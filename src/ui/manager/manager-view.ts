import { ItemView, Menu, Notice, setIcon, type ViewStateResult, type WorkspaceLeaf } from "obsidian";
import { worksheetToGrid } from "../../excel/grid";
import { loadWorkbook } from "../../excel/workbook-io";
import { compileTemplate, duplicateTemplate, newTemplateConfig, templateStatus, type TemplateConfig } from "../../store/templates";
import { validateTemplate } from "../../export/validate";
import type ExcelTemplateExportPlugin from "../../main";
import { button, confirmModal, iconButton } from "../components";
import { ExcelPickerModal } from "../excel-picker";
import { ValidationModal } from "../report-modal";
import { renderGrid } from "../sheet-render";
import { TemplateSession } from "./session";
import { CellsStep } from "./step-cells";
import { renderStepOutput } from "./step-output";
import { renderStepRelations } from "./step-relations";
import { renderStepWhat } from "./step-what";

export const MANAGER_VIEW_TYPE = "excel-export-manager";

export type ManagerRoute = { route: "gallery" } | { route: "editor"; id: string; step?: number };

const STEPS = [
	{ title: "Qué exporta", icon: "filter" },
	{ title: "Datos relacionados", icon: "link" },
	{ title: "Celdas del Excel", icon: "table" },
	{ title: "Archivo de salida", icon: "download" },
];

/** Template manager: gallery of templates and the step-by-step editor. */
export class ManagerView extends ItemView {
	private state: ManagerRoute = { route: "gallery" };
	private session: TemplateSession | null = null;
	private cells: CellsStep | null = null;
	private thumbs = new Map<string, { mtime: number; html: HTMLElement | null }>();

	constructor(leaf: WorkspaceLeaf, private plugin: ExcelTemplateExportPlugin) {
		super(leaf);
	}

	getViewType(): string {
		return MANAGER_VIEW_TYPE;
	}

	getDisplayText(): string {
		return this.session ? `Excel: ${this.session.cfg.name}` : "Templates de Excel";
	}

	getIcon(): string {
		return "file-spreadsheet";
	}

	getState(): Record<string, unknown> {
		return { ...this.state };
	}

	async setState(state: unknown, result: ViewStateResult): Promise<void> {
		await this.navigate((state as ManagerRoute) ?? { route: "gallery" });
		await super.setState(state, result);
	}

	async onOpen(): Promise<void> {
		this.contentEl.addClass("xte-manager");
		await this.render();
	}

	async onClose(): Promise<void> {
		await this.session?.flush();
	}

	async navigate(route: ManagerRoute): Promise<void> {
		await this.session?.flush();
		this.state = route;
		if (route.route === "editor") {
			const cfg = this.plugin.settings.templates.find((t) => t.id === route.id);
			if (!cfg) {
				this.state = { route: "gallery" };
				this.session = null;
			} else if (this.session?.cfg.id !== cfg.id) {
				this.session = new TemplateSession(this.app, this.plugin, cfg);
				this.cells = new CellsStep(this.session);
			}
		} else {
			this.session = null;
			this.cells = null;
		}
		(this.leaf as unknown as { updateHeader?: () => void }).updateHeader?.();
		await this.render();
	}

	/** Re-renders when templates change elsewhere (import, delete). */
	async refresh(): Promise<void> {
		if (this.state.route === "gallery") await this.render();
	}

	private async render(): Promise<void> {
		const el = this.contentEl;
		el.empty();
		if (this.state.route === "editor" && this.session) await this.renderEditor(el, this.session, this.state.step ?? 0);
		else await this.renderGallery(el);
	}

	// ------------------------------------------------------------------ gallery

	private async renderGallery(el: HTMLElement): Promise<void> {
		const templates = this.plugin.settings.templates;
		const head = el.createDiv({ cls: "xte-gallery-head" });
		const title = head.createDiv();
		title.createEl("h2", { text: "Templates de Excel" });
		title.createDiv({ cls: "xte-help", text: "Cada template es un Excel con formato que se llena con datos de tus notas." });
		const actions = head.createDiv({ cls: "xte-inline" });
		const legacy = await this.plugin.countLegacyDefinitions();
		if (legacy) button(actions, `Importar ${legacy} desde notas`, () => void this.plugin.importLegacyDefinitions(), { icon: "import" });
		button(actions, "Nuevo template", () => this.newTemplate(), { cta: true, icon: "plus" });

		if (!templates.length) {
			const empty = el.createDiv({ cls: "xte-empty" });
			setIcon(empty.createDiv({ cls: "xte-empty-icon" }), "file-spreadsheet");
			empty.createEl("h3", { text: "Todavía no tienes templates" });
			empty.createDiv({
				cls: "xte-help",
				text: "Sube un Excel con tus títulos y formato, elige de qué notas salen los datos y en qué celda va cada uno. Luego exportas con un clic.",
			});
			button(empty, "Crear mi primer template", () => this.newTemplate(), { cta: true, icon: "plus" });
			return;
		}

		const grid = el.createDiv({ cls: "xte-cards" });
		for (const cfg of templates) this.renderCard(grid, cfg);
	}

	private renderCard(parent: HTMLElement, cfg: TemplateConfig): void {
		const card = parent.createDiv({ cls: "xte-card" });
		const thumb = card.createDiv({ cls: "xte-card-thumb" });
		void this.fillThumb(thumb, cfg.template);
		thumb.onclick = () => void this.navigate({ route: "editor", id: cfg.id });

		const body = card.createDiv({ cls: "xte-card-body" });
		const top = body.createDiv({ cls: "xte-card-top" });
		top.createDiv({ cls: "xte-card-title", text: cfg.name });
		iconButton(top, "more-horizontal", "Más opciones", () => {
			const menu = new Menu();
			menu.addItem((i) => i.setTitle("Editar").setIcon("pencil").onClick(() => void this.navigate({ route: "editor", id: cfg.id })));
			menu.addItem((i) =>
				i.setTitle("Duplicar").setIcon("copy").onClick(async () => {
					await this.plugin.saveTemplate(duplicateTemplate(cfg));
					await this.render();
				}),
			);
			menu.addSeparator();
			menu.addItem((i) =>
				i.setTitle("Eliminar").setIcon("trash-2").onClick(() =>
					confirmModal(this.app, `¿Eliminar el template «${cfg.name}»?`, "El archivo Excel no se borra.", "Eliminar", async () => {
						await this.plugin.deleteTemplate(cfg.id);
						await this.render();
					}),
				),
			);
			const rect = top.getBoundingClientRect();
			menu.showAtPosition({ x: rect.right, y: rect.bottom });
		});

		const label = cfg.root.label || "Nota";
		const what =
			cfg.mode === "file-per-root"
				? `Un archivo por cada ${label}`
				: cfg.mode === "sheet-per-root"
					? `Una hoja por cada ${label}`
					: `Un archivo con todos los ${label}s`;
		body.createDiv({ cls: "xte-card-desc", text: cfg.description || what });
		const meta = body.createDiv({ cls: "xte-card-meta" });
		const file = meta.createSpan();
		setIcon(file.createSpan({ cls: "xte-btn-icon" }), "sheet");
		file.appendText(` ${(cfg.template.split("/").pop() ?? "").replace(/\.xlsx$/i, "") || "sin Excel"}`);

		const status = templateStatus(cfg);
		const foot = body.createDiv({ cls: "xte-card-foot" });
		if (status.ready) {
			const { cd } = compileTemplate(cfg);
			if (cd) {
				const n = this.plugin.countAvailable(cd);
				meta.createSpan({ text: ` · ${n} ${label}${n === 1 ? "" : "s"} disponibles` });
			}
			button(foot, "Exportar", () => this.plugin.openExportDialog(cfg), { cta: true, icon: "download" });
			button(foot, "Editar", () => void this.navigate({ route: "editor", id: cfg.id }), { icon: "pencil" });
		} else {
			foot.createDiv({ cls: "xte-card-warn", text: status.message });
			button(foot, "Continuar configurando", () => void this.navigate({ route: "editor", id: cfg.id }), { cta: true, icon: "pencil" });
		}
	}

	private async fillThumb(el: HTMLElement, path: string): Promise<void> {
		const file = this.app.vault.getAbstractFileByPath(path) as { stat?: { mtime: number } } | null;
		const mtime = file?.stat?.mtime ?? 0;
		const cached = this.thumbs.get(path);
		if (cached && cached.mtime === mtime && cached.html) {
			el.appendChild(cached.html.cloneNode(true));
			return;
		}
		try {
			const file = this.app.vault.getFileByPath(path);
			if (!file) throw new Error();
			const wb = await loadWorkbook(await this.app.vault.readBinary(file));
			const ws = wb.worksheets[0];
			const holder = createDiv({ cls: "xte-thumb-inner" });
			renderGrid(holder, worksheetToGrid(ws, { maxRows: 14, maxCols: 8, minRows: 10, minCols: 6 }));
			this.thumbs.set(path, { mtime, html: holder });
			el.appendChild(holder.cloneNode(true));
		} catch {
			setIcon(el.createDiv({ cls: "xte-thumb-missing" }), "file-question");
		}
	}

	newTemplate(): void {
		new ExcelPickerModal(
			this.app,
			this.plugin.settings.templatesFolder,
			async (path, name) => {
				const cfg = newTemplateConfig(name || "Nuevo template", path, this.plugin.settings.outputFolder);
				await this.plugin.saveTemplate(cfg);
				await this.navigate({ route: "editor", id: cfg.id, step: 0 });
			},
			"Nuevo template",
			true,
		).open();
	}

	// ------------------------------------------------------------------ editor

	private async renderEditor(el: HTMLElement, s: TemplateSession, step: number): Promise<void> {
		const cfg = s.cfg;
		const header = el.createDiv({ cls: "xte-editor-head" });
		iconButton(header, "arrow-left", "Volver a templates", () => void this.navigate({ route: "gallery" }), "xte-back");
		const nameInput = header.createEl("input", { type: "text", cls: "xte-title-input", value: cfg.name });
		nameInput.title = "Nombre del template";
		nameInput.onchange = () => {
			cfg.name = nameInput.value.trim() || cfg.name;
			s.changed();
			(this.leaf as unknown as { updateHeader?: () => void }).updateHeader?.();
		};
		const excel = header.createEl("button", { cls: "xte-excel-chip" });
		setIcon(excel.createSpan({ cls: "xte-btn-icon" }), "sheet");
		excel.createSpan({ text: (cfg.template.split("/").pop() ?? "Elegir Excel").replace(/\.xlsx$/i, "") });
		excel.title = `Excel: ${cfg.template}. Clic para cambiarlo.`;
		excel.onclick = () =>
			new ExcelPickerModal(this.app, this.plugin.settings.templatesFolder, (path) => {
				cfg.template = path;
				s.changed();
				void this.render();
			}, "Cambiar el archivo Excel").open();

		const right = header.createDiv({ cls: "xte-editor-actions" });
		const saved = right.createSpan({ cls: "xte-saved" });
		s.onSaved = (ok) => saved.setText(ok ? "Guardado" : "Guardando…");
		saved.setText("Guardado");
		button(right, "Revisar", () => void this.check(s), { icon: "check-check" });
		button(right, "Exportar…", async () => {
			await s.flush();
			this.plugin.openExportDialog(cfg);
		}, { cta: true, icon: "download" });

		// Step bar
		const status = this.stepStatus(s);
		const bar = el.createDiv({ cls: "xte-steps" });
		STEPS.forEach((st, i) => {
			const b = bar.createEl("button", { cls: `xte-step${i === step ? " is-active" : ""}${status[i] === "ok" ? " is-done" : ""}` });
			const num = b.createSpan({ cls: "xte-step-num" });
			if (status[i] === "ok" && i !== step) setIcon(num, "check");
			else num.setText(String(i + 1));
			b.createSpan({ text: st.title });
			if (status[i] === "todo" && i !== step) b.createSpan({ cls: "xte-step-dot", attr: { title: "Pendiente" } });
			b.onclick = () => void this.navigate({ route: "editor", id: cfg.id, step: i });
		});

		const content = el.createDiv({ cls: `xte-step-content${step === 2 ? " is-wide" : ""}` });
		const rerender = () => void this.navigate({ route: "editor", id: cfg.id, step });
		if (step === 0) renderStepWhat(content, s, rerender);
		else if (step === 1) renderStepRelations(content, s, rerender);
		else if (step === 2) await this.cells?.render(content);
		else await renderStepOutput(content, s, rerender);

		if (step !== 2) {
			const nav = content.createDiv({ cls: "xte-step-nav" });
			if (step > 0) button(nav, "Anterior", () => void this.navigate({ route: "editor", id: cfg.id, step: step - 1 }), { icon: "arrow-left" });
			nav.createDiv({ cls: "xte-spacer" });
			if (step < STEPS.length - 1) {
				button(nav, `Siguiente: ${STEPS[step + 1].title}`, () => void this.navigate({ route: "editor", id: cfg.id, step: step + 1 }), { cta: true });
			} else {
				button(nav, "Exportar…", async () => {
					await s.flush();
					this.plugin.openExportDialog(cfg);
				}, { cta: true, icon: "download" });
			}
		}
	}

	private stepStatus(s: TemplateSession): Array<"ok" | "todo" | "optional"> {
		const cfg = s.cfg;
		return [
			cfg.root.where.trim() && s.candidates().length ? "ok" : "todo",
			Object.keys(cfg.relations).length ? "ok" : "optional",
			Object.keys(cfg.cells).length ? "ok" : "todo",
			"optional",
		];
	}

	private async check(s: TemplateSession): Promise<void> {
		await s.flush();
		const { cd, errors } = compileTemplate(s.cfg);
		if (!cd) {
			new Notice(`Falta completar: ${errors[0]}`, 8000);
			return;
		}
		try {
			const result = await validateTemplate(cd, s.adapter, { settings: this.plugin.settings });
			new ValidationModal(this.app, s.cfg.name, result).open();
		} catch (e) {
			new Notice((e as Error).message);
		}
	}
}
