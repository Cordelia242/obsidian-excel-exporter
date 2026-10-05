import { FuzzySuggestModal, Notice, parseYaml, Plugin, TFile, type App, type TAbstractFile } from "obsidian";
import { ObsidianVaultAdapter } from "./adapter/obsidian-adapter";
import { CODEBLOCK_LANG, extractCodeBlocks } from "./config/parse";
import type { CompiledDefinition } from "./config/schema";
import { buildExport, writeExport, type RunOptions } from "./export/runner";
import { Runtime } from "./graph/context";
import { Report } from "./model/report";
import { evaluate } from "./query/evaluator";
import { DEFAULT_SETTINGS, ExcelExportSettingTab, type ExportSettings } from "./settings";
import { compileTemplate, fromRawDefinition, templateStatus, type TemplateConfig } from "./store/templates";
import { renderCodeBlock } from "./ui/codeblock";
import { errorText, findDefinitions } from "./ui/definition-picker";
import { ExportDialog } from "./ui/export-dialog";
import { MANAGER_VIEW_TYPE, ManagerView, type ManagerRoute } from "./ui/manager/manager-view";
import { PreviewModal } from "./ui/preview-modal";
import { askOverwrite, openFile, showResultNotice } from "./ui/report-modal";

class TemplatePicker extends FuzzySuggestModal<TemplateConfig> {
	constructor(app: App, private items: TemplateConfig[], private onPick: (t: TemplateConfig) => void, placeholder: string) {
		super(app);
		this.setPlaceholder(placeholder);
	}
	getItems(): TemplateConfig[] {
		return this.items;
	}
	getItemText(t: TemplateConfig): string {
		return t.name;
	}
	onChooseItem(t: TemplateConfig): void {
		this.onPick(t);
	}
}

export default class ExcelTemplateExportPlugin extends Plugin {
	settings: ExportSettings = { ...DEFAULT_SETTINGS };

	async onload(): Promise<void> {
		await this.loadSettings();
		this.addSettingTab(new ExcelExportSettingTab(this.app, this));
		this.registerView(MANAGER_VIEW_TYPE, (leaf) => new ManagerView(leaf, this));

		this.addRibbonIcon("file-spreadsheet", "Templates de Excel", () => void this.openManager({ route: "gallery" }));

		this.addCommand({
			id: "open-manager",
			name: "Abrir templates de Excel",
			callback: () => void this.openManager({ route: "gallery" }),
		});
		this.addCommand({
			id: "export",
			name: "Exportar a Excel…",
			callback: () => {
				const file = this.app.workspace.getActiveFile();
				if (file && file.extension === "md") void this.exportFromNote(file);
				else this.pickTemplate(this.readyTemplates(), (t) => this.openExportDialog(t));
			},
		});
		this.addCommand({
			id: "new-template",
			name: "Nuevo template de Excel",
			callback: async () => {
				await this.openManager({ route: "gallery" });
				this.findManager()?.newTemplate();
			},
		});
		this.addCommand({
			id: "import-definitions",
			name: "Importar definiciones desde notas",
			callback: () => void this.importLegacyDefinitions(),
		});

		this.registerMarkdownCodeBlockProcessor(CODEBLOCK_LANG, (source, el, ctx) => renderCodeBlock(this, source, el, ctx));

		this.registerEvent(
			this.app.workspace.on("file-menu", (menu, file: TAbstractFile) => {
				if (!(file instanceof TFile) || file.extension !== "md") return;
				menu.addItem((item) =>
					item
						.setTitle("Exportar a Excel…")
						.setIcon("file-spreadsheet")
						.onClick(() => void this.exportFromNote(file)),
				);
			}),
		);
	}

	async loadSettings(): Promise<void> {
		const data = ((await this.loadData()) ?? {}) as Partial<ExportSettings>;
		this.settings = { ...DEFAULT_SETTINGS, ...data, templates: data.templates ?? [] };
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.settings);
	}

	// ------------------------------------------------------------------ templates store

	async saveTemplate(cfg: TemplateConfig): Promise<void> {
		const list = this.settings.templates;
		const i = list.findIndex((t) => t.id === cfg.id);
		if (i >= 0) list[i] = cfg;
		else list.push(cfg);
		await this.saveSettings();
	}

	async deleteTemplate(id: string): Promise<void> {
		this.settings.templates = this.settings.templates.filter((t) => t.id !== id);
		await this.saveSettings();
	}

	readyTemplates(): TemplateConfig[] {
		return this.settings.templates.filter((t) => templateStatus(t).ready);
	}

	countAvailable(cd: CompiledDefinition): number {
		const adapter = new ObsidianVaultAdapter(this.app);
		const rt = new Runtime(adapter, new Report());
		let roots = adapter.listNotes().filter((n) => evaluate(cd.where, n, rt));
		if (cd.filter) roots = roots.filter((n) => evaluate(cd.filter!, n, rt));
		return roots.length;
	}

	// ------------------------------------------------------------------ navigation

	findManager(): ManagerView | null {
		const leaf = this.app.workspace.getLeavesOfType(MANAGER_VIEW_TYPE)[0];
		return (leaf?.view as ManagerView | undefined) ?? null;
	}

	async openManager(route: ManagerRoute): Promise<void> {
		const existing = this.app.workspace.getLeavesOfType(MANAGER_VIEW_TYPE)[0];
		const leaf = existing ?? this.app.workspace.getLeaf("tab");
		if (existing) await (existing.view as ManagerView).navigate(route);
		else await leaf.setViewState({ type: MANAGER_VIEW_TYPE, active: true, state: route });
		this.app.workspace.revealLeaf(leaf);
	}

	private pickTemplate(items: TemplateConfig[], onPick: (t: TemplateConfig) => void, placeholder = "¿Con qué template exportas?"): void {
		if (!items.length) {
			new Notice("Todavía no hay templates listos. Se abre el gestor para crear uno.");
			void this.openManager({ route: "gallery" });
			return;
		}
		if (items.length === 1) return onPick(items[0]);
		new TemplatePicker(this.app, items, onPick, placeholder).open();
	}

	openExportDialog(cfg: TemplateConfig, preselect: string[] = []): void {
		new ExportDialog(this, cfg, preselect).open();
	}

	/** Command / file menu on a note: templates whose notes include this one, with the note preselected. */
	async exportFromNote(file: TFile): Promise<void> {
		const adapter = new ObsidianVaultAdapter(this.app);
		const note = adapter.listNotes().find((n) => n.path === file.path);
		const rt = new Runtime(adapter, new Report());
		const matching = this.readyTemplates().filter((t) => {
			const { cd } = compileTemplate(t);
			return cd && note && evaluate(cd.where, note, rt);
		});
		if (matching.length) {
			this.pickTemplate(matching, (t) => this.openExportDialog(t, [file.path]), `Exportar «${file.basename}» con…`);
		} else {
			new Notice(`Ningún template exporta notas como «${file.basename}». Elige uno.`);
			this.pickTemplate(this.readyTemplates(), (t) => this.openExportDialog(t));
		}
	}

	// ------------------------------------------------------------------ export

	async runTemplate(cfg: TemplateConfig, rootPaths: string[]): Promise<void> {
		const { cd, errors } = compileTemplate(cfg);
		if (!cd) {
			new Notice(`El template no está terminado: ${errors[0]}`, 8000);
			return;
		}
		await this.runDefinition(cd, rootPaths);
	}

	async runDefinition(cd: CompiledDefinition, rootPaths?: string[]): Promise<void> {
		const adapter = new ObsidianVaultAdapter(this.app);
		const options: RunOptions = {
			settings: this.settings,
			rootPaths,
			confirmOverwrite: (path) => askOverwrite(this.app, path),
		};
		const progress = new Notice(`Generando «${cd.def.name}»…`, 0);
		try {
			const built = await buildExport(cd, adapter, options);
			progress.hide();
			const write = async () => {
				try {
					const result = await writeExport(built, adapter, options);
					showResultNotice(this.app, cd.def.name, result.files, result.warnings);
					if (this.settings.openAfterExport && result.files.length === 1) openFile(this.app, result.files[0]);
				} catch (e) {
					console.error("Excel Template Export", e);
					new Notice(`Error al guardar «${cd.def.name}»: ${errorText(e)}`, 10000);
				}
			};
			if (this.settings.previewBeforeExport) new PreviewModal(this.app, `Vista previa — ${cd.def.name}`, built, () => void write()).open();
			else await write();
		} catch (e) {
			progress.hide();
			console.error("Excel Template Export", e);
			new Notice(`Error en «${cd.def.name}»: ${errorText(e)}`, 10000);
		}
	}

	// ------------------------------------------------------------------ legacy (YAML in notes)

	async countLegacyDefinitions(): Promise<number> {
		const names = new Set(this.settings.templates.map((t) => t.name));
		const defs = await findDefinitions(this.app, "");
		return defs.filter((d) => d.compiled && !names.has(d.compiled.def.name)).length;
	}

	async importLegacyDefinitions(): Promise<void> {
		const names = new Set(this.settings.templates.map((t) => t.name));
		let n = 0;
		for (const d of await findDefinitions(this.app, "")) {
			if (!d.compiled || names.has(d.compiled.def.name)) continue;
			const body = extractCodeBlocks(await this.app.vault.cachedRead(d.file))[d.index] ?? "";
			try {
				this.settings.templates.push(fromRawDefinition(parseYaml(body) as Record<string, unknown>));
				names.add(d.compiled.def.name);
				n++;
			} catch (e) {
				new Notice(`No se pudo importar ${d.file.path}: ${errorText(e)}`);
			}
		}
		await this.saveSettings();
		new Notice(n ? `${n} ${n === 1 ? "template importado" : "templates importados"}.` : "No hay definiciones nuevas para importar.");
		await this.findManager()?.refresh();
		if (n) await this.openManager({ route: "gallery" });
	}
}
