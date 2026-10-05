import { Modal, setIcon } from "obsidian";
import { ObsidianVaultAdapter } from "../adapter/obsidian-adapter";
import { Runtime } from "../graph/context";
import { resolveOnValue } from "../graph/path";
import type { NoteRecord } from "../model/note-record";
import { Report } from "../model/report";
import { evaluate } from "../query/evaluator";
import { sortNotes } from "../query/sort";
import { exprToConditions } from "../setup/conditions";
import { compileTemplate, type TemplateConfig } from "../store/templates";
import { DEFAULT_NORMALIZE, normalizeValue, scalarToText } from "../values/normalize";
import type ExcelTemplateExportPlugin from "../main";
import { button, callout, checklist, choiceCards } from "./components";

/**
 * "Export" dialog of a template: the user picks which notes (e.g. which
 * Procesos) to export from a searchable list, or all the ones that pass the filter.
 */
export class ExportDialog extends Modal {
	private mode: "pick" | "all";
	private selected: string[];

	constructor(
		private plugin: ExcelTemplateExportPlugin,
		private cfg: TemplateConfig,
		preselect: string[] = [],
	) {
		super(plugin.app);
		this.selected = preselect;
		this.mode = "pick";
	}

	onOpen(): void {
		this.modalEl.addClass("xte-modal-mid");
		this.titleEl.setText(`Exportar «${this.cfg.name}»`);
		const { contentEl } = this;
		const { cd, errors } = compileTemplate(this.cfg);
		if (!cd) {
			callout(contentEl, `El template no está terminado: ${errors[0]}`, "warn");
			const f = contentEl.createDiv({ cls: "modal-button-container" });
			button(f, "Editar template", () => {
				this.close();
				void this.plugin.openManager({ route: "editor", id: this.cfg.id });
			}, { cta: true });
			return;
		}
		const adapter = new ObsidianVaultAdapter(this.app);
		const rt = new Runtime(adapter, new Report());
		const label = this.cfg.root.label || "Nota";
		const candidates = sortNotes(adapter.listNotes().filter((n) => evaluate(cd.where, n, rt)), cd.sort, rt);
		const filtered = cd.filter ? candidates.filter((n) => evaluate(cd.filter!, n, rt)) : candidates;

		// Show the properties used by the filter/sort next to each option (e.g. "status: Open").
		const detailKeys = [
			...(exprToConditions(this.cfg.root.filter) ?? []).map((c) => c.field),
			...(this.cfg.root.sort ? [this.cfg.root.sort.replace(/\s+(asc|desc)\s*$/i, "").replace(/^\["|"\]$/g, "")] : []),
		].filter((k, i, a) => k && !k.startsWith("file.") && a.indexOf(k) === i).slice(0, 2);
		const detail = (n: NoteRecord) =>
			detailKeys
				.map((k) => {
					const v = scalarToText(normalizeValue(resolveOnValue(n, [k], rt), DEFAULT_NORMALIZE));
					return v ? `${k}: ${v}` : "";
				})
				.filter(Boolean)
				.join(" · ") || (n.path.includes("/") ? n.path.slice(0, n.path.lastIndexOf("/")) : "");

		const plural = `${label}s`;
		contentEl.createDiv({ cls: "xte-q", text: `¿Qué ${plural} quieres exportar?` });
		choiceCards(
			contentEl,
			[
				{ value: "pick", icon: "mouse-pointer-click", title: `Elegir ${plural}`, description: "Busca por nombre y elige uno o varios." },
				{
					value: "all",
					icon: "list-checks",
					title: `Todos${this.cfg.root.filter ? " los que pasan el filtro" : ""} (${filtered.length})`,
					description: filtered.slice(0, 4).map((n) => n.basename).join(", ") + (filtered.length > 4 ? "…" : ""),
				},
			],
			this.mode,
			(v) => {
				this.mode = v as "pick" | "all";
				pickerBox.toggle(this.mode === "pick");
				updateSummary();
			},
		);

		const pickerBox = contentEl.createDiv({ cls: "xte-picker-box" });
		checklist(pickerBox, {
			options: candidates.map((n) => ({ value: n.path, label: n.basename, detail: detail(n) })),
			selected: this.selected,
			placeholder: `Buscar ${label.toLowerCase()} por nombre…`,
			onChange: (sel) => {
				this.selected = sel;
				updateSummary();
			},
		});
		pickerBox.toggle(this.mode === "pick");

		const summary = contentEl.createDiv({ cls: "xte-summary" });
		const footer = contentEl.createDiv({ cls: "modal-button-container" });
		button(footer, "Cancelar", () => this.close());
		const go = button(footer, this.plugin.settings.previewBeforeExport ? "Ver vista previa" : "Exportar", () => {
			const paths = this.mode === "all" ? filtered.map((n) => n.path) : this.selected;
			this.close();
			void this.plugin.runTemplate(this.cfg, paths);
		}, { cta: true, icon: this.plugin.settings.previewBeforeExport ? "eye" : "download" });

		const updateSummary = () => {
			summary.empty();
			const n = this.mode === "all" ? filtered.length : this.selected.length;
			go.disabled = n === 0;
			if (!n) {
				summary.createSpan({ cls: "xte-help", text: this.mode === "pick" ? `Elige al menos un ${label.toLowerCase()}.` : `Ningún ${label.toLowerCase()} pasa el filtro.` });
				return;
			}
			setIcon(summary.createSpan({ cls: "xte-btn-icon" }), "file-spreadsheet");
			const files = this.cfg.mode === "file-per-root" ? n : 1;
			const folder = this.cfg.output.folder ?? this.plugin.settings.outputFolder;
			const extra = this.cfg.mode === "sheet-per-root" ? ` con ${n} ${n === 1 ? "hoja" : "hojas"}` : "";
			summary.createSpan({ text: ` Se ${files === 1 ? "generará 1 archivo" : `generarán ${files} archivos`}${extra} en «${folder}».` });
		};
		updateSummary();
	}

	onClose(): void {
		this.contentEl.empty();
	}
}
