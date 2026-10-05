import { Modal, type App } from "obsidian";
import { worksheetToGrid } from "../excel/grid";
import type { BuiltExport } from "../export/runner";
import { ReportModal } from "./report-modal";
import { renderGrid, renderTabs } from "./sheet-render";

/**
 * Shows the generated workbooks inside Obsidian before writing them.
 * `onConfirm` is omitted for a read-only preview (Setup).
 */
export class PreviewModal extends Modal {
	private fileIdx = 0;
	private sheetIdx = 0;
	private decided = false;

	constructor(
		app: App,
		private title: string,
		private built: BuiltExport,
		private onConfirm?: () => void,
		private onCancel?: () => void,
	) {
		super(app);
	}

	onOpen(): void {
		this.modalEl.addClass("xte-modal-wide");
		this.titleEl.setText(this.title);
		this.render();
	}

	private render(): void {
		const { contentEl, built } = this;
		contentEl.empty();
		const n = built.outputs.length;
		const w = built.warnings.length;

		const info = contentEl.createDiv({ cls: "xte-preview-info" });
		info.createSpan({
			text: `${n} ${n === 1 ? "archivo" : "archivos"} · ${built.run.roots.length} ${built.run.roots.length === 1 ? "raíz" : "raíces"}`,
		});
		if (w) {
			info.createSpan({ text: " · " });
			const link = info.createEl("a", { text: `${w} ${w === 1 ? "warning" : "warnings"}`, href: "#" });
			link.onclick = (ev) => {
				ev.preventDefault();
				new ReportModal(this.app, `Warnings — ${this.title}`, built.warnings).open();
			};
		}

		if (!n) {
			contentEl.createEl("p", { text: "No se generó ningún archivo: ninguna nota cumple el filtro de la raíz." });
		} else {
			if (n > 1) {
				const select = contentEl.createEl("select", { cls: "dropdown xte-file-select" });
				built.outputs.forEach((o, i) => {
					const opt = select.createEl("option", { text: o.fileName, value: String(i) });
					if (i === this.fileIdx) opt.selected = true;
				});
				select.onchange = () => {
					this.fileIdx = Number(select.value);
					this.sheetIdx = 0;
					this.render();
				};
			} else {
				contentEl.createDiv({ text: built.outputs[0].fileName, cls: "xte-muted" });
			}
			const wb = built.outputs[this.fileIdx].workbook;
			const sheets = wb.worksheets;
			renderTabs(contentEl, sheets.map((s) => s.name), this.sheetIdx, (i) => {
				this.sheetIdx = i;
				this.render();
			});
			const ws = sheets[this.sheetIdx];
			if (ws) renderGrid(contentEl.createDiv({ cls: "xte-preview-grid" }), worksheetToGrid(ws, { minRows: 10, minCols: 6 }));
			contentEl.createDiv({
				cls: "xte-muted",
				text: "Las fórmulas se muestran como =FÓRMULA; Excel las calcula al abrir el archivo.",
			});
		}

		const buttons = contentEl.createDiv({ cls: "modal-button-container" });
		if (this.onConfirm) {
			const ok = buttons.createEl("button", { text: n === 1 ? "Exportar" : `Exportar ${n} archivos`, cls: "mod-cta" });
			ok.disabled = n === 0;
			ok.onclick = () => {
				this.decided = true;
				this.close();
				this.onConfirm?.();
			};
			buttons.createEl("button", { text: "Cancelar" }).onclick = () => this.close();
		} else {
			buttons.createEl("button", { text: "Cerrar" }).onclick = () => this.close();
		}
	}

	onClose(): void {
		this.contentEl.empty();
		if (!this.decided) this.onCancel?.();
	}
}
