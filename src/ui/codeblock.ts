import type { MarkdownPostProcessorContext } from "obsidian";
import { countRoots } from "../export/runner";
import type ExcelTemplateExportPlugin from "../main";
import { ObsidianVaultAdapter } from "../adapter/obsidian-adapter";
import { compileSource, errorText } from "./definition-picker";

/** Renders an `excel-export` block: summary, matching roots and Run/Validate buttons. */
export function renderCodeBlock(
	plugin: ExcelTemplateExportPlugin,
	source: string,
	el: HTMLElement,
	_ctx: MarkdownPostProcessorContext,
): void {
	const box = el.createDiv({ cls: "xte-block" });
	const { compiled, error } = compileSource(source);
	if (!compiled) {
		box.addClass("xte-block-error");
		box.createEl("strong", { text: "Excel Export: definición inválida" });
		box.createEl("pre", { text: error ?? "" });
		return;
	}
	const def = compiled.def;
	box.createEl("div", { text: def.name, cls: "xte-block-title" });
	const meta = box.createEl("div", { cls: "xte-muted" });
	meta.createSpan({ text: `Template: ${def.template} · Modo: ${def.mode} · ` });
	const count = meta.createSpan({ text: "contando raíces…" });
	window.setTimeout(() => {
		try {
			const n = countRoots(compiled, new ObsidianVaultAdapter(plugin.app));
			count.setText(`${n} ${n === 1 ? "raíz" : "raíces"}`);
		} catch (e) {
			count.setText(`error: ${errorText(e)}`);
		}
	}, 0);
	const rels = Object.keys(def.relations);
	if (rels.length) box.createEl("div", { text: `Relaciones: ${rels.join(", ")}`, cls: "xte-muted" });
	const nCells = compiled.mapping.cells.length;
	const nRows = compiled.mapping.rows.length;
	if (nCells || nRows) {
		box.createEl("div", {
			text: `Setup: ${nCells} ${nCells === 1 ? "celda" : "celdas"}, ${nRows} ${nRows === 1 ? "fila repetida" : "filas repetidas"}`,
			cls: "xte-muted",
		});
	}
	const buttons = box.createDiv({ cls: "xte-block-buttons" });
	buttons.createEl("button", { text: "Importar al gestor de templates", cls: "mod-cta" }).onclick = () => void plugin.importLegacyDefinitions();
	buttons.createEl("button", { text: "Exportar" }).onclick = () => void plugin.runDefinition(compiled);
	box.createDiv({ cls: "xte-muted", text: "Los templates ahora se gestionan desde «Templates de Excel» (icono de la barra lateral)." });
}

