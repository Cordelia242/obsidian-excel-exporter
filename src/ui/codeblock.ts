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
	ctx: MarkdownPostProcessorContext,
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
	buttons.createEl("button", { text: "Setup" }).onclick = () => void plugin.openSetup(ctx.sourcePath, blockIndex(ctx, el));
	buttons.createEl("button", { text: "Ejecutar", cls: "mod-cta" }).onclick = () => void plugin.runDefinition(compiled);
	buttons.createEl("button", { text: "Validar" }).onclick = () => void plugin.validateDefinition(compiled);
}

/** Index of this block among the note's excel-export blocks. */
function blockIndex(ctx: MarkdownPostProcessorContext, el: HTMLElement): number {
	const info = ctx.getSectionInfo(el);
	if (!info) return 0;
	const before = info.text.split("\n").slice(0, info.lineStart);
	return before.filter((l) => /^\s*(`{3,}|~{3,})\s*excel-export\b/.test(l)).length;
}
