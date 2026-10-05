import type { CompiledMapping } from "../config/mapping";
import { targetsSheet } from "../config/mapping";
import type { Report } from "../model/report";
import type { EachSpec } from "./scanner";
import type { Workbook, Worksheet } from "./workbook-io";

/** Writes the Setup cell mapping into the (in-memory) template as placeholder text. */
export function applyCellMapping(wb: Workbook, mapping: CompiledMapping, report: Report): void {
	const sheets = wb.worksheets;
	if (!sheets.length) return;
	const first = sheets[0].name;
	for (const m of mapping.cells) {
		const ws = sheets.find((s) => targetsSheet(m.sheet, s.name, first));
		if (!ws) {
			report.warn("other", `La hoja "${m.sheet}" del mapeo no existe en el template`, { location: `${m.sheet}!${m.address}` });
			continue;
		}
		let cell = ws.getCell(m.row, m.col);
		if (cell.isMerged && cell.master !== cell) cell = cell.master;
		cell.value = m.expr;
	}
}

/** Row repeats (Setup `rows`) that apply to a sheet of the template. */
export function rowRepeatsFor(mapping: CompiledMapping, ws: Worksheet, firstSheet: string): Map<number, EachSpec> {
	const out = new Map<number, EachSpec>();
	for (const r of mapping.rows) {
		if (targetsSheet(r.sheet, ws.name, firstSheet)) out.set(r.row, { src: r.src, path: r.path });
	}
	return out;
}
