import type ExcelJS from "exceljs";
import type { EmptyBlockMode } from "../config/schema";
import type { Runtime, Scope } from "../graph/context";
import { isDateOnly, scalarToText, type NormalizeOptions } from "../values/normalize";
import type { Rendered } from "../values/pipes";
import { copyMapper, insertionMapper, mapFormulaRefs, type RefMapper } from "./formulas";
import { iterateCollection, renderTemplate, type CellTemplate, type RenderContext } from "./placeholders";
import { scanWorksheet } from "./scanner";
import { isFormulaValue, toExcelDate, type Workbook, type Worksheet } from "./workbook-io";

export interface FillOptions {
	rt: Runtime;
	opts: NormalizeOptions;
	emptyBlock: EmptyBlockMode;
}

interface MergeRect {
	top: number;
	left: number;
	bottom: number;
	right: number;
}

function clone<T>(v: T): T {
	return v === undefined || v === null || typeof v !== "object" ? v : structuredClone(v);
}

function listMerges(ws: Worksheet): MergeRect[] {
	const merges = (ws as unknown as { _merges: Record<string, MergeRect> })._merges ?? {};
	return Object.values(merges).map(({ top, left, bottom, right }) => ({ top, left, bottom, right }));
}

function merge(ws: Worksheet, m: MergeRect): void {
	const w = ws as unknown as { mergeCellsWithoutStyle?: (...a: number[]) => void };
	if (w.mergeCellsWithoutStyle) w.mergeCellsWithoutStyle(m.top, m.left, m.bottom, m.right);
	else ws.mergeCells(m.top, m.left, m.bottom, m.right);
}

function colName(col: number): string {
	let s = "";
	for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
	return s;
}

/** Writes a rendered value, keeping the cell's style; dates get a date format if the cell has none. */
export function writeRendered(cell: ExcelJS.Cell, rendered: Rendered): void {
	let value = rendered.value;
	if (value instanceof Date) {
		const fmt = cell.numFmt;
		if (!fmt || fmt === "General") {
			cell.style = { ...cell.style, numFmt: isDateOnly(value) ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm" };
		}
		value = toExcelDate(value);
	}
	if (rendered.hyperlink) {
		cell.value = { text: scalarToText(rendered.value), hyperlink: rendered.hyperlink };
	} else {
		cell.value = value;
	}
}

/** Applies a reference mapper to every formula in the workbook that points to `sheetName`. */
function adjustFormulas(
	wb: Workbook,
	sheetName: string,
	mapper: RefMapper,
	rt: Runtime,
	skip?: { sheet: Worksheet; from: number; to: number },
): void {
	wb.eachSheet((ws) => {
		ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
			if (skip && ws === skip.sheet && rowNumber >= skip.from && rowNumber <= skip.to) return;
			row.eachCell({ includeEmpty: false }, (cell) => {
				const v = cell.value;
				if (!isFormulaValue(v) || !("formula" in v) || !v.formula) return;
				const next = mapFormulaRefs(v.formula, ws.name, sheetName, mapper);
				if (next !== v.formula) {
					if (next.includes("#REF!") && !v.formula.includes("#REF!")) {
						rt.report.warn(
							"formula",
							`La fórmula =${v.formula} quedó como =${next} porque la fila #each no tuvo elementos (usá emptyBlock: blank para conservarla)`,
							{ location: `${ws.name}!${cell.address}` },
						);
					}
					cell.value = { ...v, formula: next, result: undefined } as ExcelJS.CellFormulaValue;
					wb.calcProperties = { ...wb.calcProperties, fullCalcOnLoad: true };
				}
			});
		});
	});
}

interface Anchorish {
	nativeRow?: number;
}

function shiftImages(ws: Worksheet, at: number, delta: number): void {
	const media = (ws as unknown as { _media?: Array<{ range?: { tl?: Anchorish; br?: Anchorish } }> })._media ?? [];
	for (const m of media) {
		const tl = m.range?.tl;
		// nativeRow is 0-based: rows strictly below the template row have nativeRow >= at.
		if (tl && typeof tl.nativeRow === "number" && tl.nativeRow >= at) {
			tl.nativeRow += delta;
			const br = m.range?.br;
			if (br && typeof br.nativeRow === "number") br.nativeRow += delta;
		}
	}
}

/** Fills one worksheet: scalar placeholders first, then `#each` rows from bottom to top. */
export function fillWorksheet(wb: Workbook, ws: Worksheet, scope: Scope, fo: FillOptions): void {
	const report = fo.rt.report;
	const scan = scanWorksheet(ws);
	for (const issue of scan.issues) report.warn("other", issue.message, { location: `${ws.name}!${issue.address}` });
	const tables = (ws as unknown as { tables?: Record<string, unknown> }).tables;
	if (tables && Object.keys(tables).length && scan.eachRows.length) {
		report.warn("table", `La hoja "${ws.name}" tiene Tablas de Excel; no se redimensionan al insertar filas`);
	}

	const eachRowNumbers = new Set(scan.eachRows.map((e) => e.row));
	for (const c of scan.cells) {
		if (eachRowNumbers.has(c.row)) continue;
		const ctx: RenderContext = { rt: fo.rt, opts: fo.opts, location: `${ws.name}!${c.address}` };
		writeRendered(ws.getCell(c.row, c.col), renderTemplate(c.tpl, scope, ctx));
	}

	for (const e of [...scan.eachRows].sort((a, b) => b.row - a.row)) {
		const each = e.tpl.each;
		if (!each) continue;
		const tplByCol = new Map<number, CellTemplate>();
		for (const c of scan.cells) if (c.row === e.row) tplByCol.set(c.col, c.tpl);
		let scopes: Scope[] = [];
		if (each.error) {
			report.warn("unresolved", `{{#each ${each.src}}}: ${each.error}`, { location: `${ws.name}!${e.address}` });
		} else {
			const items = iterateCollection(each.path, scope, fo.rt);
			if (items === null) {
				report.warn("unresolved", `{{#each ${each.src}}}: colección desconocida "${each.path[0]}"`, {
					location: `${ws.name}!${e.address}`,
				});
			} else {
				scopes = items;
			}
		}
		expandRow(wb, ws, e.row, tplByCol, scopes, fo);
	}
}

function expandRow(
	wb: Workbook,
	ws: Worksheet,
	r: number,
	tplByCol: Map<number, CellTemplate>,
	scopes: Scope[],
	fo: FillOptions,
): void {
	const report = fo.rt.report;
	const n = scopes.length;

	// Unmerge everything at or below the template row; merges are re-created after moving rows.
	const affected = listMerges(ws).filter((m) => m.bottom >= r);
	for (const m of affected) ws.unMergeCells(m.top, m.left, m.bottom, m.right);

	const row = ws.getRow(r);
	const height = row.height;
	const tplCells: Array<{ col: number; value: ExcelJS.CellValue; style: Partial<ExcelJS.Style> }> = [];
	row.eachCell({ includeEmpty: true }, (cell, col) => {
		tplCells.push({ col, value: clone(cell.value), style: clone(cell.style) });
	});

	if (n === 0 && fo.emptyBlock === "blank") {
		for (const t of tplCells) if (tplByCol.has(t.col)) ws.getCell(r, t.col).value = null;
		for (const m of affected) merge(ws, m);
		return;
	}

	const delta = n === 0 ? -1 : n - 1;
	if (delta < 0) ws.spliceRows(r, 1);
	else if (delta > 0) ws.spliceRows(r + 1, 0, ...Array.from({ length: delta }, () => []));

	if (delta !== 0) {
		adjustFormulas(wb, ws.name, insertionMapper(r, delta), fo.rt, n > 0 ? { sheet: ws, from: r, to: r + n - 1 } : undefined);
		shiftImages(ws, r, delta);
		const cf = (ws as unknown as { conditionalFormattings?: unknown[] }).conditionalFormattings;
		const dv = (ws as unknown as { dataValidations?: { model?: Record<string, unknown> } }).dataValidations?.model;
		if ((cf && cf.length) || (dv && Object.keys(dv).length)) {
			report.warn(
				"other",
				`La hoja "${ws.name}" tiene formato condicional o validaciones de datos; sus rangos no se ajustan al insertar filas`,
			);
		}
	}

	for (let i = 0; i < n; i++) {
		const target = ws.getRow(r + i);
		if (height) target.height = height;
		for (const t of tplCells) {
			const cell = target.getCell(t.col);
			cell.style = clone(t.style);
			const tpl = tplByCol.get(t.col);
			if (tpl) {
				const ctx: RenderContext = { rt: fo.rt, opts: fo.opts, location: `${ws.name}!${colName(t.col)}${r}` };
				writeRendered(cell, renderTemplate(tpl, scopes[i], ctx));
			} else if (isFormulaValue(t.value) && "formula" in t.value && t.value.formula) {
				const formula = mapFormulaRefs(t.value.formula, ws.name, ws.name, copyMapper(r, delta, i));
				cell.value = { formula } as ExcelJS.CellFormulaValue;
			} else {
				cell.value = clone(t.value);
			}
		}
	}

	for (const m of affected) {
		if (m.top === r && m.bottom === r) {
			for (let i = 0; i < n; i++) merge(ws, { ...m, top: r + i, bottom: r + i });
		} else if (m.top > r) {
			merge(ws, { ...m, top: m.top + delta, bottom: m.bottom + delta });
		} else {
			report.warn(
				"merge",
				`Celda combinada ${colName(m.left)}${m.top}:${colName(m.right)}${m.bottom} cruza la fila #each ${r} de "${ws.name}"; se descombina`,
			);
		}
	}
}

const INVALID_SHEET_CHARS = /[[\]:*?/\\]/g;

/** Excel sheet names: max 31 chars, none of []:*?/\, not empty, unique (case-insensitive). */
export function sanitizeSheetName(name: string, taken: Set<string>): string {
	let base = name.replace(INVALID_SHEET_CHARS, " ").replace(/\s+/g, " ").trim().replace(/^'+|'+$/g, "");
	if (!base) base = "Hoja";
	base = base.slice(0, 31);
	let candidate = base;
	for (let i = 2; taken.has(candidate.toLowerCase()); i++) {
		const suffix = ` (${i})`;
		candidate = base.slice(0, 31 - suffix.length) + suffix;
	}
	taken.add(candidate.toLowerCase());
	return candidate;
}

/** Copies a worksheet (values, styles, widths, merges, images) into a new sheet. */
export function cloneWorksheet(wb: Workbook, src: Worksheet, name: string, rt: Runtime): Worksheet {
	const ws = wb.addWorksheet(name);
	const model = structuredClone(src.model) as unknown as Record<string, unknown>;
	model.id = ws.id;
	model.name = name;
	model.mergeCells = model.merges;
	if (Array.isArray(model.tables) && model.tables.length) {
		rt.report.warn("table", `Las Tablas de Excel de "${src.name}" no se copian a las hojas clonadas`);
		model.tables = [];
	}
	(ws as unknown as { model: unknown }).model = model;
	return ws;
}
