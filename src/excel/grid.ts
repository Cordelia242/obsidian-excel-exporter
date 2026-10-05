import type ExcelJS from "exceljs";
import { colLetters } from "../config/mapping";
import { anyToText, formatDate } from "../values/normalize";
import type { Worksheet } from "./workbook-io";

export interface GridStyle {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	color?: string;
	bg?: string;
	align?: string;
	valign?: string;
	wrap?: boolean;
	size?: number;
	borders?: { top?: boolean; right?: boolean; bottom?: boolean; left?: boolean };
}

export interface GridCell {
	row: number;
	col: number;
	address: string;
	text: string;
	kind: "empty" | "text" | "number" | "date" | "bool" | "formula" | "link" | "error";
	style: GridStyle;
	rowSpan: number;
	colSpan: number;
	/** Covered by a merge: not rendered. */
	hidden: boolean;
	/** Raw formula, when the cell has one. */
	formula?: string;
}

export interface GridRow {
	index: number;
	height: number;
	cells: GridCell[];
}

export interface Grid {
	sheet: string;
	cols: Array<{ index: number; letter: string; width: number }>;
	rows: GridRow[];
	/** True when the sheet is larger than what was rendered. */
	truncated: boolean;
}

export interface GridOptions {
	maxRows?: number;
	maxCols?: number;
	/** Render at least this many rows/cols (Setup needs empty space to click). */
	minRows?: number;
	minCols?: number;
}

function argbToCss(argb: string | undefined): string | undefined {
	if (!argb || !/^[0-9a-f]{8}$/i.test(argb)) return undefined;
	return `#${argb.slice(2)}`;
}

/** ExcelJS stores dates as UTC wall-clock; show them with the cell's numFmt when it is a simple date format. */
export function excelDateText(d: Date, numFmt?: string): string {
	const local = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds());
	const fmt = (numFmt ?? "").toLowerCase();
	const hasTime = /h/.test(fmt) || local.getHours() !== 0 || local.getMinutes() !== 0;
	let datePart = "yyyy-MM-dd";
	if (/d{1,2}[/.-]m{1,2}[/.-]y{2,4}/.test(fmt)) datePart = "dd/MM/yyyy";
	else if (/m{1,2}[/.-]d{1,2}[/.-]y{2,4}/.test(fmt)) datePart = "MM/dd/yyyy";
	return formatDate(local, hasTime ? `${datePart} HH:mm` : datePart);
}

function numberText(n: number, numFmt?: string): string {
	if (numFmt && numFmt.includes("%")) return `${+(n * 100).toFixed(2)}%`;
	if (Number.isInteger(n)) return String(n);
	return String(+n.toFixed(4));
}

/** Text and kind of a cell value, as Excel would roughly show it. */
export function cellDisplay(value: ExcelJS.CellValue, numFmt?: string): { text: string; kind: GridCell["kind"]; formula?: string } {
	if (value === null || value === undefined || value === "") return { text: "", kind: "empty" };
	if (typeof value === "string") return { text: value, kind: "text" };
	if (typeof value === "number") return { text: numberText(value, numFmt), kind: "number" };
	if (typeof value === "boolean") return { text: value ? "VERDADERO" : "FALSO", kind: "bool" };
	if (value instanceof Date) return { text: excelDateText(value, numFmt), kind: "date" };
	if (typeof value === "object") {
		if ("richText" in value && Array.isArray(value.richText)) return { text: value.richText.map((r) => r.text).join(""), kind: "text" };
		if ("hyperlink" in value) {
			const t = (value as { text?: unknown }).text;
			return { text: typeof t === "string" ? t : String(value.hyperlink), kind: "link" };
		}
		if ("formula" in value || "sharedFormula" in value) {
			const f = value as { formula?: string; result?: unknown };
			const formula = f.formula ?? "";
			if (f.result !== undefined && f.result !== null && typeof f.result !== "object") {
				return { ...cellDisplay(f.result as ExcelJS.CellValue, numFmt), kind: "formula", formula };
			}
			return { text: `=${formula}`, kind: "formula", formula };
		}
		if ("error" in value) return { text: String(value.error), kind: "error" };
	}
	return { text: anyToText(value), kind: "text" };
}

function styleOf(cell: ExcelJS.Cell): GridStyle {
	const s: GridStyle = {};
	const font = cell.font;
	if (font) {
		if (font.bold) s.bold = true;
		if (font.italic) s.italic = true;
		if (font.underline) s.underline = true;
		if (font.size) s.size = font.size;
		s.color = argbToCss(font.color?.argb);
	}
	const fill = cell.fill as ExcelJS.FillPattern | undefined;
	if (fill && fill.type === "pattern" && fill.pattern === "solid") s.bg = argbToCss(fill.fgColor?.argb);
	const al = cell.alignment;
	if (al) {
		if (al.horizontal) s.align = al.horizontal === "centerContinuous" ? "center" : al.horizontal;
		if (al.vertical) s.valign = al.vertical === "middle" ? "middle" : al.vertical;
		if (al.wrapText) s.wrap = true;
	}
	const b = cell.border;
	if (b && (b.top || b.right || b.bottom || b.left)) {
		s.borders = { top: !!b.top?.style, right: !!b.right?.style, bottom: !!b.bottom?.style, left: !!b.left?.style };
	}
	return s;
}

/** Converts a worksheet into a renderable grid (pure: no DOM). */
export function worksheetToGrid(ws: Worksheet, opts: GridOptions = {}): Grid {
	const maxRows = opts.maxRows ?? 300;
	const maxCols = opts.maxCols ?? 40;
	const usedRows = ws.rowCount;
	const usedCols = ws.columnCount;
	const nRows = Math.min(Math.max(usedRows, opts.minRows ?? 0), maxRows);
	const nCols = Math.min(Math.max(usedCols, opts.minCols ?? 0), maxCols);

	const merges = Object.values((ws as unknown as { _merges: Record<string, { top: number; left: number; bottom: number; right: number }> })._merges ?? {});
	const covered = new Set<string>();
	const spans = new Map<string, { rowSpan: number; colSpan: number }>();
	for (const m of merges) {
		spans.set(`${m.top}:${m.left}`, { rowSpan: m.bottom - m.top + 1, colSpan: m.right - m.left + 1 });
		for (let r = m.top; r <= m.bottom; r++) for (let c = m.left; c <= m.right; c++) if (r !== m.top || c !== m.left) covered.add(`${r}:${c}`);
	}

	const cols = [];
	for (let c = 1; c <= nCols; c++) {
		const w = ws.getColumn(c).width;
		cols.push({ index: c, letter: colLetters(c), width: Math.round((w ?? 8.43) * 7 + 5) });
	}
	const rows: GridRow[] = [];
	for (let r = 1; r <= nRows; r++) {
		const row = ws.findRow(r);
		const cells: GridCell[] = [];
		for (let c = 1; c <= nCols; c++) {
			const cell = row?.findCell(c);
			const key = `${r}:${c}`;
			const span = spans.get(key) ?? { rowSpan: 1, colSpan: 1 };
			const disp = cell ? cellDisplay(cell.isMerged && cell.master !== cell ? null : cell.value, cell.numFmt) : { text: "", kind: "empty" as const };
			cells.push({
				row: r,
				col: c,
				address: `${colLetters(c)}${r}`,
				text: disp.text,
				kind: disp.kind,
				formula: "formula" in disp ? disp.formula : undefined,
				style: cell ? styleOf(cell) : {},
				rowSpan: Math.min(span.rowSpan, nRows - r + 1),
				colSpan: Math.min(span.colSpan, nCols - c + 1),
				hidden: covered.has(key),
			});
		}
		rows.push({ index: r, height: Math.round((row?.height ?? 15) * (4 / 3)), cells });
	}
	return { sheet: ws.name, cols, rows, truncated: usedRows > maxRows || usedCols > maxCols };
}
