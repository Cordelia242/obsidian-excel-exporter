import { parsePath } from "../graph/path";
import { parseCellText } from "../excel/placeholders";

/** A cell filled by the Setup mapping. `sheet` null = first sheet. */
export interface CellMapping {
	sheet: string | null;
	address: string;
	row: number;
	col: number;
	/** Template text, always with `{{ }}`. */
	expr: string;
}

export interface RowMapping {
	sheet: string | null;
	row: number;
	src: string;
	path: string[];
}

export interface CompiledMapping {
	cells: CellMapping[];
	rows: RowMapping[];
}

const CELL_KEY = /^(?:(.+)!)?\$?([A-Za-z]{1,3})\$?(\d+)$/;
const ROW_KEY = /^(?:(.+)!)?(\d+)$/;

export function colNumber(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
	return n;
}

export function colLetters(col: number): string {
	let s = "";
	for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
	return s;
}

function unquoteSheet(s: string | undefined): string | null {
	if (!s) return null;
	const t = s.trim();
	return t.startsWith("'") && t.endsWith("'") ? t.slice(1, -1).replace(/''/g, "'") : t;
}

/** `Hoja!B2` key for a cell; quotes sheet names that need it. */
export function cellKey(sheet: string, address: string): string {
	return `${sheet}!${address}`;
}

export function rowKey(sheet: string, row: number): string {
	return `${sheet}!${row}`;
}

/** A bare path (`proceso.team | upper`) becomes `{{proceso.team | upper}}`. */
export function toTemplateText(value: string): string {
	return value.includes("{{") ? value : `{{${value.trim()}}}`;
}

/** Strips one surrounding `{{ }}` when the whole text is a single placeholder. */
export function toBareExpression(value: string): string {
	const m = /^\s*\{\{([^{}]*)\}\}\s*$/.exec(value);
	return m ? m[1].trim() : value;
}

export function compileMapping(cells: Record<string, string>, rows: Record<string, string>, errors: string[]): CompiledMapping {
	const out: CompiledMapping = { cells: [], rows: [] };
	for (const [key, value] of Object.entries(cells)) {
		const m = CELL_KEY.exec(key.trim());
		if (!m) {
			errors.push(`\`cells\`: celda inválida "${key}" (formato Hoja!B2)`);
			continue;
		}
		const expr = toTemplateText(value);
		const tpl = parseCellText(expr);
		for (const issue of tpl?.issues ?? []) errors.push(`\`cells.${key}\`: ${issue}`);
		for (const part of tpl?.parts ?? []) {
			if (part.kind === "ph" && part.error) errors.push(`\`cells.${key}\`: ${part.error}`);
		}
		const col = colNumber(m[2]);
		const row = Number(m[3]);
		out.cells.push({ sheet: unquoteSheet(m[1]), address: `${colLetters(col)}${row}`, row, col, expr });
	}
	for (const [key, value] of Object.entries(rows)) {
		const m = ROW_KEY.exec(key.trim());
		if (!m) {
			errors.push(`\`rows\`: fila inválida "${key}" (formato Hoja!7)`);
			continue;
		}
		const src = value.replace(/^\s*\{\{\s*(#each\s+)?|\s*\}\}\s*$/g, "").trim();
		try {
			out.rows.push({ sheet: unquoteSheet(m[1]), row: Number(m[2]), src, path: parsePath(src) });
		} catch (e) {
			errors.push(`\`rows.${key}\`: ${(e as Error).message}`);
		}
	}
	return out;
}

/** Whether a mapping entry targets the given sheet (null = the first sheet). */
export function targetsSheet(entrySheet: string | null, sheetName: string, firstSheet: string): boolean {
	const target = entrySheet ?? firstSheet;
	return target.toLowerCase() === sheetName.toLowerCase();
}
