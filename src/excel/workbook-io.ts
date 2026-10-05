import ExcelJS from "exceljs";

export type Workbook = ExcelJS.Workbook;
export type Worksheet = ExcelJS.Worksheet;

export async function loadWorkbook(data: ArrayBuffer): Promise<Workbook> {
	const wb = new ExcelJS.Workbook();
	await wb.xlsx.load(data);
	return wb;
}

export async function saveWorkbook(wb: Workbook): Promise<ArrayBuffer> {
	const out: unknown = await wb.xlsx.writeBuffer();
	if (out instanceof ArrayBuffer) return out;
	const view = out as Uint8Array;
	return view.buffer.slice(view.byteOffset, view.byteOffset + view.byteLength) as ArrayBuffer;
}

/** Cell text if the cell holds text (plain, rich text or hyperlink text); otherwise null. */
export function cellText(value: ExcelJS.CellValue): string | null {
	if (typeof value === "string") return value;
	if (value && typeof value === "object") {
		if ("richText" in value && Array.isArray(value.richText)) return value.richText.map((r) => r.text).join("");
		if ("hyperlink" in value && "text" in value && typeof value.text === "string") return value.text;
	}
	return null;
}

export function isFormulaValue(value: ExcelJS.CellValue): value is ExcelJS.CellFormulaValue {
	return !!value && typeof value === "object" && ("formula" in value || "sharedFormula" in value);
}

/** Excel stores wall-clock time; ExcelJS converts Dates using UTC, so encode local components as UTC. */
export function toExcelDate(d: Date): Date {
	return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()));
}

/**
 * Replaces shared formulas by plain per-cell formulas, so rows can be moved
 * and rewritten independently.
 */
export function unshareFormulas(wb: Workbook): void {
	wb.eachSheet((ws) => {
		const pending: Array<{ cell: ExcelJS.Cell; formula: string; result: unknown }> = [];
		ws.eachRow({ includeEmpty: false }, (row) => {
			row.eachCell({ includeEmpty: false }, (cell) => {
				const v = cell.value;
				if (isFormulaValue(v) && ("sharedFormula" in v || (v as { shareType?: string }).shareType === "shared")) {
					pending.push({ cell, formula: cell.formula, result: (v as { result?: unknown }).result });
				}
			});
		});
		for (const p of pending) {
			p.cell.value = { formula: p.formula, result: p.result } as ExcelJS.CellFormulaValue;
		}
	});
}
