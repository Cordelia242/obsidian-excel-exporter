import { parseCellText, type CellTemplate } from "./placeholders";
import { cellText, isFormulaValue, type Worksheet } from "./workbook-io";

export interface ScannedCell {
	row: number;
	col: number;
	address: string;
	text: string;
	tpl: CellTemplate;
}

export interface EachRow {
	row: number;
	col: number;
	address: string;
	tpl: CellTemplate;
}

export interface SheetScan {
	cells: ScannedCell[];
	eachRows: EachRow[];
	issues: Array<{ address: string; message: string }>;
}

/** Finds placeholders and `#each` rows in a worksheet (§6). */
export function scanWorksheet(ws: Worksheet): SheetScan {
	const cells: ScannedCell[] = [];
	const eachRows: EachRow[] = [];
	const issues: SheetScan["issues"] = [];
	ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
		let eachInRow: EachRow | undefined;
		row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
			if (cell.isMerged && cell.master !== cell) return;
			if (isFormulaValue(cell.value)) return;
			const text = cellText(cell.value);
			if (text === null) return;
			const tpl = parseCellText(text);
			if (!tpl) return;
			for (const message of tpl.issues) issues.push({ address: cell.address, message });
			cells.push({ row: rowNumber, col: colNumber, address: cell.address, text, tpl });
			if (tpl.each) {
				if (eachInRow) {
					issues.push({
						address: cell.address,
						message: `Solo se permite un {{#each}} por fila; se usa el de ${eachInRow.address}`,
					});
				} else {
					eachInRow = { row: rowNumber, col: colNumber, address: cell.address, tpl };
					eachRows.push(eachInRow);
				}
			}
		});
	});
	return { cells, eachRows, issues };
}
