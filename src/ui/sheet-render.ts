import type { Grid, GridCell } from "../excel/grid";

export interface GridRenderOptions {
	/** Overrides what a cell shows (e.g. the mapped field in Setup). */
	cellContent?: (cell: GridCell) => { text: string; cls?: string; title?: string } | null;
	onCellClick?: (cell: GridCell) => void;
	onRowClick?: (row: number) => void;
	rowClass?: (row: number) => string | undefined;
	rowTitle?: (row: number) => string | undefined;
}

export interface RenderedGrid {
	cells: Map<string, HTMLTableCellElement>;
	rows: Map<number, HTMLTableRowElement>;
	refreshCell: (address: string) => void;
}

/** Cell formatting from the Excel file: fixed choices as CSS classes, colors/sizes (data) via setCssStyles. */
function applyStyle(td: HTMLTableCellElement, cell: GridCell): void {
	const s = cell.style;
	if (s.bold) td.addClass("xte-b");
	if (s.italic) td.addClass("xte-i");
	if (s.underline) td.addClass("xte-u");
	if (s.wrap) td.addClass("xte-wrap");
	const align = s.align ?? (cell.kind === "number" || cell.kind === "date" ? "right" : cell.kind === "bool" ? "center" : "left");
	td.addClass(`xte-al-${align}`);
	if (s.valign) td.addClass(`xte-va-${s.valign}`);
	const b = s.borders;
	if (b?.top) td.addClass("xte-bt");
	if (b?.right) td.addClass("xte-br");
	if (b?.bottom) td.addClass("xte-bb");
	if (b?.left) td.addClass("xte-bl");
	const dynamic: Partial<CSSStyleDeclaration> = {};
	if (s.color) dynamic.color = s.color;
	if (s.bg) dynamic.backgroundColor = s.bg;
	if (s.size) dynamic.fontSize = `${Math.round(s.size * 1.1)}px`;
	if (Object.keys(dynamic).length) td.setCssStyles(dynamic);
}

/** Renders a grid as an Excel-like HTML table. */
export function renderGrid(parent: HTMLElement, grid: Grid, o: GridRenderOptions = {}): RenderedGrid {
	const wrap = parent.createDiv({ cls: "xte-grid-wrap" });
	const table = wrap.createEl("table", { cls: "xte-grid" });
	const colgroup = table.createEl("colgroup");
	colgroup.createEl("col", { cls: "xte-col-head" });
	// Column widths and row heights come from the Excel file.
	for (const c of grid.cols) colgroup.createEl("col").setCssStyles({ width: `${c.width}px` });
	// Fixed width so columns keep their Excel widths instead of shrinking to the container.
	table.setCssStyles({ width: `${44 + grid.cols.reduce((sum, c) => sum + c.width, 0)}px` });
	const head = table.createEl("thead").createEl("tr");
	head.createEl("th", { cls: "xte-corner" });
	for (const c of grid.cols) head.createEl("th", { text: c.letter });

	const body = table.createEl("tbody");
	const cells = new Map<string, HTMLTableCellElement>();
	const rows = new Map<number, HTMLTableRowElement>();
	const byAddress = new Map<string, GridCell>();

	const fill = (td: HTMLTableCellElement, cell: GridCell) => {
		td.empty();
		td.className = `xte-cell xte-kind-${cell.kind}`;
		applyStyle(td, cell);
		const custom = o.cellContent?.(cell);
		if (custom) {
			td.createSpan({ text: custom.text, cls: custom.cls });
			if (custom.title) td.title = custom.title;
		} else {
			td.setText(cell.text);
			td.title = cell.formula ? `=${cell.formula}` : cell.text;
		}
	};

	for (const r of grid.rows) {
		const tr = body.createEl("tr");
		tr.setCssStyles({ height: `${r.height}px` });
		const extra = o.rowClass?.(r.index);
		if (extra) tr.addClass(extra);
		const th = tr.createEl("th", { text: String(r.index), cls: "xte-rowhead" });
		const title = o.rowTitle?.(r.index);
		if (title) th.title = title;
		if (o.onRowClick) th.onclick = () => o.onRowClick?.(r.index);
		rows.set(r.index, tr);
		for (const cell of r.cells) {
			if (cell.hidden) continue;
			const td = tr.createEl("td");
			if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
			if (cell.colSpan > 1) td.colSpan = cell.colSpan;
			fill(td, cell);
			if (o.onCellClick) td.onclick = () => o.onCellClick?.(cell);
			cells.set(cell.address, td);
			byAddress.set(cell.address, cell);
		}
	}
	if (grid.truncated) wrap.createDiv({ cls: "xte-muted", text: "La hoja es más grande: se muestra solo una parte." });

	return {
		cells,
		rows,
		refreshCell: (address) => {
			const td = cells.get(address);
			const cell = byAddress.get(address);
			if (td && cell) fill(td, cell);
		},
	};
}

/** Sheet tabs; returns nothing, calls `onSelect` with the index. */
export function renderTabs(parent: HTMLElement, names: string[], active: number, onSelect: (i: number) => void): void {
	const tabs = parent.createDiv({ cls: "xte-tabs" });
	names.forEach((name, i) => {
		const t = tabs.createEl("button", { text: name, cls: i === active ? "xte-tab is-active" : "xte-tab" });
		t.onclick = () => onSelect(i);
	});
}
