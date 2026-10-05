import { setIcon } from "obsidian";
import { colLetters, toBareExpression } from "../../config/mapping";
import { worksheetToGrid, type GridCell } from "../../excel/grid";
import type { Workbook, Worksheet } from "../../excel/workbook-io";
import { parsePath } from "../../graph/path";
import {
	baseScope,
	cellSources,
	collectionOptions,
	composeExpr,
	fieldsAt,
	formatOptions,
	friendlyLabel,
	parseExprParts,
	previewExpression,
	rootLabel,
	rowSampleScope,
	variableLabel,
	type CellContext,
	type ExprParts,
	type FieldEntry,
	type Source,
} from "../../setup/fields";
import { button, callout, choiceCards, combobox, iconButton, pillChoices } from "../components";
import { renderGrid, type RenderedGrid } from "../sheet-render";
import type { TemplateSession } from "./session";

type Selection = { type: "cell"; row: number; col: number; address: string } | { type: "row"; row: number } | null;

const KIND_ICON: Record<string, string> = {
	text: "type",
	number: "hash",
	date: "calendar",
	stars: "star",
	link: "link",
	list: "list",
	bool: "check-square",
	note: "file-text",
	collection: "layers",
	file: "file",
	special: "at-sign",
	empty: "circle",
};

/** Step 3: the workbook with clickable cells and rows. Keeps its UI state between renders. */
export class CellsStep {
	private sheetIdx = 0;
	private sampleIdx = 0;
	private showValues = false;
	private selection: Selection = null;
	private sourceId: string | null = null;
	private trail: string[] = [];
	private query = "";
	private grid: RenderedGrid | null = null;
	private tabCounts = new Map<string, HTMLElement>();
	private wb: Workbook | null = null;

	constructor(private s: TemplateSession) {}

	private get cfg() {
		return this.s.cfg;
	}

	private sheet(): Worksheet | undefined {
		return this.wb?.worksheets[this.sheetIdx];
	}

	private perRoot(ws: Worksheet): boolean {
		const cfg = this.cfg;
		if (cfg.mode === "file-per-root") return true;
		if (cfg.mode === "single") return false;
		const name = cfg.output.repeatSheet;
		const repeated = (name && this.wb?.worksheets.find((w) => w.name.toLowerCase() === name.toLowerCase())) || this.wb?.worksheets[0];
		return repeated === ws;
	}

	private key(row: number, col?: number): string {
		const ws = this.sheet();
		return col === undefined ? `${ws?.name}!${row}` : `${ws?.name}!${colLetters(col)}${row}`;
	}

	private rowPath(row: number): string[] | undefined {
		const src = this.cfg.rows[this.key(row)];
		if (!src) return undefined;
		try {
			return parsePath(src);
		} catch {
			return undefined;
		}
	}

	private ctx(row: number): CellContext | null {
		const run = this.s.run();
		const ws = this.sheet();
		if (!run || !ws) return null;
		return { run, perRoot: this.perRoot(ws), rowPath: this.rowPath(row), sampleIndex: this.sampleIdx };
	}

	private sampleValue(expr: string, row: number): string {
		const run = this.s.run();
		const ws = this.sheet();
		if (!run || !ws) return "";
		const base = baseScope(run, this.perRoot(ws), this.sampleIdx);
		if (!base) return "";
		const scope = rowSampleScope(base, this.rowPath(row), run.rt);
		if (!scope) return "(sin elementos en la muestra)";
		try {
			return previewExpression(expr, scope, run);
		} catch (e) {
			return `⚠ ${(e as Error).message}`;
		}
	}

	/** Moves mapping keys without a sheet ("B2") to the first sheet ("Hoja!B2"). */
	private normalizeKeys(): void {
		const first = this.wb?.worksheets[0]?.name;
		if (!first) return;
		for (const map of [this.cfg.cells, this.cfg.rows]) {
			for (const k of Object.keys(map)) {
				if (!k.includes("!")) {
					map[`${first}!${k}`] = map[k];
					delete map[k];
				}
			}
		}
	}

	async render(el: HTMLElement): Promise<void> {
		this.wb = await this.s.workbook();
		el.empty();
		if (!this.wb) {
			callout(el, `No se encontró el archivo Excel «${this.cfg.template}». Cámbialo con el botón «Excel» de arriba.`, "warn");
			return;
		}
		const run = this.s.run();
		if (!run) {
			callout(el, "Primero elige en el paso 1 qué notas exporta el template: los campos disponibles salen de esas notas.", "warn");
			return;
		}
		this.normalizeKeys();
		if (this.sheetIdx >= this.wb.worksheets.length) this.sheetIdx = 0;
		if (this.sampleIdx >= run.roots.length) this.sampleIdx = 0;
		const ws = this.sheet() as Worksheet;

		// Toolbar: sheets + sample
		const bar = el.createDiv({ cls: "xte-cells-bar" });
		const tabs = bar.createDiv({ cls: "xte-sheet-tabs" });
		this.wb.worksheets.forEach((w, i) => {
			const n = Object.keys(this.cfg.cells).filter((k) => k.startsWith(`${w.name}!`)).length;
			const t = tabs.createEl("button", { cls: `xte-sheet-tab${i === this.sheetIdx ? " is-active" : ""}` });
			setIcon(t.createSpan({ cls: "xte-btn-icon" }), this.perRoot(w) && this.cfg.mode === "sheet-per-root" ? "copy" : "sheet");
			t.createSpan({ text: w.name });
			this.tabCounts.set(w.name, t.createSpan({ cls: "xte-tab-count", text: String(n) }));
			t.title = `${n} ${n === 1 ? "celda con dato" : "celdas con dato"}`;
			t.onclick = () => {
				this.sheetIdx = i;
				this.selection = null;
				void this.render(el);
			};
		});
		const sampleBox = bar.createDiv({ cls: "xte-inline xte-sample" });
		sampleBox.createSpan({ cls: "xte-help", text: `${rootLabel(run.cd.def)} de ejemplo:` });
		combobox(sampleBox, {
			value: String(this.sampleIdx),
			options: run.roots.map((r, i) => ({ value: String(i), label: r.basename })),
			onChange: (v) => {
				this.sampleIdx = Number(v);
				void this.render(el);
			},
		});
		const toggle = sampleBox.createEl("label", { cls: "xte-switch" });
		const cb = toggle.createEl("input", { type: "checkbox" });
		cb.checked = this.showValues;
		toggle.appendText(" Ver valores");
		cb.onchange = () => {
			this.showValues = cb.checked;
			this.drawGrid(gridEl, panel);
		};

		// What this sheet is
		const label = rootLabel(run.cd.def);
		const how =
			this.cfg.mode === "file-per-root"
				? `Esta hoja se llena una vez por cada ${label} (un archivo por ${label}).`
				: this.cfg.mode === "single"
					? `Esta hoja se llena una sola vez con todos los ${label}s: usa una fila repetida para listarlos.`
					: this.perRoot(ws)
						? `Esta hoja se copia una vez por cada ${label} (una hoja por ${label}).`
						: `Esta hoja aparece una sola vez (resumen con todos los ${label}s).`;
		const info = el.createDiv({ cls: "xte-sheet-info" });
		info.createSpan({ text: how });
		const legend = info.createDiv({ cls: "xte-legend" });
		legend.createSpan({ cls: "xte-legend-cell", text: "celda con dato" });
		legend.createSpan({ cls: "xte-legend-row", text: "fila que se repite" });

		const body = el.createDiv({ cls: "xte-cells-body" });
		const gridEl = body.createDiv({ cls: "xte-cells-grid" });
		const panel = body.createDiv({ cls: "xte-cells-panel" });
		this.drawGrid(gridEl, panel);
	}

	private drawGrid(gridEl: HTMLElement, panel: HTMLElement): void {
		const ws = this.sheet();
		const run = this.s.run();
		if (!ws || !run) return;
		gridEl.empty();
		const def = run.cd.def;
		const grid = worksheetToGrid(ws, { minRows: Math.max(ws.rowCount + 8, 25), minCols: Math.max(ws.columnCount + 2, 10) });
		this.grid = renderGrid(gridEl, grid, {
			cellContent: (cell: GridCell) => {
				const expr = this.cfg.cells[this.key(cell.row, cell.col)];
				if (!expr) return null;
				let nice = friendlyLabel(expr, def);
				// Inside a repeated row, "Interviews › decision" reads better as just "decision".
				const rp = this.rowPath(cell.row);
				if (rp) {
					const prefix = `${variableLabel(def, rp[rp.length - 1])} › `;
					if (nice.startsWith(prefix)) nice = `↻ ${nice.slice(prefix.length)}`;
				}
				if (this.showValues) {
					const v = this.sampleValue(expr, cell.row);
					return { text: v || "(vacío)", cls: "xte-mapped-value", title: `${nice}\n→ ${v}` };
				}
				return { text: nice, cls: "xte-mapped", title: nice };
			},
			onCellClick: (cell) => this.select({ type: "cell", row: cell.row, col: cell.col, address: cell.address }, panel),
			onRowClick: (row) => this.select({ type: "row", row }, panel),
			rowClass: (row) => (this.cfg.rows[this.key(row)] ? "xte-row-repeat" : undefined),
			rowTitle: (row) => {
				const p = this.rowPath(row);
				return p ? `Se repite por cada ${variableLabel(def, p[p.length - 1])}` : "Clic para que esta fila se repita";
			},
		});
		for (const [row, tr] of this.grid.rows) {
			const p = this.rowPath(row);
			if (p) {
				const th = tr.querySelector("th");
				if (th) {
					th.empty();
					setIcon(th.createSpan({ cls: "xte-row-icon" }), "repeat");
					th.createSpan({ text: String(row) });
				}
			}
		}
		this.highlight();
		this.renderPanel(panel);
	}

	private highlight(): void {
		if (!this.grid) return;
		for (const td of this.grid.cells.values()) td.removeClass("is-selected");
		for (const tr of this.grid.rows.values()) tr.removeClass("is-selected");
		const sel = this.selection;
		if (sel?.type === "cell") this.grid.cells.get(sel.address)?.addClass("is-selected");
		if (sel?.type === "row") this.grid.rows.get(sel.row)?.addClass("is-selected");
	}

	private select(sel: Selection, panel: HTMLElement): void {
		this.selection = sel;
		this.query = "";
		this.sourceId = null;
		this.trail = [];
		if (sel?.type === "cell") this.syncNavigation(sel.row);
		this.highlight();
		this.renderPanel(panel);
	}

	/** Opens the source/trail matching the cell's current mapping. */
	private syncNavigation(row: number): void {
		const sel = this.selection;
		const ctx = this.ctx(row);
		if (sel?.type !== "cell" || !ctx) return;
		const parts = parseExprParts(this.cfg.cells[this.key(sel.row, sel.col)] ?? "");
		const sources = cellSources(ctx);
		if (!parts) return;
		if (parts.segs[0].startsWith("@")) {
			this.sourceId = "special";
			return;
		}
		const match = sources
			.filter((src) => src.base.length && src.base.every((b, i) => parts.segs[i]?.toLowerCase() === b.toLowerCase()))
			.sort((a, b) => b.base.length - a.base.length)[0];
		if (!match) return;
		this.sourceId = match.id;
		const rest = parts.segs.slice(match.base.length);
		this.trail = rest.length > 1 ? rest.slice(0, -1) : [];
		if (rest[rest.length - 2] === "file") this.trail = rest.slice(0, -2);
	}

	// ------------------------------------------------------------------ panel

	private renderPanel(panel: HTMLElement): void {
		panel.empty();
		const sel = this.selection;
		if (!sel) this.renderOverview(panel);
		else if (sel.type === "row") this.renderRowPanel(panel, sel.row);
		else this.renderCellPanel(panel, sel.row, sel.col);
	}

	private renderOverview(panel: HTMLElement): void {
		const run = this.s.run();
		const ws = this.sheet();
		if (!run || !ws) return;
		const def = run.cd.def;
		panel.createEl("h4", { text: "Cómo configurar esta hoja" });
		const ol = panel.createEl("ol", { cls: "xte-howto" });
		ol.createEl("li", { text: "Haz clic en una celda vacía del Excel y elige qué dato va ahí." });
		ol.createEl("li", { text: "Para una tabla (una fila por Interview, por ejemplo), haz clic en el número de la fila y elige por qué se repite." });
		ol.createEl("li", { text: "Activa «Ver valores» para comprobar cómo queda con un ejemplo real." });

		const cells = Object.entries(this.cfg.cells).filter(([k]) => k.startsWith(`${ws.name}!`));
		const rows = Object.entries(this.cfg.rows).filter(([k]) => k.startsWith(`${ws.name}!`));
		if (rows.length) {
			panel.createEl("h5", { text: "Filas que se repiten" });
			for (const [k, v] of rows) {
				const row = Number(k.slice(k.lastIndexOf("!") + 1));
				const item = panel.createDiv({ cls: "xte-overview-item" });
				setIcon(item.createSpan({ cls: "xte-btn-icon" }), "repeat");
				item.createSpan({ text: `Fila ${row}: una por cada ${variableLabel(def, parsePathSafe(v).slice(-1)[0] ?? v)}` });
				item.onclick = () => this.select({ type: "row", row }, panel);
			}
		}
		panel.createEl("h5", { text: `Celdas con dato (${cells.length})` });
		if (!cells.length) panel.createDiv({ cls: "xte-help", text: "Ninguna todavía." });
		for (const [k, v] of cells) {
			const addr = k.slice(k.lastIndexOf("!") + 1);
			const m = /^([A-Z]+)(\d+)$/.exec(addr);
			if (!m) continue;
			const item = panel.createDiv({ cls: "xte-overview-item" });
			item.createSpan({ cls: "xte-addr", text: addr });
			item.createSpan({ text: friendlyLabel(v, def) });
			const col = m[1].split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
			item.onclick = () => this.select({ type: "cell", row: Number(m[2]), col, address: addr }, panel);
		}
	}

	private renderRowPanel(panel: HTMLElement, row: number): void {
		const run = this.s.run();
		const ws = this.sheet();
		if (!run || !ws) return;
		const def = run.cd.def;
		const head = panel.createDiv({ cls: "xte-panel-head" });
		head.createEl("h4", { text: `Fila ${row}` });
		iconButton(head, "x", "Cerrar", () => this.select(null, panel));

		panel.createDiv({ cls: "xte-q", text: "¿Esta fila se repite?" });
		const key = this.key(row);
		const current = this.cfg.rows[key] ?? "";
		const options = collectionOptions(run, this.perRoot(ws), this.sampleIdx);
		const sample = run.roots[this.sampleIdx]?.basename ?? "";
		const cards = [
			{ value: "", title: "No, es una fila normal", icon: "minus" },
			...options.map((o) => ({
				value: o.value,
				title: `Sí, una fila por cada ${o.label}`,
				description: `${o.count} ${o.count === 1 ? "fila" : "filas"} con «${sample}»`,
				icon: "repeat",
			})),
		];
		if (current && !options.some((o) => o.value === current)) cards.push({ value: current, title: `Una fila por cada ${current}`, icon: "repeat" });
		choiceCards(panel, cards, current, (v) => {
			if (v) this.cfg.rows[key] = v;
			else delete this.cfg.rows[key];
			this.s.changed();
			const gridEl = panel.parentElement?.querySelector(".xte-cells-grid") as HTMLElement | null;
			if (gridEl) this.drawGrid(gridEl, panel);
		});
		if (!options.length) {
			callout(panel, `Para repetir filas necesitas datos relacionados (paso 2), por ejemplo las Interviews de cada ${rootLabel(def)}.`, "info");
		} else if (current) {
			callout(panel, `Ahora haz clic en las celdas de la fila ${row} y elige «${variableLabel(def, parsePathSafe(current).slice(-1)[0] ?? current)} de esta fila» como origen del dato.`, "info", "mouse-pointer-click");
		}
		panel.createDiv({
			cls: "xte-help",
			text: "Las filas de abajo se desplazan solas y las fórmulas que las usan (sumas, promedios) se ajustan. El formato de la fila se copia en cada repetición.",
		});
	}

	private renderCellPanel(panel: HTMLElement, row: number, col: number): void {
		const ctx = this.ctx(row);
		const ws = this.sheet();
		if (!ctx || !ws) return;
		const def = ctx.run.cd.def;
		const key = this.key(row, col);
		const expr = this.cfg.cells[key] ?? "";
		const parts = parseExprParts(expr);

		const head = panel.createDiv({ cls: "xte-panel-head" });
		head.createEl("h4", { text: `Celda ${colLetters(col)}${row}` });
		iconButton(head, "x", "Cerrar", () => this.select(null, panel));
		if (ctx.rowPath) {
			const pill = panel.createDiv({ cls: "xte-pill" });
			setIcon(pill.createSpan({ cls: "xte-btn-icon" }), "repeat");
			pill.createSpan({ text: `La fila ${row} se repite por cada ${variableLabel(def, ctx.rowPath[ctx.rowPath.length - 1])}` });
		}
		const tplText = ws.getCell(row, col).text;
		if (tplText && !expr) panel.createDiv({ cls: "xte-help", text: `En el Excel dice «${tplText}». Si eliges un dato, se reemplaza.` });

		const set = (next: string, focusTop = true) => {
			if (next.trim()) this.cfg.cells[key] = next.trim();
			else delete this.cfg.cells[key];
			this.s.changed();
			this.tabCounts.get(ws.name)?.setText(String(Object.keys(this.cfg.cells).filter((k) => k.startsWith(`${ws.name}!`)).length));
			this.grid?.refreshCell(`${colLetters(col)}${row}`);
			this.renderPanel(panel);
			if (focusTop) panel.scrollTop = 0;
		};

		const sources = cellSources(ctx);
		if (!this.sourceId) this.sourceId = sources[0]?.id ?? null;
		const source = sources.find((x) => x.id === this.sourceId);

		// The chosen data and how it is shown, on top.
		if (expr) {
			const cur = panel.createDiv({ cls: "xte-current" });
			const top = cur.createDiv({ cls: "xte-current-top" });
			const t = top.createDiv({ cls: "xte-current-text" });
			t.createDiv({ cls: "xte-current-caption", text: "Dato de esta celda" });
			t.createDiv({ cls: "xte-current-label", text: friendlyLabel(expr, def) });
			const v = this.sampleValue(expr, row);
			t.createDiv({ cls: "xte-current-value", text: v ? `Ejemplo: ${v}` : "Ejemplo: (vacío)" });
			iconButton(top, "eraser", "Quitar el dato de esta celda", () => set(""));
			if (parts) {
				const entry = this.findEntry(sources, parts, ctx);
				const opts = formatOptions(entry?.kind ?? "text", entry?.isList ?? false);
				if (opts.length > 1 || parts.format) {
					cur.createDiv({ cls: "xte-current-caption", text: "¿Cómo mostrarlo?" });
					const known = opts.some((o) => o.pipe === parts.format);
					const all = known ? opts : [...opts, { pipe: parts.format, label: `Personalizado (${parts.format})` }];
					pillChoices(cur, all.map((o) => ({ value: o.pipe, label: o.label })), parts.format, (pipe) => set(composeExpr({ ...parts, format: pipe })));
				}
				const fb = cur.createDiv({ cls: "xte-inline xte-fallback" });
				fb.createSpan({ cls: "xte-help", text: "Si no hay dato, mostrar:" });
				const fallback = fb.createEl("input", { type: "text", cls: "xte-input-sm", value: parts.fallback });
				fallback.placeholder = "(celda vacía)";
				fallback.onchange = () => set(composeExpr({ ...parts, fallback: fallback.value }));
			}
		}

		// Choosing (or changing) the data.
		const pick = panel.createDiv({ cls: "xte-qblock" });
		pick.createDiv({ cls: "xte-q", text: expr ? "Cambiar el dato" : "¿Qué dato va en esta celda?" });
		pick.createDiv({ cls: "xte-step-caption", text: "1. Elige de dónde sale" });
		pillChoices(
			pick,
			sources.map((src) => ({
				value: src.id,
				label: src.label,
				title: src.hint,
				icon: src.kind === "row" ? "repeat" : src.kind === "relation" ? "layers" : src.kind === "special" ? "calendar" : "file-text",
			})),
			this.sourceId ?? "",
			(id) => {
				this.sourceId = id;
				this.trail = [];
				this.query = "";
				this.renderPanel(panel);
			},
		);
		if (source) pick.createDiv({ cls: "xte-help", text: source.hint });
		if (source) this.renderFieldPicker(pick, source, ctx, parts, (e) => set(e), panel);

		// Advanced
		const adv = panel.createEl("details", { cls: "xte-advanced" });
		if (expr && !parts) adv.open = true;
		adv.createEl("summary", { text: "Avanzado: combinar texto y datos" });
		adv.createDiv({ cls: "xte-help", text: "Escribe texto y pon los datos entre llaves. Ejemplo: Rol: {{proceso.role}} ({{proceso.quantity}} vacantes)" });
		const ta = adv.createEl("textarea", { cls: "xte-expr-text" });
		ta.value = expr && !expr.includes("{{") ? `{{${expr}}}` : expr;
		ta.rows = 2;
		const prev = adv.createDiv({ cls: "xte-help" });
		const showPrev = () => prev.setText(ta.value.trim() ? `Ejemplo: ${this.sampleValue(ta.value, row) || "(vacío)"}` : "");
		showPrev();
		ta.oninput = showPrev;
		button(adv, "Aplicar", () => set(toBareExpression(ta.value)), { cls: "xte-link-btn" });
	}

	/** The field entry a mapping points to (for its type and list-ness). */
	private findEntry(sources: Source[], parts: ExprParts, ctx: CellContext): FieldEntry | undefined {
		if (parts.segs[0].startsWith("@")) return fieldsAt(sources[sources.length - 1], [], ctx).find((e) => sameSegs(e.segs, parts.segs));
		const candidates = sources
			.filter((src) => src.base.length && src.base.every((b, i) => parts.segs[i]?.toLowerCase() === b.toLowerCase()))
			.sort((a, b) => b.base.length - a.base.length);
		for (const src of candidates) {
			const rest = parts.segs.slice(src.base.length);
			for (const trail of [rest.slice(0, -1), rest, rest.slice(0, -2)]) {
				const e = fieldsAt(src, trail, ctx).find((x) => sameSegs(x.segs, parts.segs));
				if (e) return e;
			}
		}
		return undefined;
	}

	private renderFieldPicker(
		parent: HTMLElement,
		source: Source,
		ctx: CellContext,
		parts: ExprParts | null,
		set: (e: string) => void,
		panel: HTMLElement,
	): void {
		const def = ctx.run.cd.def;
		const step2 = parent.createDiv({ cls: "xte-fieldpicker" });
		step2.createDiv({ cls: "xte-step-caption", text: "2. Elige el campo" });

		// Breadcrumb
		if (this.trail.length) {
			const crumbs = step2.createDiv({ cls: "xte-crumbs" });
			const names = [source.label, ...this.trail.map((t) => (def.relations[t] ? variableLabel(def, t) : t))];
			names.forEach((n, i) => {
				if (i) setIcon(crumbs.createSpan({ cls: "xte-crumb-sep" }), "chevron-right");
				const c = crumbs.createEl("a", { cls: "xte-crumb", text: n, href: "#" });
				c.onclick = (ev) => {
					ev.preventDefault();
					this.trail = this.trail.slice(0, i);
					this.query = "";
					this.renderPanel(panel);
					panel.querySelector(".xte-fieldpicker")?.scrollIntoView({ block: "start" });
				};
			});
		}

		const entries = fieldsAt(source, this.trail, ctx);
		const search = step2.createEl("input", { type: "search", cls: "xte-field-search", placeholder: "Buscar campo…", value: this.query });
		const list = step2.createDiv({ cls: "xte-fields" });
		const draw = () => {
			list.empty();
			const q = this.query.trim().toLowerCase();
			const shown = entries.filter((e) => !q || e.label.toLowerCase().includes(q) || e.sample.toLowerCase().includes(q));
			if (!shown.length) list.createDiv({ cls: "xte-help", text: "Ningún campo coincide." });
			for (const e of shown) this.fieldRow(list, e, parts, set, panel);
		};
		search.oninput = () => {
			this.query = search.value;
			draw();
		};
		draw();
		if (source.isList && !this.trail.length) {
			step2.createDiv({
				cls: "xte-help",
				text: "Estos datos tienen varios valores: luego eliges si mostrar la cantidad, la lista o solo el primero. Para una fila por cada uno, repite la fila (clic en su número).",
			});
		}
	}

	private fieldRow(list: HTMLElement, e: FieldEntry, parts: ExprParts | null, set: (expr: string) => void, panel: HTMLElement): void {
		const selected = parts && sameSegs(parts.segs, e.segs);
		const row = list.createDiv({ cls: `xte-field-row${selected ? " is-selected" : ""}` });
		const main = row.createDiv({ cls: "xte-field-main" });
		setIcon(main.createSpan({ cls: "xte-field-kind" }), KIND_ICON[e.kind] ?? "circle");
		const txt = main.createDiv({ cls: "xte-field-text" });
		txt.createDiv({ cls: "xte-field-name", text: e.label });
		if (e.sample) txt.createDiv({ cls: "xte-field-ex", text: e.sample });
		main.title = "Usar este dato";
		main.onclick = () => {
			// A whole list (e.g. "Contratados (todas)") is usually wanted as a number.
			const defaultFormat = e.kind === "collection" || (e.kind === "note" && e.isList) ? "count" : "";
			set(composeExpr({ segs: e.segs, format: defaultFormat, fallback: parts?.fallback ?? "" }));
		};
		if (e.navigable) {
			const go = row.createEl("button", { cls: "xte-field-go", attr: { "aria-label": `Ver campos de ${e.label}`, title: `Ver campos de ${e.label}` } });
			go.createSpan({ text: "ver campos" });
			setIcon(go.createSpan(), "chevron-right");
			go.onclick = () => {
				const source = this.sourceId;
				const base = cellSources(this.ctx((this.selection as { row: number }).row) as CellContext).find((x) => x.id === source)?.base ?? [];
				this.trail = e.segs.slice(base.length);
				this.query = "";
				this.renderPanel(panel);
				panel.querySelector(".xte-fieldpicker")?.scrollIntoView({ block: "start" });
			};
		}
	}
}

function sameSegs(a: string[], b: string[]): boolean {
	return a.length === b.length && a.every((x, i) => x.toLowerCase() === b[i].toLowerCase());
}

function parsePathSafe(src: string): string[] {
	try {
		return parsePath(src);
	} catch {
		return [];
	}
}

