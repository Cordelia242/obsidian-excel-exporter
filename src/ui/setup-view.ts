import { ItemView, Notice, parseYaml, stringifyYaml, TFile, type ViewStateResult, type WorkspaceLeaf } from "obsidian";
import { ObsidianVaultAdapter } from "../adapter/obsidian-adapter";
import { colLetters, toBareExpression } from "../config/mapping";
import { extractCodeBlocks, orderDefinitionKeys, replaceCodeBlock } from "../config/parse";
import { compileDefinition, validateDefinition, type CompiledDefinition } from "../config/schema";
import { worksheetToGrid, type GridCell } from "../excel/grid";
import { loadWorkbook, type Workbook } from "../excel/workbook-io";
import { buildExport, prepareRun, type PreparedRun } from "../export/runner";
import { parsePath } from "../graph/path";
import { parseExpression } from "../values/pipes";
import {
	baseScope,
	collectionOptions,
	contextFields,
	isPerRootSheet,
	previewExpression,
	rowSampleScope,
	type FieldNode,
} from "../setup/fields";
import type ExcelTemplateExportPlugin from "../main";
import { errorText } from "./definition-picker";
import { PreviewModal } from "./preview-modal";
import { renderGrid, renderTabs, type RenderedGrid } from "./sheet-render";

export const SETUP_VIEW_TYPE = "excel-export-setup";

interface SetupState {
	file?: string;
	index?: number;
}

type Selection = { type: "cell"; row: number; col: number; address: string } | { type: "row"; row: number } | null;

const FORMATS: Array<{ pipe: string; label: string }> = [
	{ pipe: "", label: "Tal cual" },
	{ pipe: "count", label: "Contar elementos" },
	{ pipe: "stars", label: "Estrellas ⭐ → número" },
	{ pipe: 'date:"dd/MM/yyyy"', label: "Fecha como texto dd/MM/yyyy" },
	{ pipe: "first", label: "Solo el primero de la lista" },
	{ pipe: 'join:"; "', label: "Lista separada por ;" },
	{ pipe: "target", label: "Nombre de la nota (sin alias)" },
	{ pipe: "upper", label: "MAYÚSCULAS" },
	{ pipe: "lower", label: "minúsculas" },
	{ pipe: "link", label: "Hipervínculo a la nota" },
	{ pipe: "raw", label: "Crudo [[...]]" },
];

const KIND_ICON: Record<string, string> = {
	text: "Aa",
	number: "#",
	date: "📅",
	stars: "⭐",
	link: "🔗",
	list: "☰",
	collection: "▦",
	file: "📄",
	special: "@",
	group: "",
	value: "·",
};

/** Visual editor: shows the template and lets the user map cells and repeated rows to note data. */
export class SetupView extends ItemView {
	private state: SetupState = {};
	private file: TFile | null = null;
	private raw: Record<string, unknown> = {};
	private cd: CompiledDefinition | null = null;
	private run: PreparedRun | null = null;
	private wb: Workbook | null = null;
	private cells = new Map<string, string>();
	private rows = new Map<string, string>();
	private sheetIdx = 0;
	private sampleIdx = 0;
	private selection: Selection = null;
	private showSample = false;
	private grid: RenderedGrid | null = null;
	private gridEl: HTMLElement | null = null;
	private panelEl: HTMLElement | null = null;
	private statusEl: HTMLElement | null = null;
	private saveTimer: number | null = null;
	private fieldQuery = "";

	constructor(leaf: WorkspaceLeaf, private plugin: ExcelTemplateExportPlugin) {
		super(leaf);
	}

	getViewType(): string {
		return SETUP_VIEW_TYPE;
	}

	getDisplayText(): string {
		const name = typeof this.raw.name === "string" ? this.raw.name : this.file?.basename;
		return name ? `Setup: ${name}` : "Excel Export Setup";
	}

	getIcon(): string {
		return "table";
	}

	getState(): Record<string, unknown> {
		return { ...this.state };
	}

	async setState(state: unknown, result: ViewStateResult): Promise<void> {
		this.state = (state ?? {}) as SetupState;
		await this.loadDefinition();
		await super.setState(state, result);
	}

	async onClose(): Promise<void> {
		await this.flushSave();
	}

	// ---------------------------------------------------------------- loading

	private sheetNames(): string[] {
		return this.wb?.worksheets.map((w) => w.name) ?? [];
	}

	private currentSheet(): string {
		return this.sheetNames()[this.sheetIdx] ?? "";
	}

	private async loadDefinition(): Promise<void> {
		const root = this.contentEl;
		root.empty();
		root.addClass("xte-setup");
		const { file, index = 0 } = this.state;
		const f = file ? this.app.vault.getAbstractFileByPath(file) : null;
		if (!(f instanceof TFile)) {
			root.createEl("p", { text: "Abre el Setup desde el botón «Setup» de un bloque excel-export." });
			return;
		}
		this.file = f;
		try {
			const source = extractCodeBlocks(await this.app.vault.read(f))[index];
			if (source === undefined) throw new Error(`La nota ya no tiene el bloque excel-export #${index + 1}`);
			const raw = parseYaml(source);
			if (!raw || typeof raw !== "object" || Array.isArray(raw)) throw new Error("La definición debe ser un objeto YAML");
			this.raw = raw as Record<string, unknown>;
			this.cd = compileDefinition(validateDefinition(this.raw));
			this.run = prepareRun(this.cd, new ObsidianVaultAdapter(this.app), { settings: this.plugin.settings });
			await this.loadTemplate();
		} catch (e) {
			this.renderError(errorText(e));
			return;
		}
		this.renderAll();
		(this.leaf as unknown as { updateHeader?: () => void }).updateHeader?.();
	}

	private async loadTemplate(): Promise<void> {
		this.wb = null;
		const cd = this.cd;
		if (!cd) return;
		const path = cd.def.template;
		if (path && (await this.app.vault.adapter.exists(path))) {
			this.wb = await loadWorkbook(await this.app.vault.adapter.readBinary(path));
		}
		// Canonical keys "Hoja!B2" / "Hoja!7" (entries without sheet go to the first sheet).
		const first = this.sheetNames()[0] ?? "Hoja1";
		this.cells = new Map();
		this.rows = new Map();
		const defined = Object.values(cd.def.cells);
		cd.mapping.cells.forEach((m, i) => {
			const original = defined[i] ?? m.expr;
			this.cells.set(`${m.sheet ?? first}!${m.address}`, toBareExpression(String(original)));
		});
		for (const r of cd.mapping.rows) this.rows.set(`${r.sheet ?? first}!${r.row}`, r.src);
		if (this.sheetIdx >= this.sheetNames().length) this.sheetIdx = 0;
	}

	private renderError(message: string): void {
		const root = this.contentEl;
		root.empty();
		const box = root.createDiv({ cls: "xte-block xte-block-error" });
		box.createEl("strong", { text: "No se puede abrir el Setup" });
		box.createEl("pre", { text: message });
		const btns = box.createDiv({ cls: "xte-block-buttons" });
		if (this.file) {
			const file = this.file;
			btns.createEl("button", { text: "Abrir la nota" }).onclick = () => void this.app.workspace.getLeaf(false).openFile(file);
		}
		btns.createEl("button", { text: "Reintentar" }).onclick = () => void this.loadDefinition();
	}

	// ---------------------------------------------------------------- rendering

	private renderAll(): void {
		const root = this.contentEl;
		root.empty();
		const cd = this.cd;
		const run = this.run;
		if (!cd || !run) return;

		const bar = root.createDiv({ cls: "xte-setup-bar" });
		bar.createEl("strong", { text: cd.def.name });

		const tplLabel = bar.createEl("label", { text: "Template " });
		const tplSelect = tplLabel.createEl("select", { cls: "dropdown" });
		const xlsx = this.app.vault.getFiles().filter((f) => f.extension.toLowerCase() === "xlsx").map((f) => f.path).sort();
		if (!xlsx.includes(cd.def.template)) tplSelect.createEl("option", { text: cd.def.template || "(elige un .xlsx)", value: cd.def.template });
		for (const p of xlsx) tplSelect.createEl("option", { text: p, value: p });
		tplSelect.value = cd.def.template;
		tplSelect.onchange = () => void this.changeTemplate(tplSelect.value);

		if (run.roots.length) {
			const sLabel = bar.createEl("label", { text: "Nota de muestra " });
			const sSelect = sLabel.createEl("select", { cls: "dropdown" });
			run.roots.slice(0, 200).forEach((r, i) => sSelect.createEl("option", { text: r.basename, value: String(i) }));
			sSelect.value = String(this.sampleIdx);
			sSelect.onchange = () => {
				this.sampleIdx = Number(sSelect.value);
				this.renderGridArea();
				this.renderPanel();
			};
		}

		const toggle = bar.createEl("label", { cls: "xte-toggle" });
		const cb = toggle.createEl("input", { type: "checkbox" });
		cb.checked = this.showSample;
		toggle.appendText(" Ver datos de muestra");
		cb.onchange = () => {
			this.showSample = cb.checked;
			this.renderGridArea();
		};

		const actions = bar.createDiv({ cls: "xte-setup-actions" });
		this.statusEl = actions.createSpan({ cls: "xte-muted" });
		actions.createEl("button", { text: "Previsualizar" }).onclick = () => void this.preview();
		actions.createEl("button", { text: "Exportar", cls: "mod-cta" }).onclick = () => void this.exportNow();

		if (!this.wb) {
			root.createDiv({ cls: "xte-setup-empty" }).createEl("p", {
				text: `No se encontró el template "${cd.def.template}". Elige un archivo .xlsx del vault en el selector de arriba.`,
			});
			return;
		}

		const body = root.createDiv({ cls: "xte-setup-body" });
		const left = body.createDiv({ cls: "xte-setup-left" });
		renderTabs(left, this.sheetNames(), this.sheetIdx, (i) => {
			this.sheetIdx = i;
			this.selection = null;
			this.renderAll();
		});
		this.gridEl = left.createDiv({ cls: "xte-setup-grid" });
		this.panelEl = body.createDiv({ cls: "xte-setup-panel" });
		this.renderGridArea();
		this.renderPanel();
	}

	private perRoot(): boolean {
		return this.cd ? isPerRootSheet(this.cd.def.mode, this.sheetIdx) : true;
	}

	private rowPath(row: number): string[] | undefined {
		const src = this.rows.get(`${this.currentSheet()}!${row}`);
		if (!src) return undefined;
		try {
			return parsePath(src);
		} catch {
			return undefined;
		}
	}

	private sampleValue(expr: string, row: number): string {
		const run = this.run;
		if (!run) return "";
		const base = baseScope(run, this.perRoot(), this.sampleIdx);
		if (!base) return "(sin notas de muestra)";
		const scope = rowSampleScope(base, this.rowPath(row), run.rt);
		if (!scope) return "(la colección está vacía en la muestra)";
		try {
			return previewExpression(expr, scope, run);
		} catch (e) {
			return `⚠ ${(e as Error).message}`;
		}
	}

	private renderGridArea(): void {
		const el = this.gridEl;
		const ws = this.wb?.worksheets[this.sheetIdx];
		if (!el || !ws) return;
		el.empty();
		const sheet = ws.name;
		const grid = worksheetToGrid(ws, { minRows: Math.max(ws.rowCount + 10, 30), minCols: Math.max(ws.columnCount + 3, 12) });
		this.grid = renderGrid(el, grid, {
			cellContent: (cell: GridCell) => {
				const expr = this.cells.get(`${sheet}!${cell.address}`);
				if (!expr) return null;
				if (this.showSample) {
					const v = this.sampleValue(expr, cell.row);
					return { text: v, cls: "xte-mapped-sample", title: `${expr}\n→ ${v}` };
				}
				return { text: expr, cls: "xte-mapped", title: expr };
			},
			onCellClick: (cell) => this.select({ type: "cell", row: cell.row, col: cell.col, address: cell.address }),
			onRowClick: (row) => this.select({ type: "row", row }),
			rowClass: (row) => (this.rows.has(`${sheet}!${row}`) ? "xte-row-repeat" : undefined),
			rowTitle: (row) => {
				const src = this.rows.get(`${sheet}!${row}`);
				return src ? `Se repite por cada elemento de "${src}"` : "Clic para repetir esta fila por una colección";
			},
		});
		for (const [row, tr] of this.grid.rows) {
			const src = this.rows.get(`${sheet}!${row}`);
			if (src) tr.querySelector("th")?.setText(`↻ ${row}`);
		}
		this.highlightSelection();
	}

	private highlightSelection(): void {
		if (!this.grid) return;
		for (const td of this.grid.cells.values()) td.removeClass("is-selected");
		for (const tr of this.grid.rows.values()) tr.removeClass("is-selected");
		const sel = this.selection;
		if (sel?.type === "cell") this.grid.cells.get(sel.address)?.addClass("is-selected");
		if (sel?.type === "row") this.grid.rows.get(sel.row)?.addClass("is-selected");
	}

	private select(sel: Selection): void {
		this.selection = sel;
		this.fieldQuery = "";
		this.highlightSelection();
		this.renderPanel();
	}

	// ---------------------------------------------------------------- side panel

	private renderPanel(): void {
		const panel = this.panelEl;
		if (!panel) return;
		panel.empty();
		const sel = this.selection;
		if (!sel) {
			panel.createEl("h4", { text: "Setup" });
			panel.createEl("p", { text: "Haz clic en una celda para elegir con qué dato se llena." });
			panel.createEl("p", { text: "Haz clic en el número de una fila para repetirla por cada elemento de una colección (por ejemplo, una fila por interview)." });
			const n = [...this.cells.keys()].filter((k) => k.startsWith(`${this.currentSheet()}!`)).length;
			panel.createEl("p", { cls: "xte-muted", text: `${n} ${n === 1 ? "celda configurada" : "celdas configuradas"} en esta hoja.` });
			return;
		}
		if (sel.type === "row") this.renderRowPanel(panel, sel.row);
		else this.renderCellPanel(panel, sel.row, sel.col, sel.address);
	}

	private renderRowPanel(panel: HTMLElement, row: number): void {
		const run = this.run;
		if (!run) return;
		const key = `${this.currentSheet()}!${row}`;
		panel.createEl("h4", { text: `Fila ${row}` });
		panel.createEl("p", {
			cls: "xte-muted",
			text: "Una fila repetida se duplica una vez por elemento, copiando su formato. Las filas de abajo se desplazan y las fórmulas se ajustan.",
		});
		const label = panel.createEl("label", { text: "Repetir esta fila por: " });
		const select = label.createEl("select", { cls: "dropdown" });
		select.createEl("option", { text: "(no repetir)", value: "" });
		const options = collectionOptions(run, this.perRoot());
		const current = this.rows.get(key) ?? "";
		for (const o of options) select.createEl("option", { text: o.label, value: o.value });
		if (current && !options.some((o) => o.value === current)) select.createEl("option", { text: current, value: current });
		select.value = current;
		select.onchange = () => {
			if (select.value) this.rows.set(key, select.value);
			else this.rows.delete(key);
			this.changed();
			this.renderGridArea();
			this.renderPanel();
		};
		if (current) {
			const base = baseScope(run, this.perRoot(), this.sampleIdx);
			const count = base ? rowSampleScope(base, this.rowPath(row), run.rt) : null;
			const n = base ? (count ? (count.specials.count as number) : 0) : 0;
			panel.createEl("p", { text: `En la nota de muestra: ${n} ${n === 1 ? "fila" : "filas"}.` });
			panel.createEl("p", {
				cls: "xte-muted",
				text: `Ahora haz clic en las celdas de esta fila y elige campos dentro de «Fila: cada ${this.rowPath(row)?.slice(-1)[0] ?? current}».`,
			});
		}
		if (!options.length) panel.createEl("p", { cls: "xte-muted", text: "La definición no tiene relaciones. Agrégalas en el bloque YAML (relations)." });
	}

	private renderCellPanel(panel: HTMLElement, row: number, col: number, address: string): void {
		const run = this.run;
		if (!run) return;
		const sheet = this.currentSheet();
		const key = `${sheet}!${address}`;
		const rowPath = this.rowPath(row);
		const expr = this.cells.get(key) ?? "";

		panel.createEl("h4", { text: `Celda ${colLetters(col)}${row}` });
		if (rowPath) panel.createDiv({ cls: "xte-pill", text: `↻ fila repetida por ${rowPath.join(".")}` });
		const tplText = this.wb?.worksheets[this.sheetIdx]?.getCell(address).text;
		if (tplText && !this.cells.has(key)) {
			panel.createEl("p", { cls: "xte-muted", text: `Texto en el template: «${tplText}». Si eliges un campo, se reemplaza.` });
		}

		// Expression + format
		const exprBox = panel.createDiv({ cls: "xte-expr" });
		exprBox.createEl("label", { text: "Valor" });
		const input = exprBox.createEl("input", { type: "text", value: expr, placeholder: "Elige un campo abajo o escribe, ej. proceso.role" });
		const isSingle = (e: string) => e.trim() !== "" && !e.includes("{{");
		const split = (e: string) => {
			const p = parseExpression(e);
			return { path: p.path, pipe: e.includes("|") ? e.slice(e.indexOf("|") + 1).trim() : "" };
		};

		const fmtLabel = exprBox.createEl("label", { text: "Formato" });
		const fmt = fmtLabel.createEl("select", { cls: "dropdown" });
		for (const f of FORMATS) fmt.createEl("option", { text: f.label, value: f.pipe });
		const custom = fmt.createEl("option", { text: "Personalizado", value: "__custom" });
		const syncFmt = () => {
			const e = input.value;
			fmt.disabled = !isSingle(e);
			const pipe = isSingle(e) ? split(e).pipe : "";
			fmt.value = FORMATS.some((f) => f.pipe === pipe) ? pipe : "__custom";
			custom.hidden = fmt.value !== "__custom";
		};
		syncFmt();

		const previewEl = exprBox.createDiv({ cls: "xte-preview-value" });
		const updatePreview = () => {
			const e = input.value.trim();
			previewEl.empty();
			if (!e) return;
			previewEl.createSpan({ cls: "xte-muted", text: "Muestra: " });
			previewEl.createSpan({ text: this.sampleValue(e, row) || "(vacío)" });
		};
		updatePreview();

		const setExpr = (e: string) => {
			input.value = e;
			if (e.trim()) this.cells.set(key, e.trim());
			else this.cells.delete(key);
			syncFmt();
			updatePreview();
			this.grid?.refreshCell(address);
			this.changed();
		};
		input.oninput = () => setExpr(input.value);
		fmt.onchange = () => {
			if (fmt.value === "__custom" || !isSingle(input.value)) return;
			const { path } = split(input.value);
			setExpr(fmt.value ? `${path} | ${fmt.value}` : path);
		};

		const btns = exprBox.createDiv({ cls: "xte-block-buttons" });
		btns.createEl("button", { text: "Vaciar celda" }).onclick = () => setExpr("");
		btns.createEl("button", { text: "Insertar campo en texto" }).onclick = () => {
			// Switches to mixed mode: "Texto {{campo}}".
			const e = input.value.trim();
			setExpr(isSingle(e) ? `{{${e}}}` : e || "Texto {{ }}");
			input.focus();
		};
		exprBox.createDiv({
			cls: "xte-muted",
			text: "Para mezclar texto y datos escribe, por ejemplo: Rol: {{proceso.role}}",
		});

		// Field tree
		panel.createEl("h5", { text: "Campos disponibles" });
		const search = panel.createEl("input", { type: "search", placeholder: "Buscar campo…", value: this.fieldQuery, cls: "xte-field-search" });
		const treeEl = panel.createDiv({ cls: "xte-field-tree" });
		const pick = (node: FieldNode) => {
			const current = input.value.trim();
			if (current.includes("{{") && current.includes("{{ }}")) {
				setExpr(current.replace("{{ }}", `{{${node.expr}}}`));
				return;
			}
			const suggested = node.kind === "collection" ? "count" : node.kind === "stars" ? "stars" : "";
			setExpr(suggested ? `${node.expr} | ${suggested}` : node.expr);
		};
		const groups = contextFields({ run, perRoot: this.perRoot(), rowPath });
		const draw = () => {
			treeEl.empty();
			this.renderFieldNodes(treeEl, groups, pick, this.fieldQuery.trim().toLowerCase(), 0);
		};
		search.oninput = () => {
			this.fieldQuery = search.value;
			draw();
		};
		draw();
	}

	private renderFieldNodes(parent: HTMLElement, nodes: FieldNode[], pick: (n: FieldNode) => void, query: string, depth: number): number {
		let shown = 0;
		for (const node of nodes) {
			const matches = !query || node.label.toLowerCase().includes(query) || node.expr.toLowerCase().includes(query);
			const item = parent.createDiv({ cls: "xte-field" });
			const line = item.createDiv({ cls: "xte-field-line" });
			const hasChildren = !!node.children;
			const caret = line.createSpan({ cls: "xte-caret", text: hasChildren ? "▸" : "" });
			line.createSpan({ cls: "xte-field-icon", text: KIND_ICON[node.kind] ?? "" });
			const label = line.createSpan({ cls: node.kind === "group" ? "xte-field-group" : "xte-field-label", text: node.label });
			if (node.sample) line.createSpan({ cls: "xte-field-sample", text: node.sample });
			const childrenEl = item.createDiv({ cls: "xte-field-children" });
			let expanded = false;
			const expand = (open: boolean, q = "") => {
				expanded = open;
				caret.setText(hasChildren ? (open ? "▾" : "▸") : "");
				childrenEl.empty();
				if (open && node.children) return this.renderFieldNodes(childrenEl, node.children(), pick, q, depth + 1);
				return 0;
			};
			if (hasChildren) caret.onclick = () => void expand(!expanded);
			if (node.kind === "group") label.onclick = () => void expand(!expanded);
			else {
				label.onclick = () => pick(node);
				label.title = `Usar ${node.expr}`;
			}
			// Groups open by default at the top level; when searching, open up to two levels deep.
			let childMatches = 0;
			if (hasChildren && (depth === 0 || (query && depth < 2))) childMatches = expand(true, matches && query ? "" : query);
			if (query && !matches && childMatches === 0) item.remove();
			else shown++;
		}
		return shown;
	}

	// ---------------------------------------------------------------- persistence

	private changed(): void {
		this.statusEl?.setText("Guardando…");
		if (this.saveTimer !== null) window.clearTimeout(this.saveTimer);
		this.saveTimer = window.setTimeout(() => void this.flushSave(), 700);
	}

	private mappingObjects(): { cells: Record<string, string>; rows: Record<string, string> } {
		const order = this.sheetNames();
		const sortKey = (k: string) => {
			const i = k.lastIndexOf("!");
			const sheet = k.slice(0, i);
			const addr = k.slice(i + 1);
			const m = /^([A-Z]+)(\d+)$/.exec(addr);
			const row = m ? Number(m[2]) : Number(addr);
			const col = m ? m[1].padStart(3, " ") : "";
			return `${String(order.indexOf(sheet)).padStart(3, "0")}|${String(row).padStart(7, "0")}|${col}`;
		};
		const sorted = (map: Map<string, string>) =>
			Object.fromEntries([...map.entries()].sort((a, b) => sortKey(a[0]).localeCompare(sortKey(b[0]))));
		return { cells: sorted(this.cells), rows: sorted(this.rows) };
	}

	/** The definition with the current Setup mapping (unsaved changes included). */
	private currentRaw(): Record<string, unknown> {
		const { cells, rows } = this.mappingObjects();
		const raw: Record<string, unknown> = { ...this.raw };
		if (Object.keys(cells).length) raw.cells = cells;
		else delete raw.cells;
		if (Object.keys(rows).length) raw.rows = rows;
		else delete raw.rows;
		return orderDefinitionKeys(raw);
	}

	private async flushSave(): Promise<void> {
		if (this.saveTimer === null || !this.file) return;
		window.clearTimeout(this.saveTimer);
		this.saveTimer = null;
		const raw = this.currentRaw();
		try {
			const cd = compileDefinition(validateDefinition(raw));
			const yaml = stringifyYaml(raw);
			await this.app.vault.process(this.file, (md) => replaceCodeBlock(md, this.state.index ?? 0, yaml));
			this.raw = raw;
			this.cd = cd;
			if (this.run) this.run.cd = cd;
			this.statusEl?.setText("Guardado ✓");
		} catch (e) {
			// Usually an expression still being typed: keep editing, save on the next valid change.
			const first = errorText(e).split("\n")[0];
			this.statusEl?.setText(`⚠ Sin guardar: ${first}`);
		}
	}

	private async changeTemplate(path: string): Promise<void> {
		this.raw = { ...this.raw, template: path };
		const keep = { cells: this.cells, rows: this.rows };
		try {
			this.cd = compileDefinition(validateDefinition(this.currentRaw()));
			await this.loadTemplate();
			// Keep the mapping the user already made (same keys if sheets match).
			this.cells = keep.cells;
			this.rows = keep.rows;
		} catch (e) {
			new Notice(errorText(e));
		}
		this.selection = null;
		this.renderAll();
		this.changed();
	}

	private async buildCurrent(sampleOnly: boolean) {
		await this.flushSave();
		const cd = compileDefinition(validateDefinition(this.currentRaw()));
		const sample = this.run?.roots[this.sampleIdx];
		return {
			cd,
			built: await buildExport(cd, new ObsidianVaultAdapter(this.app), {
				settings: this.plugin.settings,
				activeNotePath: sampleOnly && cd.def.mode !== "single" && sample ? sample.path : undefined,
			}),
		};
	}

	private async preview(): Promise<void> {
		try {
			const { cd, built } = await this.buildCurrent(true);
			new PreviewModal(this.app, `Vista previa — ${cd.def.name}`, built).open();
		} catch (e) {
			new Notice(`Excel Export: ${errorText(e)}`, 8000);
		}
	}

	private async exportNow(): Promise<void> {
		await this.flushSave();
		try {
			const cd = compileDefinition(validateDefinition(this.currentRaw()));
			await this.plugin.runDefinition(cd);
		} catch (e) {
			new Notice(`Excel Export: ${errorText(e)}`, 8000);
		}
	}
}
