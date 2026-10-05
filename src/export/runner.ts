import type { VaultAdapter } from "../adapter/vault-adapter";
import { DEFAULT_SETTINGS, type ExportSettings } from "../config/defaults";
import type { CompiledDefinition } from "../config/schema";
import { makeScope, Runtime, type Scope } from "../graph/context";
import { RelationBuilder } from "../graph/relations";
import type { NoteRecord } from "../model/note-record";
import { Report, type ExportWarning } from "../model/report";
import { evaluate } from "../query/evaluator";
import { sortNotes } from "../query/sort";
import type { NormalizeOptions } from "../values/normalize";
import { cloneWorksheet, fillWorksheet, sanitizeSheetName, type FillOptions } from "../excel/engine";
import { renderString } from "../excel/placeholders";
import { applyCellMapping, rowRepeatsFor } from "../excel/mapping";
import { loadWorkbook, saveWorkbook, unshareFormulas, type Workbook, type Worksheet } from "../excel/workbook-io";

export type OverwriteAnswer = "overwrite" | "suffix" | "skip";

export interface RunOptions {
	settings?: Partial<ExportSettings>;
	/** Export only this note as root (command "Exportar nota activa"): `where` is checked, `filter` ignored. */
	activeNotePath?: string;
	now?: Date;
	/** Called when `overwrite: ask` and the output exists. Defaults to "suffix". */
	confirmOverwrite?: (path: string) => Promise<OverwriteAnswer>;
}

export interface ExportResult {
	files: string[];
	roots: number;
	warnings: ExportWarning[];
}

/** Everything prepared for a run: runtime, roots with their relations, options. */
export interface PreparedRun {
	cd: CompiledDefinition;
	rt: Runtime;
	roots: NoteRecord[];
	settings: ExportSettings;
	normalize: NormalizeOptions;
}

export function normalizeOptions(cd: CompiledDefinition, settings: ExportSettings): NormalizeOptions {
	return {
		links: cd.def.normalize?.links ?? settings.links,
		listSeparator: cd.def.normalize?.listSeparator ?? settings.listSeparator,
		stars: cd.def.normalize?.stars ?? settings.stars,
		emptyValue: cd.def.normalize?.emptyValue ?? settings.emptyValue,
	};
}

/** Selects roots (where → filter → sort) and computes their relations. */
export function prepareRun(cd: CompiledDefinition, adapter: VaultAdapter, options: RunOptions = {}): PreparedRun {
	const settings = { ...DEFAULT_SETTINGS, ...options.settings };
	const report = new Report();
	const rt = new Runtime(adapter, report, options.now ?? new Date(), cd.def.name);
	const notes = adapter.listNotes();

	let roots: NoteRecord[];
	if (options.activeNotePath) {
		const note = notes.find((n) => n.path === options.activeNotePath);
		if (!note) throw new Error(`No se encontró la nota "${options.activeNotePath}"`);
		if (!evaluate(cd.where, note, rt)) {
			report.warn("skipped-note", `La nota no cumple root.where (${cd.def.root.where}); se exporta igual`, {
				note: note.path,
			});
		}
		roots = [note];
	} else {
		roots = notes.filter((n) => evaluate(cd.where, n, rt));
		if (cd.filter) {
			const f = cd.filter;
			roots = roots.filter((n) => evaluate(f, n, rt));
		}
		roots = sortNotes(roots, cd.sort, rt);
	}

	const builder = new RelationBuilder(cd.relations, notes, cd.def.root.alias, rt);
	for (const root of roots) builder.build(root);
	return { cd, rt, roots, settings, normalize: normalizeOptions(cd, settings) };
}

/** Counts the roots a definition would export (code block preview). */
export function countRoots(cd: CompiledDefinition, adapter: VaultAdapter): number {
	const rt = new Runtime(adapter, new Report());
	let roots = adapter.listNotes().filter((n) => evaluate(cd.where, n, rt));
	if (cd.filter) {
		const f = cd.filter;
		roots = roots.filter((n) => evaluate(f, n, rt));
	}
	return roots.length;
}

function baseSpecials(run: PreparedRun): Record<string, unknown> {
	return { today: run.rt.today, now: run.rt.now, exportName: run.cd.def.name };
}

export function rootScope(run: PreparedRun, root: NoteRecord, index: number): Scope {
	const rels = run.rt.relationsOf(root) ?? new Map();
	return makeScope(
		{ ...Object.fromEntries(rels), [run.cd.def.root.alias]: root },
		{ ...baseSpecials(run), index: index + 1, count: run.roots.length },
	);
}

export function collectionScope(run: PreparedRun): Scope {
	return makeScope({ [run.cd.def.root.alias]: run.roots }, { ...baseSpecials(run), count: run.roots.length });
}

const INVALID_FILE_CHARS = /[\\/:*?"<>|]/g;

export function sanitizeFileName(name: string): string {
	let base = name.replace(INVALID_FILE_CHARS, "-").replace(/\s+/g, " ").trim();
	if (!base || base === ".xlsx") base = "export.xlsx";
	if (!/\.xlsx$/i.test(base)) base += ".xlsx";
	return base;
}

function joinPath(folder: string, file: string): string {
	const f = folder.replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
	return f ? `${f}/${file}` : file;
}

function withSuffix(path: string, n: number): string {
	return path.replace(/(\.xlsx)?$/i, ` (${n})$1`);
}

class OutputWriter {
	private used = new Set<string>();
	readonly files: string[] = [];

	constructor(private adapter: VaultAdapter, private run: PreparedRun, private options: RunOptions) {}

	async write(fileName: string, data: ArrayBuffer): Promise<void> {
		const folder = this.run.cd.def.output.folder ?? this.run.settings.outputFolder;
		let path = joinPath(folder, sanitizeFileName(fileName));
		if (this.used.has(path.toLowerCase())) {
			this.run.rt.report.warn("other", `Nombre de archivo repetido "${path}"; se agrega un sufijo`);
			path = await this.freePath(path);
		}
		let overwrite = false;
		if (await this.adapter.exists(path)) {
			const mode = this.run.cd.def.output.overwrite ?? this.run.settings.overwrite;
			let answer: OverwriteAnswer = mode === "overwrite" ? "overwrite" : "suffix";
			if (mode === "ask") answer = this.options.confirmOverwrite ? await this.options.confirmOverwrite(path) : "suffix";
			if (answer === "skip") {
				this.run.rt.report.warn("other", `Se omitió "${path}" porque ya existe`);
				return;
			}
			if (answer === "overwrite") overwrite = true;
			else path = await this.freePath(path);
		}
		await this.adapter.writeBinary(path, data, overwrite);
		this.used.add(path.toLowerCase());
		this.files.push(path);
	}

	private async freePath(path: string): Promise<string> {
		for (let n = 2; ; n++) {
			const candidate = withSuffix(path, n);
			if (!this.used.has(candidate.toLowerCase()) && !(await this.adapter.exists(candidate))) return candidate;
		}
	}
}

function fillOptions(run: PreparedRun): FillOptions {
	return { rt: run.rt, opts: run.normalize, emptyBlock: run.cd.def.emptyBlock ?? run.settings.emptyBlock };
}

function renderName(run: PreparedRun, src: string, scope: Scope, what: string): string {
	return renderString(src, scope, { rt: run.rt, opts: { ...run.normalize, emptyValue: "" }, location: what });
}

/** Removes defined names (e.g. print areas) that only point at a deleted sheet. */
function dropDefinedNamesFor(wb: Workbook, sheetName: string): void {
	try {
		const dn = wb.definedNames as unknown as { model: Array<{ name: string; ranges: string[] }> };
		const prefix = (r: string) => r.split("!")[0].replace(/^'|'$/g, "").replace(/''/g, "'");
		dn.model = dn.model
			.map((n) => ({ ...n, ranges: n.ranges.filter((r) => prefix(r) !== sheetName) }))
			.filter((n) => n.ranges.length);
	} catch {
		// Best effort: ExcelJS internals.
	}
}

export async function loadTemplate(adapter: VaultAdapter, path: string): Promise<ArrayBuffer> {
	if (!(await adapter.exists(path))) throw new Error(`No se encontró el template "${path}"`);
	return adapter.readBinary(path);
}

export interface BuiltOutput {
	/** File name (not yet sanitized/deduplicated against the vault). */
	fileName: string;
	workbook: Workbook;
}

/** Workbooks generated in memory, ready to preview and then write. */
export interface BuiltExport {
	run: PreparedRun;
	outputs: BuiltOutput[];
	warnings: ExportWarning[];
}

/** Loads the template and applies the Setup cell mapping. */
export async function loadPreparedTemplate(data: ArrayBuffer, cd: CompiledDefinition, run: PreparedRun): Promise<Workbook> {
	const wb = await loadWorkbook(data.slice(0));
	unshareFormulas(wb);
	applyCellMapping(wb, cd.mapping, run.rt.report);
	return wb;
}

/** Generates every output workbook in memory (§6.4) without writing anything. */
export async function buildExport(cd: CompiledDefinition, adapter: VaultAdapter, options: RunOptions = {}): Promise<BuiltExport> {
	const template = await loadTemplate(adapter, cd.def.template);
	const run = prepareRun(cd, adapter, options);
	const { def } = cd;
	const fo = fillOptions(run);
	const outputs: BuiltOutput[] = [];

	if (run.roots.length === 0) {
		run.rt.report.warn("other", "Ninguna nota cumple el filtro de la raíz; no se generó ningún archivo");
		return { run, outputs, warnings: run.rt.report.warnings };
	}

	const repeats = (wb: Workbook, ws: Worksheet) => rowRepeatsFor(cd.mapping, ws, wb.worksheets[0]?.name ?? ws.name);

	if (def.mode === "file-per-root") {
		const filename = def.output.filename ?? `{{${def.root.alias}.file.name}}.xlsx`;
		for (let i = 0; i < run.roots.length; i++) {
			const wb = await loadPreparedTemplate(template, cd, run);
			const scope = rootScope(run, run.roots[i], i);
			for (const ws of [...wb.worksheets]) fillWorksheet(wb, ws, scope, fo, repeats(wb, ws));
			outputs.push({ fileName: renderName(run, filename, scope, "output.filename"), workbook: wb });
		}
	} else if (def.mode === "sheet-per-root") {
		const wb = await loadPreparedTemplate(template, cd, run);
		const [tplSheet, ...others] = wb.worksheets;
		const tplRepeats = repeats(wb, tplSheet);
		const otherRepeats = new Map(others.map((ws) => [ws, repeats(wb, ws)]));
		const taken = new Set(others.map((ws) => ws.name.toLowerCase()));
		const sheetName = def.output.sheetName ?? `{{${def.root.alias}.file.name}}`;
		const clones: Array<{ ws: Worksheet; name: string }> = [];
		for (let i = 0; i < run.roots.length; i++) {
			const scope = rootScope(run, run.roots[i], i);
			const name = sanitizeSheetName(renderName(run, sheetName, scope, "output.sheetName"), taken);
			const ws = cloneWorksheet(wb, tplSheet, `__xte_${i}`, run.rt);
			fillWorksheet(wb, ws, scope, fo, tplRepeats);
			clones.push({ ws, name });
		}
		const collection = collectionScope(run);
		for (const ws of others) fillWorksheet(wb, ws, collection, fo, otherRepeats.get(ws));
		const order = (ws: Worksheet) => (ws as unknown as { orderNo: number }).orderNo;
		const tplName = tplSheet.name;
		const tplOrder = order(tplSheet);
		wb.removeWorksheet(tplSheet.id);
		dropDefinedNamesFor(wb, tplName);
		for (const c of clones) c.ws.name = c.name;
		// Clones take the template's position; the remaining sheets keep their relative order.
		const ordered = [
			...others.filter((ws) => order(ws) < tplOrder),
			...clones.map((c) => c.ws),
			...others.filter((ws) => order(ws) > tplOrder),
		];
		ordered.forEach((ws, i) => ((ws as unknown as { orderNo: number }).orderNo = i + 1));
		const filename = def.output.filename ?? "{{@exportName}}.xlsx";
		outputs.push({ fileName: renderName(run, filename, collection, "output.filename"), workbook: wb });
	} else {
		const wb = await loadPreparedTemplate(template, cd, run);
		const scope = collectionScope(run);
		for (const ws of [...wb.worksheets]) fillWorksheet(wb, ws, scope, fo, repeats(wb, ws));
		const filename = def.output.filename ?? "{{@exportName}}.xlsx";
		outputs.push({ fileName: renderName(run, filename, scope, "output.filename"), workbook: wb });
	}

	return { run, outputs, warnings: run.rt.report.warnings };
}

/** Saves the built workbooks into the vault, honoring `overwrite`. */
export async function writeExport(built: BuiltExport, adapter: VaultAdapter, options: RunOptions = {}): Promise<ExportResult> {
	const writer = new OutputWriter(adapter, built.run, options);
	for (const out of built.outputs) await writer.write(out.fileName, await saveWorkbook(out.workbook));
	return { files: writer.files, roots: built.run.roots.length, warnings: built.run.rt.report.warnings };
}

/** Runs a definition end to end and writes the resulting files. */
export async function runExport(cd: CompiledDefinition, adapter: VaultAdapter, options: RunOptions = {}): Promise<ExportResult> {
	return writeExport(await buildExport(cd, adapter, options), adapter, options);
}
