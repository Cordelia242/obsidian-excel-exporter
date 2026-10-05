import { toTemplateText } from "../config/mapping";
import type { ExportDefinition } from "../config/schema";
import { iterateCollection, parseCellText, renderTemplate } from "../excel/placeholders";
import { collectionScope, rootScope, type PreparedRun } from "../export/runner";
import type { Runtime, Scope } from "../graph/context";
import { parsePath } from "../graph/path";
import { isNoteRecord, type NoteRecord } from "../model/note-record";
import { LinkRef, wrapFrontmatterValue } from "../values/link";
import { isEmptyValue, scalarToText } from "../values/normalize";
import { parseExpression } from "../values/pipes";
import { propertiesOf, type PropKind } from "./discovery";

// ------------------------------------------------------------------ labels

export function capitalize(s: string): string {
	return s ? s[0].toUpperCase() + s.slice(1) : s;
}

export function rootLabel(def: ExportDefinition): string {
	return def.root.label || capitalize(def.root.alias);
}

export function relationLabel(def: ExportDefinition, name: string): string {
	return def.relations[name]?.label || capitalize(name.replace(/_/g, " "));
}

/** Friendly name of a path variable (alias, relation or property). */
export function variableLabel(def: ExportDefinition, name: string): string {
	if (name.toLowerCase() === def.root.alias.toLowerCase()) return rootLabel(def);
	const rel = Object.keys(def.relations).find((r) => r.toLowerCase() === name.toLowerCase());
	return rel ? relationLabel(def, rel) : capitalize(name);
}

const SPECIAL_LABELS: Record<string, string> = {
	"@today": "Fecha de hoy",
	"@now": "Fecha y hora actual",
	"@exportName": "Nombre del template",
	"@index": "Nº de fila",
	"@count": "Total",
};

const FILE_LABELS: Record<string, string> = {
	name: "Nombre de la nota",
	basename: "Nombre de la nota",
	path: "Ruta",
	folder: "Carpeta",
	ctime: "Fecha de creación",
	mtime: "Última modificación",
	link: "Link a la nota",
	ext: "Extensión",
};

const PIPE_LABELS: Record<string, string> = {
	count: "cantidad",
	first: "el primero",
	stars: "número",
	raw: "texto original",
	target: "nombre real",
	link: "hipervínculo",
	upper: "MAYÚSCULAS",
	lower: "minúsculas",
	date: "fecha como texto",
	join: "lista",
};

/** Joins path segments; segments with dots or brackets use `["..."]`. */
export function pathExpr(segs: string[]): string {
	let out = "";
	segs.forEach((s, i) => {
		if (i === 0) out = s;
		else if (/[.[\]"]/.test(s)) out += `["${s}"]`;
		else out += `.${s}`;
	});
	return out;
}

/** Human-readable label for a mapping expression: `proceso.hires | count` → `Proceso › Hires (cantidad)`. */
export function friendlyLabel(expr: string, def: ExportDefinition): string {
	const one = (src: string): string => {
		const { path, pipes } = parseExpression(src);
		let segs: string[];
		try {
			segs = parsePath(path);
		} catch {
			return src;
		}
		const parts: string[] = [];
		for (let i = 0; i < segs.length; i++) {
			const s = segs[i];
			if (i === 0 && s.startsWith("@")) parts.push(SPECIAL_LABELS[s] ?? s);
			else if (s === "file" && i + 1 < segs.length) {
				parts.push(FILE_LABELS[segs[i + 1]] ?? segs[i + 1]);
				i++;
			} else parts.push(i === 0 || def.relations[s] ? variableLabel(def, s) : s);
		}
		const fmt = pipes.filter((p) => p.name !== "default").map((p) => PIPE_LABELS[p.name] ?? p.name);
		return parts.join(" › ") + (fmt.length ? ` (${fmt.join(", ")})` : "");
	};
	if (!expr.includes("{{")) return one(expr);
	return expr.replace(/\{\{([^{}]*)\}\}/g, (_m, inner: string) => `[${one(inner.trim())}]`);
}

// ------------------------------------------------------------------ expression <-> parts

export interface ExprParts {
	segs: string[];
	/** Format pipe text (e.g. `count`, `date:"dd/MM/yyyy"`), "" when none. */
	format: string;
	/** Value shown when empty (`default` pipe), "" when none. */
	fallback: string;
}

function pipeText(name: string, arg?: string): string {
	return arg === undefined ? name : `${name}:"${arg.replace(/"/g, '\\"')}"`;
}

/** Splits a simple expression into path + one format + fallback. Null when it is more complex (mixed text…). */
export function parseExprParts(expr: string): ExprParts | null {
	if (!expr.trim() || expr.includes("{{")) return null;
	const { path, pipes } = parseExpression(expr);
	let segs: string[];
	try {
		segs = parsePath(path);
	} catch {
		return null;
	}
	const fallbackPipe = pipes.find((p) => p.name === "default");
	const formats = pipes.filter((p) => p.name !== "default");
	if (formats.length > 1) return null;
	return {
		segs,
		format: formats[0] ? pipeText(formats[0].name, formats[0].arg) : "",
		fallback: fallbackPipe?.arg ?? "",
	};
}

export function composeExpr(parts: ExprParts): string {
	let e = pathExpr(parts.segs);
	if (parts.format) e += ` | ${parts.format}`;
	if (parts.fallback) e += ` | ${pipeText("default", parts.fallback)}`;
	return e;
}

// ------------------------------------------------------------------ sources and fields

export type FieldKind = PropKind | "note" | "collection" | "file" | "special";

export interface Source {
	id: string;
	label: string;
	hint: string;
	base: string[];
	/** The source yields several values (a relation outside a repeated row, or all roots). */
	isList: boolean;
	notes: NoteRecord[];
	kind: "row" | "root" | "relation" | "special";
}

export interface FieldEntry {
	label: string;
	/** Full path (base + trail + key). */
	segs: string[];
	kind: FieldKind;
	sample: string;
	/** Can be opened to see the fields of the linked notes. */
	navigable: boolean;
	isList: boolean;
}

export interface CellContext {
	run: PreparedRun;
	perRoot: boolean;
	rowPath?: string[];
	sampleIndex?: number;
}

function uniqueNotes(notes: NoteRecord[]): NoteRecord[] {
	const seen = new Set<string>();
	return notes.filter((n) => (seen.has(n.path) ? false : (seen.add(n.path), true)));
}

/** The scope a sheet is filled with, using one root as sample. */
export function baseScope(run: PreparedRun, perRoot: boolean, sampleIndex = 0): Scope | null {
	if (!perRoot) return collectionScope(run);
	const root = run.roots[sampleIndex];
	return root ? rootScope(run, root, sampleIndex) : null;
}

/** Scope of the first element of a repeated row (or the base scope). */
export function rowSampleScope(base: Scope, rowPath: string[] | undefined, rt: Runtime): Scope | null {
	if (!rowPath) return base;
	return iterateCollection(rowPath, base, rt)?.[0] ?? null;
}

/** Notes bound to each variable of a repeated row, across all roots. */
function rowVariableNotes(ctx: CellContext): Map<string, NoteRecord[]> {
	const out = new Map<string, NoteRecord[]>();
	if (!ctx.rowPath) return out;
	const scopes: Scope[] = ctx.perRoot ? ctx.run.roots.slice(0, 200).map((r, i) => rootScope(ctx.run, r, i)) : [collectionScope(ctx.run)];
	for (const scope of scopes) {
		for (const item of iterateCollection(ctx.rowPath, scope, ctx.run.rt) ?? []) {
			for (const name of ctx.rowPath) {
				let v = item.vars[name];
				if (v instanceof LinkRef) v = ctx.run.rt.resolveLink(v, false) ?? undefined;
				if (isNoteRecord(v)) out.set(name, [...(out.get(name) ?? []), v]);
			}
		}
	}
	for (const [k, v] of out) out.set(k, uniqueNotes(v));
	return out;
}

/** Where a cell's data can come from, most relevant first. */
export function cellSources(ctx: CellContext): Source[] {
	const { run } = ctx;
	const def = run.cd.def;
	const alias = def.root.alias;
	const out: Source[] = [];
	const rowVars = rowVariableNotes(ctx);
	const rootBoundByRow = !!ctx.rowPath && ctx.rowPath[0].toLowerCase() === alias.toLowerCase();

	if (ctx.rowPath) {
		for (const name of [...ctx.rowPath].reverse()) {
			if (name.toLowerCase() === alias.toLowerCase()) continue;
			out.push({
				id: `row:${name}`,
				label: `${variableLabel(def, name)} de esta fila`,
				hint: "Cambia en cada fila repetida",
				base: [name],
				isList: false,
				notes: rowVars.get(name) ?? [],
				kind: "row",
			});
		}
	}

	const rootIsOne = ctx.perRoot || rootBoundByRow;
	out.push({
		id: "root",
		label: rootIsOne ? rootLabel(def) : `Todos los ${rootLabel(def)}`,
		hint: rootIsOne ? (rootBoundByRow ? "El de esta fila" : `El ${rootLabel(def).toLowerCase()} que se exporta`) : "Todas las notas exportadas",
		base: [alias],
		isList: !rootIsOne,
		notes: run.roots,
		kind: "root",
	});

	if (rootIsOne) {
		for (const rel of run.cd.relations) {
			const items = uniqueNotes(run.roots.flatMap((r) => run.rt.relationsOf(r)?.get(rel.name) ?? []));
			out.push({
				id: `rel:${rel.name}`,
				label: `${relationLabel(def, rel.name)} (todas)`,
				hint: `Todas las del ${rootLabel(def).toLowerCase()}: contar, listar…`,
				base: [alias, rel.name],
				isList: true,
				notes: items,
				kind: "relation",
			});
		}
	}

	out.push({ id: "special", label: "Fecha y otros", hint: "Fecha de hoy, número de fila…", base: [], isList: false, notes: [], kind: "special" });
	return out;
}

function atoms(v: unknown): unknown[] {
	return Array.isArray(v) ? v.flatMap(atoms) : [v];
}

/** Notes reached from `notes` by following `seg` (a relation or a link property). */
function followSegment(notes: NoteRecord[], seg: string, rt: Runtime): { notes: NoteRecord[]; list: boolean } {
	const out: NoteRecord[] = [];
	let list = false;
	for (const n of notes) {
		const rel = rt.relationsOf(n)?.get(seg);
		if (rel) {
			out.push(...rel);
			list = true;
			continue;
		}
		const key = Object.keys(n.frontmatter).find((k) => k.toLowerCase() === seg.toLowerCase());
		if (key === undefined) continue;
		const v = wrapFrontmatterValue(n.frontmatter[key], n.path);
		if (Array.isArray(v) && v.length > 1) list = true;
		for (const a of atoms(v)) {
			if (a instanceof LinkRef) {
				const t = rt.resolveLink(a, false);
				if (t) out.push(t);
			}
		}
	}
	return { notes: uniqueNotes(out), list };
}

/** Fields visible in the browser for a source and a navigation trail (segments after the base). */
export function fieldsAt(source: Source, trail: string[], ctx: CellContext): FieldEntry[] {
	const { run } = ctx;
	const rt = run.rt;
	if (source.kind === "special") {
		const items: Array<[string, FieldKind, string]> = [
			["@today", "date", scalarToText(rt.today)],
			["@now", "date", ""],
			["@exportName", "text", run.cd.def.name],
		];
		if (ctx.rowPath) items.unshift(["@index", "number", "1, 2, 3…"]);
		items.push(["@count", "number", ctx.rowPath ? "Total de filas" : "Total de notas exportadas"]);
		return items.map(([e, kind, sample]) => ({ label: SPECIAL_LABELS[e], segs: [e], kind, sample, navigable: false, isList: false }));
	}

	let notes = source.notes;
	let isList = source.isList;
	for (const seg of trail) {
		const next = followSegment(notes, seg, rt);
		notes = next.notes;
		isList = isList || next.list;
	}
	const prefix = [...source.base, ...trail];
	const out: FieldEntry[] = [];
	out.push({
		label: "Nombre de la nota",
		segs: prefix,
		kind: "note",
		sample: notes.slice(0, 3).map((n) => n.basename).join(", "),
		navigable: false,
		isList,
	});

	const relNames = new Set<string>();
	for (const n of notes.slice(0, 200)) for (const name of rt.relationsOf(n)?.keys() ?? []) relNames.add(name);
	for (const name of relNames) {
		if (source.kind === "relation" && trail.length === 0 && name === source.base[1]) continue;
		const first = notes[0] ? rt.relationsOf(notes[0])?.get(name)?.length ?? 0 : 0;
		out.push({
			label: relationLabel(run.cd.def, name),
			segs: [...prefix, name],
			kind: "collection",
			sample: `${first} en la muestra`,
			navigable: true,
			isList: true,
		});
	}

	for (const p of propertiesOf(notes)) {
		const segs = [...prefix, p.key];
		const list = isList || p.kind === "list";
		const linkish = p.kind === "link" || (p.kind === "list" && followSegment(notes.slice(0, 50), p.key, rt).notes.length > 0);
		out.push({ label: p.key, segs, kind: p.kind, sample: p.sample, navigable: linkish, isList: list });
	}

	const sample = notes[0] ? rt.fileInfo(notes[0]) : undefined;
	for (const f of ["ctime", "mtime", "folder", "link"] as const) {
		out.push({
			label: FILE_LABELS[f],
			segs: [...prefix, "file", f],
			kind: f === "ctime" || f === "mtime" ? "date" : "file",
			sample: sample ? scalarToText(sample[f]) : "",
			navigable: false,
			isList,
		});
	}
	return out;
}

/** Trail labels for the breadcrumb. */
export function trailLabels(source: Source, trail: string[], def: ExportDefinition): string[] {
	return [source.label, ...trail.map((t) => (def.relations[t] ? relationLabel(def, t) : t))];
}

// ------------------------------------------------------------------ formats

export interface FormatOption {
	pipe: string;
	label: string;
}

/** "How to show it" options that make sense for a field. */
export function formatOptions(kind: FieldKind, isList: boolean): FormatOption[] {
	if (isList || kind === "collection") {
		return [
			{ pipe: "", label: "Todos, separados por coma" },
			{ pipe: "count", label: "Cantidad (número)" },
			{ pipe: "first", label: "Solo el primero" },
			{ pipe: 'join:"; "', label: "Todos, separados por ;" },
		];
	}
	switch (kind) {
		case "note":
		case "link":
			return [
				{ pipe: "", label: "Nombre (como se ve en Obsidian)" },
				{ pipe: "target", label: "Nombre real de la nota" },
				{ pipe: "link", label: "Nombre con hipervínculo a Obsidian" },
			];
		case "date":
			return [
				{ pipe: "", label: "Fecha de Excel (usa el formato de la celda)" },
				{ pipe: 'date:"dd/MM/yyyy"', label: "Texto 31/12/2026" },
				{ pipe: 'date:"yyyy-MM-dd"', label: "Texto 2026-12-31" },
			];
		case "stars":
			return [
				{ pipe: "", label: "Número (⭐⭐⭐ → 3)" },
				{ pipe: "raw", label: "Estrellas tal cual" },
			];
		case "text":
			return [
				{ pipe: "", label: "Tal cual" },
				{ pipe: "upper", label: "MAYÚSCULAS" },
				{ pipe: "lower", label: "minúsculas" },
			];
		default:
			return [{ pipe: "", label: "Tal cual" }];
	}
}

// ------------------------------------------------------------------ repeat options + preview

export interface CollectionOption {
	value: string;
	label: string;
	count: number;
}

/** Collections a row can repeat over, with the count for the sample root. */
export function collectionOptions(run: PreparedRun, perRoot: boolean, sampleIndex = 0): CollectionOption[] {
	const def = run.cd.def;
	const { alias } = def.root;
	const base = baseScope(run, perRoot, sampleIndex);
	const count = (path: string) => (base ? iterateCollection(parsePath(path), base, run.rt)?.length ?? 0 : 0);
	const rels = run.cd.relations.map((r) => r.name);
	if (perRoot) return rels.map((r) => ({ value: r, label: relationLabel(def, r), count: count(r) }));
	return [
		{ value: alias, label: `${rootLabel(def)} (una fila por nota exportada)`, count: count(alias) },
		...rels.map((r) => ({
			value: `${alias}.${r}`,
			label: `${relationLabel(def, r)} de todos los ${rootLabel(def)}`,
			count: count(`${alias}.${r}`),
		})),
	];
}

/** Evaluates a mapping expression against a sample scope. */
export function previewExpression(expr: string, scope: Scope, run: PreparedRun): string {
	const tpl = parseCellText(toTemplateText(expr));
	if (!tpl) return expr;
	if (tpl.issues.length) return `⚠ ${tpl.issues[0]}`;
	const r = renderTemplate({ ...tpl, each: undefined }, scope, { rt: run.rt, opts: run.normalize, location: "setup", warnUnresolved: false });
	return scalarToText(r.value);
}

export function isEmptyPreview(text: string): boolean {
	return isEmptyValue(text);
}
