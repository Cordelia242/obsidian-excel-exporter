import { toTemplateText } from "../config/mapping";
import type { ExportMode } from "../config/schema";
import { lookupKey, type Runtime, type Scope } from "../graph/context";
import { iterateCollection, parseCellText, renderTemplate } from "../excel/placeholders";
import { collectionScope, rootScope, type PreparedRun } from "../export/runner";
import { isNoteRecord, type NoteRecord } from "../model/note-record";
import { LinkRef, wrapFrontmatterValue } from "../values/link";
import { isEmptyValue, normalizeValue, parseIsoDate, scalarToText, starsToNumber, DEFAULT_NORMALIZE } from "../values/normalize";

export type FieldKind = "value" | "text" | "number" | "date" | "stars" | "link" | "list" | "collection" | "file" | "special" | "group";

/** A field the user can pick in Setup. `expr` is a placeholder path (without braces). */
export interface FieldNode {
	label: string;
	expr: string;
	kind: FieldKind;
	sample?: string;
	/** Lazily computed children (links are navigable, relations list their item fields). */
	children?: () => FieldNode[];
}

const MAX_NOTES = 200;
const MAX_DEPTH = 3;

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

function sampleText(v: unknown): string {
	const t = scalarToText(normalizeValue(v, DEFAULT_NORMALIZE));
	return t.length > 60 ? `${t.slice(0, 57)}…` : t;
}

function flatValues(v: unknown): unknown[] {
	return Array.isArray(v) ? v.flatMap(flatValues) : [v];
}

function kindOf(values: unknown[]): FieldKind {
	const atoms = values.flatMap(flatValues).filter((x) => !isEmptyValue(x));
	if (!atoms.length) return "value";
	const isList = values.some((v) => Array.isArray(v) && v.length > 1);
	if (atoms.some((a) => a instanceof LinkRef)) return isList ? "list" : "link";
	if (isList) return "list";
	if (atoms.every((a) => typeof a === "number")) return "number";
	if (atoms.every((a) => typeof a === "string" && parseIsoDate(a))) return "date";
	if (atoms.every((a) => typeof a === "string" && starsToNumber(a) !== null && /⭐/.test(a))) return "stars";
	return "text";
}

function uniqueNotes(notes: NoteRecord[]): NoteRecord[] {
	const seen = new Set<string>();
	return notes.filter((n) => (seen.has(n.path) ? false : (seen.add(n.path), true)));
}

/** Fields of a set of notes reachable at `prefix` (frontmatter keys, file.*, relations). */
export function noteFields(prefix: string[], notes: NoteRecord[], rt: Runtime, depth = 0): FieldNode[] {
	const sample = notes.slice(0, MAX_NOTES);
	const out: FieldNode[] = [];

	// Relations (only roots have them).
	const relNames = new Map<string, NoteRecord[]>();
	for (const n of sample) {
		for (const [name, items] of rt.relationsOf(n) ?? []) relNames.set(name, [...(relNames.get(name) ?? []), ...items]);
	}
	for (const [name, items] of relNames) {
		const first = rt.relationsOf(sample[0])?.get(name)?.length ?? 0;
		const segs = [...prefix, name];
		out.push({
			label: name,
			expr: pathExpr(segs),
			kind: "collection",
			sample: `${first} ${first === 1 ? "elemento" : "elementos"} en la muestra`,
			children: depth < MAX_DEPTH ? () => noteFields(segs, uniqueNotes(items), rt, depth + 1) : undefined,
		});
	}

	// Frontmatter keys, first-seen casing.
	const keys = new Map<string, string>();
	for (const n of sample) for (const k of Object.keys(n.frontmatter)) if (!keys.has(k.toLowerCase())) keys.set(k.toLowerCase(), k);
	const sortedKeys = [...keys.values()].sort((a, b) => a.localeCompare(b));
	for (const key of sortedKeys) {
		const values = sample
			.map((n) => {
				const [found, v] = lookupKey(n.frontmatter, key);
				return found ? wrapFrontmatterValue(v, n.path) : undefined;
			})
			.filter((v) => !isEmptyValue(v));
		const segs = [...prefix, key];
		const kind = kindOf(values);
		const node: FieldNode = { label: key, expr: pathExpr(segs), kind, sample: values.length ? sampleText(values[0]) : "(vacío)" };
		if ((kind === "link" || kind === "list") && depth < MAX_DEPTH) {
			const links = values.flatMap(flatValues).filter((v): v is LinkRef => v instanceof LinkRef);
			node.children = () => {
				const targets = uniqueNotes(links.map((l) => rt.resolveLink(l, false)).filter((n): n is NoteRecord => !!n));
				return targets.length ? noteFields(segs, targets, rt, depth + 1) : [];
			};
		}
		out.push(node);
	}

	const first = sample[0];
	const fileSample = first ? rt.fileInfo(first) : undefined;
	out.push({
		label: "Archivo (file.*)",
		expr: pathExpr([...prefix, "file"]),
		kind: "group",
		children: () =>
			(["name", "path", "folder", "ctime", "mtime", "link"] as const).map((p) => ({
				label: `file.${p}`,
				expr: pathExpr([...prefix, "file", p]),
				kind: "file" as const,
				sample: fileSample ? sampleText(fileSample[p]) : undefined,
			})),
	});
	return out;
}

export interface FieldContext {
	run: PreparedRun;
	/** Whether the sheet is filled once per root (file-per-root, or the first sheet of sheet-per-root). */
	perRoot: boolean;
	/** Collection path of the row being edited, if the row repeats. */
	rowPath?: string[];
}

export function isPerRootSheet(mode: ExportMode, sheetIndex: number): boolean {
	return mode === "file-per-root" || (mode === "sheet-per-root" && sheetIndex === 0);
}

/** The scope a sheet is filled with, using one root as sample. */
export function baseScope(run: PreparedRun, perRoot: boolean, sampleIndex = 0): Scope | null {
	if (!perRoot) return collectionScope(run);
	const root = run.roots[sampleIndex];
	return root ? rootScope(run, root, sampleIndex) : null;
}

/** Notes bound to each variable of a repeated row, across all roots (for field discovery). */
function rowVariableNotes(ctx: FieldContext): Map<string, NoteRecord[]> {
	const out = new Map<string, NoteRecord[]>();
	if (!ctx.rowPath) return out;
	const scopes: Scope[] = ctx.perRoot
		? ctx.run.roots.slice(0, MAX_NOTES).map((r, i) => rootScope(ctx.run, r, i))
		: [collectionScope(ctx.run)];
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

/** Top-level field tree for a cell (§ Setup). */
export function contextFields(ctx: FieldContext): FieldNode[] {
	const { run } = ctx;
	const alias = run.cd.def.root.alias;
	const rt = run.rt;
	const groups: FieldNode[] = [];

	if (ctx.rowPath) {
		const vars = rowVariableNotes(ctx);
		for (const name of [...ctx.rowPath].reverse()) {
			const notes = vars.get(name) ?? [];
			groups.push({
				label: `Fila: cada ${name}`,
				expr: name,
				kind: "group",
				sample: `${notes.length} notas`,
				children: () => [
					{ label: `${name} (nombre de la nota)`, expr: name, kind: "link", sample: notes[0]?.basename },
					...noteFields([name], notes, rt),
				],
			});
		}
	}

	const roots = run.roots;
	const rootIsItem = ctx.perRoot || (ctx.rowPath && ctx.rowPath[0].toLowerCase() === alias.toLowerCase());
	if (!(ctx.rowPath && ctx.rowPath[0].toLowerCase() === alias.toLowerCase())) {
		groups.push({
			label: rootIsItem ? `${alias} (nota raíz)` : `${alias} (todas las raíces)`,
			expr: alias,
			kind: rootIsItem ? "group" : "collection",
			sample: rootIsItem ? roots[0]?.basename : `${roots.length} notas`,
			children: () => [
				{ label: `${alias} (nombre de la nota)`, expr: alias, kind: "link", sample: roots[0]?.basename },
				...noteFields([alias], roots, rt),
			],
		});
	}

	const specials: FieldNode[] = [
		{ label: "@today (fecha de hoy)", expr: "@today", kind: "special", sample: scalarToText(rt.today) },
		{ label: "@now (fecha y hora)", expr: "@now", kind: "special" },
		{ label: "@exportName (nombre del export)", expr: "@exportName", kind: "special", sample: run.cd.def.name },
		{ label: "@count (cantidad)", expr: "@count", kind: "special" },
	];
	if (ctx.rowPath) specials.unshift({ label: "@index (nº de fila 1, 2, 3…)", expr: "@index", kind: "special" });
	groups.push({ label: "Variables especiales", expr: "@", kind: "group", children: () => specials });
	return groups;
}

export interface CollectionOption {
	value: string;
	label: string;
}

/** Collections a row can repeat over. */
export function collectionOptions(run: PreparedRun, perRoot: boolean): CollectionOption[] {
	const { alias } = run.cd.def.root;
	const rels = run.cd.relations.map((r) => r.name);
	if (perRoot) return rels.map((r) => ({ value: r, label: `cada elemento de "${r}"` }));
	return [
		{ value: alias, label: `cada ${alias} (una fila por nota raíz)` },
		...rels.map((r) => ({ value: `${alias}.${r}`, label: `cada "${r}" de cada ${alias} (flatten)` })),
	];
}

/** Evaluates a mapping expression against a sample scope, for the Setup preview. */
export function previewExpression(expr: string, scope: Scope, run: PreparedRun): string {
	const tpl = parseCellText(toTemplateText(expr));
	if (!tpl) return expr;
	if (tpl.issues.length) return `⚠ ${tpl.issues[0]}`;
	const r = renderTemplate({ ...tpl, each: undefined }, scope, { rt: run.rt, opts: run.normalize, location: "setup", warnUnresolved: false });
	return scalarToText(r.value);
}

/** Scope of the first element of a repeated row (or the base scope). */
export function rowSampleScope(base: Scope, rowPath: string[] | undefined, rt: Runtime): Scope | null {
	if (!rowPath) return base;
	return iterateCollection(rowPath, base, rt)?.[0] ?? null;
}
