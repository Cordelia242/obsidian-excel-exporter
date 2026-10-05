import type { Runtime } from "../graph/context";
import { isNoteRecord, type NoteRecord } from "../model/note-record";
import { LinkRef, wrapFrontmatterValue } from "../values/link";
import { isEmptyValue, parseIsoDate, scalarToText, normalizeValue, DEFAULT_NORMALIZE } from "../values/normalize";
import type { Condition } from "./conditions";

export type PropKind = "text" | "number" | "date" | "stars" | "link" | "list" | "bool" | "empty";

export interface PropertyInfo {
	key: string;
	kind: PropKind;
	/** Notes (of the given set) that have a value. */
	count: number;
	sample: string;
}

function atoms(v: unknown): unknown[] {
	return Array.isArray(v) ? v.flatMap(atoms) : [v];
}

function atomText(a: unknown): string {
	if (a instanceof LinkRef) return a.targetName;
	if (isNoteRecord(a)) return a.basename;
	return scalarToText(normalizeValue(a, { ...DEFAULT_NORMALIZE, stars: "raw" }));
}

function kindOfValues(values: unknown[]): PropKind {
	const flat = values.flatMap(atoms).filter((x) => !isEmptyValue(x));
	if (!flat.length) return "empty";
	if (values.some((v) => Array.isArray(v) && v.length > 1)) return "list";
	if (flat.some((a) => a instanceof LinkRef)) return "link";
	if (flat.every((a) => typeof a === "number")) return "number";
	if (flat.every((a) => typeof a === "boolean")) return "bool";
	if (flat.every((a) => typeof a === "string" && /^[⭐½️\s]+$/.test(a))) return "stars";
	if (flat.every((a) => typeof a === "string" && parseIsoDate(a))) return "date";
	return "text";
}

function valuesOf(notes: NoteRecord[], key: string): unknown[] {
	const lower = key.toLowerCase();
	const out: unknown[] = [];
	for (const n of notes) {
		const k = Object.keys(n.frontmatter).find((x) => x.toLowerCase() === lower);
		if (k === undefined) continue;
		const v = wrapFrontmatterValue(n.frontmatter[k], n.path);
		if (!isEmptyValue(v)) out.push(v);
	}
	return out;
}

/** Properties present in a set of notes, with type and an example value. */
export function propertiesOf(notes: NoteRecord[]): PropertyInfo[] {
	const keys = new Map<string, string>();
	for (const n of notes.slice(0, 500)) for (const k of Object.keys(n.frontmatter)) if (!keys.has(k.toLowerCase())) keys.set(k.toLowerCase(), k);
	return [...keys.values()]
		.map((key) => {
			const values = valuesOf(notes, key);
			const first = values[0];
			const sample = first === undefined ? "" : atoms(first).map(atomText).join(", ");
			return { key, kind: kindOfValues(values), count: values.length, sample: sample.length > 50 ? `${sample.slice(0, 47)}…` : sample };
		})
		.sort((a, b) => b.count - a.count || a.key.localeCompare(b.key));
}

export interface ValueCount {
	value: string;
	count: number;
}

/** Distinct values of a property (list items and link names counted separately), most frequent first. */
export function valuesOfProperty(notes: NoteRecord[], key: string, limit = 100): ValueCount[] {
	if (key === "file.name") return notes.map((n) => ({ value: n.basename, count: 1 })).slice(0, limit);
	if (key === "file.folder") {
		const counts = new Map<string, number>();
		for (const n of notes) {
			const f = n.path.includes("/") ? n.path.slice(0, n.path.lastIndexOf("/")) : "";
			counts.set(f, (counts.get(f) ?? 0) + 1);
		}
		return [...counts].map(([value, count]) => ({ value, count })).sort((a, b) => b.count - a.count).slice(0, limit);
	}
	const counts = new Map<string, number>();
	for (const v of valuesOf(notes, key)) {
		const seen = new Set<string>();
		for (const a of atoms(v)) {
			if (isEmptyValue(a)) continue;
			const t = atomText(a);
			if (seen.has(t)) continue;
			seen.add(t);
			counts.set(t, (counts.get(t) ?? 0) + 1);
		}
	}
	return [...counts]
		.map(([value, count]) => ({ value, count }))
		.sort((a, b) => b.count - a.count || a.value.localeCompare(b.value))
		.slice(0, limit);
}

const TYPE_KEYS = ["categories", "category", "categoria", "categorias", "type", "tipo", "tags"];

export interface TypeSuggestion {
	label: string;
	condition: Condition;
	count: number;
	examples: string[];
}

/** "Kinds of notes" in the vault (by categories/type/tags values, then by folder), for one-click root selection. */
export function noteTypeSuggestions(notes: NoteRecord[], limit = 12): TypeSuggestion[] {
	const out: TypeSuggestion[] = [];
	const present = new Set(notes.flatMap((n) => Object.keys(n.frontmatter).map((k) => k.toLowerCase())));
	for (const key of TYPE_KEYS.filter((k) => present.has(k))) {
		const realKey = notes.flatMap((n) => Object.keys(n.frontmatter)).find((k) => k.toLowerCase() === key) ?? key;
		for (const { value, count } of valuesOfProperty(notes, realKey, 50)) {
			const matching = notes.filter((n) => valuesOf([n], realKey).flatMap(atoms).some((a) => atomText(a) === value));
			out.push({
				label: value,
				condition: { field: realKey, op: "contains", value },
				count,
				examples: matching.slice(0, 3).map((n) => n.basename),
			});
		}
	}
	const folders = valuesOfProperty(notes, "file.folder", 30).filter((f) => f.value && f.count > 1);
	for (const f of folders) {
		if (out.some((s) => s.count === f.count && s.examples.length)) continue;
		const matching = notes.filter((n) => n.path.startsWith(`${f.value}/`) && !n.path.slice(f.value.length + 1).includes("/"));
		out.push({
			label: `Carpeta ${f.value}`,
			condition: { field: "file.folder", op: "==", value: f.value },
			count: f.count,
			examples: matching.slice(0, 3).map((n) => n.basename),
		});
	}
	return out.sort((a, b) => b.count - a.count).slice(0, limit);
}

export interface RelationSuggestion {
	/** Suggested internal name (identifier). */
	name: string;
	label: string;
	/** Field of the related notes that links to the root. */
	field: string;
	/** Condition that identifies the related notes (may be null). */
	from: Condition | null;
	/** Related notes linking to at least one root. */
	count: number;
	examples: string[];
}

export function slugify(text: string): string {
	const s = text
		.normalize("NFD")
		.replace(/[̀-ͯ]/g, "")
		.toLowerCase()
		.replace(/[^a-z0-9]+/g, "_")
		.replace(/^_+|_+$/g, "");
	return /^[a-z_]/.test(s) ? s : `r_${s || "rel"}`;
}

/**
 * Finds notes that point to the roots through some property (backlinks), e.g.
 * "4 notes Interview Evaluations point to processes through `process`".
 */
export function relationSuggestions(notes: NoteRecord[], roots: NoteRecord[], rt: Runtime): RelationSuggestion[] {
	const rootPaths = new Set(roots.map((r) => r.path));
	const byField = new Map<string, NoteRecord[]>();
	for (const n of notes) {
		if (rootPaths.has(n.path)) continue;
		for (const [key, raw] of Object.entries(n.frontmatter)) {
			const linked = atoms(wrapFrontmatterValue(raw, n.path)).some(
				(a) => a instanceof LinkRef && rootPaths.has(rt.resolveLink(a, false)?.path ?? ""),
			);
			if (linked) byField.set(key, [...(byField.get(key) ?? []), n]);
		}
	}
	const out: RelationSuggestion[] = [];
	for (const [field, linking] of byField) {
		// Most specific type value shared by all linking notes.
		let from: Condition | null = null;
		let best = Infinity;
		const present = new Set(linking.flatMap((n) => Object.keys(n.frontmatter)));
		for (const key of [...present].filter((k) => TYPE_KEYS.includes(k.toLowerCase()))) {
			for (const { value, count } of valuesOfProperty(linking, key)) {
				if (count !== linking.length) continue;
				const vaultCount = valuesOfProperty(notes, key, 1000).find((v) => v.value === value)?.count ?? Infinity;
				if (vaultCount < best) {
					best = vaultCount;
					from = { field: key, op: "contains", value };
				}
			}
		}
		const label = from ? from.value : field;
		out.push({ name: slugify(label), label, field, from, count: linking.length, examples: linking.slice(0, 3).map((n) => n.basename) });
	}
	return out.sort((a, b) => b.count - a.count);
}

