import type { Runtime, Scope } from "../graph/context";
import { resolveHead, resolveOnValue, resolvePath } from "../graph/path";
import { isNoteRecord, type NoteRecord } from "../model/note-record";
import { LinkRef, parseWikilink } from "../values/link";
import { isDateOnly, isEmptyValue, MISSING, parseIsoDate } from "../values/normalize";
import type { CmpOp, Expr, ValueNode } from "./parser";

/** Evaluates a filter against a note. `scope` provides alias references (e.g. the root). */
export function evaluate(expr: Expr, note: NoteRecord, rt: Runtime, scope?: Scope): boolean {
	switch (expr.type) {
		case "and":
			return evaluate(expr.left, note, rt, scope) && evaluate(expr.right, note, rt, scope);
		case "or":
			return evaluate(expr.left, note, rt, scope) || evaluate(expr.right, note, rt, scope);
		case "not":
			return !evaluate(expr.expr, note, rt, scope);
		case "exists": {
			const present = !isEmptyValue(resolveOnValue(note, expr.path, rt));
			return expr.negate ? !present : present;
		}
		case "cmp":
			return compare(resolveOnValue(note, expr.path, rt), expr.op, valueOf(expr.value, rt, scope), rt);
	}
}

function valueOf(node: ValueNode, rt: Runtime, scope?: Scope): unknown {
	if (node.kind === "literal") return node.value;
	if (node.kind === "today") return rt.today;
	if (scope && resolveHead(node.path[0], scope) !== MISSING) return resolvePath(node.path, scope, rt);
	// Unknown bare word: treat as text so `status == Closed` works.
	return node.text;
}

function flatten(v: unknown): unknown[] {
	if (Array.isArray(v)) return v.flatMap(flatten);
	return [v];
}

export function compare(left: unknown, op: CmpOp, right: unknown, rt: Runtime): boolean {
	const missing = isEmptyValue(left);
	if (right === null) {
		if (op === "==") return missing;
		if (op === "!=") return !missing;
		return false;
	}
	if (missing) return op === "!=" || op === "!contains";
	const atoms = flatten(left).filter((x) => !isEmptyValue(x));
	const rights = Array.isArray(right) ? flatten(right) : [right];
	const anyEq = () => atoms.some((a) => rights.some((b) => equals(a, b, rt)));
	switch (op) {
		case "==":
			return anyEq();
		case "!=":
			return !anyEq();
		case "contains":
		case "!contains": {
			let result: boolean;
			if (Array.isArray(left)) result = anyEq();
			else result = rights.some((b) => scalarContains(atoms[0], b, rt));
			return op === "contains" ? result : !result;
		}
		default:
			return atoms.some((a) =>
				rights.some((b) => {
					const c = order(a, b);
					if (c === null) return false;
					if (op === ">") return c > 0;
					if (op === ">=") return c >= 0;
					if (op === "<") return c < 0;
					return c <= 0;
				}),
			);
	}
}

interface NoteKey {
	path: string | null;
	names: string[];
}

function noteKey(v: unknown, rt: Runtime): NoteKey | null {
	if (isNoteRecord(v)) {
		return { path: v.path, names: [v.basename.toLowerCase(), v.path.replace(/\.md$/i, "").toLowerCase()] };
	}
	let link: LinkRef | null = v instanceof LinkRef ? v : null;
	if (!link && typeof v === "string") {
		const p = parseWikilink(v);
		if (p) link = new LinkRef(v, p.linkpath, p.subpath, p.alias, "");
	}
	if (!link) return null;
	const resolved = link.sourcePath ? rt.resolveLink(link, false) : null;
	return {
		path: resolved?.path ?? null,
		names: [link.targetName.toLowerCase(), link.linkpath.replace(/\.md$/i, "").toLowerCase()],
	};
}

function asNumber(v: unknown): number | null {
	if (typeof v === "number") return v;
	if (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v))) return Number(v);
	return null;
}

function asDate(v: unknown): Date | null {
	if (v instanceof Date) return v;
	if (typeof v === "string") return parseIsoDate(v);
	return null;
}

function text(v: unknown): string {
	if (v instanceof LinkRef) return v.targetName;
	if (isNoteRecord(v)) return v.basename;
	if (v instanceof Date) return v.toISOString();
	return String(v);
}

function dateCompare(a: Date, b: Date): number {
	if (isDateOnly(a) || isDateOnly(b)) {
		const da = new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime();
		const db = new Date(b.getFullYear(), b.getMonth(), b.getDate()).getTime();
		return Math.sign(da - db);
	}
	return Math.sign(a.getTime() - b.getTime());
}

export function equals(a: unknown, b: unknown, rt: Runtime): boolean {
	const ka = noteKey(a, rt);
	const kb = noteKey(b, rt);
	if (ka || kb) {
		if (ka && kb) {
			if (ka.path && kb.path) return ka.path === kb.path;
			return ka.names.some((n) => kb.names.includes(n));
		}
		const key = (ka ?? kb) as NoteKey;
		const other = ka ? b : a;
		return typeof other === "string" && key.names.includes(other.trim().replace(/\.md$/i, "").toLowerCase());
	}
	if (typeof a === "boolean" || typeof b === "boolean") {
		return String(a).toLowerCase() === String(b).toLowerCase();
	}
	const na = asNumber(a);
	const nb = asNumber(b);
	if (na !== null && nb !== null && (typeof a === "number" || typeof b === "number")) return na === nb;
	const da = asDate(a);
	const db = asDate(b);
	if (da && db) return dateCompare(da, db) === 0;
	return text(a).toLowerCase() === text(b).toLowerCase();
}

function scalarContains(a: unknown, b: unknown, rt: Runtime): boolean {
	if (typeof a === "string" && !parseWikilink(a)) {
		return a.toLowerCase().includes(text(b).toLowerCase());
	}
	if (a instanceof LinkRef) {
		const needle = text(b).toLowerCase();
		return a.targetName.toLowerCase().includes(needle) || (a.alias ?? "").toLowerCase().includes(needle);
	}
	return equals(a, b, rt);
}

/** Ordering for >, <, sort. Returns null when the values are not comparable. */
export function order(a: unknown, b: unknown): number | null {
	const na = asNumber(a);
	const nb = asNumber(b);
	if (na !== null && nb !== null) return Math.sign(na - nb);
	const da = asDate(a);
	const db = asDate(b);
	if (da && db) return dateCompare(da, db);
	if (da || db) return null;
	if (typeof a === "boolean" || typeof b === "boolean") return null;
	return Math.sign(text(a).toLowerCase().localeCompare(text(b).toLowerCase()));
}
