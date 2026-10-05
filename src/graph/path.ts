import { FileInfo, isNoteRecord } from "../model/note-record";
import { LinkRef, wrapFrontmatterValue } from "../values/link";
import { MISSING } from "../values/normalize";
import { lookupKey, type Runtime, type Scope } from "./context";

export class PathError extends Error {}

/**
 * Parses `a.b c.d` / `a["x.y"].b` into segments. Segments may contain spaces
 * (`interviews.tech rating`); brackets allow dots inside a segment.
 */
export function parsePath(src: string): string[] {
	const segs: string[] = [];
	let buf = "";
	let afterBracket = false;
	let i = 0;
	const flush = (atDot: boolean) => {
		const t = buf.trim();
		if (t) segs.push(t);
		else if (atDot && !afterBracket) throw new PathError(`Segmento vacío en "${src}"`);
		buf = "";
		afterBracket = false;
	};
	while (i < src.length) {
		const ch = src[i];
		if (ch === ".") {
			flush(true);
			i++;
		} else if (ch === "[") {
			if (buf.trim()) flush(false);
			buf = "";
			let j = i + 1;
			while (j < src.length && src[j] === " ") j++;
			const quote = src[j];
			if (quote !== '"' && quote !== "'") throw new PathError(`Se esperaba una comilla después de "[" en "${src}"`);
			const end = src.indexOf(quote, j + 1);
			if (end < 0) throw new PathError(`Comilla sin cerrar en "${src}"`);
			const name = src.slice(j + 1, end);
			let k = end + 1;
			while (k < src.length && src[k] === " ") k++;
			if (src[k] !== "]") throw new PathError(`Se esperaba "]" en "${src}"`);
			segs.push(name);
			afterBracket = true;
			i = k + 1;
		} else {
			if (afterBracket && ch.trim()) throw new PathError(`Texto inesperado después de "]" en "${src}"`);
			buf += ch;
			i++;
		}
	}
	const t = buf.trim();
	if (t) segs.push(t);
	else if (!afterBracket) throw new PathError(segs.length ? `Segmento vacío al final de "${src}"` : "Path vacío");
	return segs;
}

/** Resolves the first segment against the scope (`@special` or a variable). */
export function resolveHead(seg: string, scope: Scope): unknown {
	if (seg.startsWith("@")) {
		const [found, v] = lookupKey(scope.specials, seg.slice(1));
		return found ? v : MISSING;
	}
	const [found, v] = lookupKey(scope.vars, seg);
	return found ? v : MISSING;
}

export function resolvePath(segs: string[], scope: Scope, rt: Runtime): unknown {
	let value = resolveHead(segs[0], scope);
	for (let i = 1; i < segs.length; i++) value = step(value, segs[i], rt);
	return value;
}

/** Resolves a path relative to a note (used by filters and `on`). */
export function resolveOnValue(base: unknown, segs: string[], rt: Runtime): unknown {
	let value = base;
	for (const seg of segs) value = step(value, seg, rt);
	return value;
}

/**
 * One navigation step. Links are resolved and traversed; lists are mapped
 * (and flattened one level), so `interviews.people.role` yields a list.
 */
export function step(value: unknown, seg: string, rt: Runtime): unknown {
	if (value === MISSING || value === null || value === undefined) return MISSING;
	if (Array.isArray(value)) {
		const out: unknown[] = [];
		let anyFound = false;
		for (const item of value) {
			const r = step(item, seg, rt);
			if (r === MISSING) continue;
			anyFound = true;
			if (Array.isArray(r)) out.push(...(r as unknown[]));
			else out.push(r);
		}
		return anyFound || value.length === 0 ? out : MISSING;
	}
	if (value instanceof LinkRef) {
		const note = rt.resolveLink(value);
		return note ? step(note, seg, rt) : MISSING;
	}
	if (isNoteRecord(value)) {
		const rels = rt.relationsOf(value);
		if (rels) {
			for (const [name, items] of rels) {
				if (name === seg || name.toLowerCase() === seg.toLowerCase()) return items;
			}
		}
		if (seg.toLowerCase() === "file") return rt.fileInfo(value);
		const [found, v] = lookupKey(value.frontmatter, seg);
		return found ? wrapFrontmatterValue(v, value.path) : MISSING;
	}
	if (value instanceof FileInfo) {
		const [found, v] = lookupKey(value as unknown as Record<string, unknown>, seg);
		return found ? v : MISSING;
	}
	if (typeof value === "object" && !(value instanceof Date)) {
		const [found, v] = lookupKey(value as Record<string, unknown>, seg);
		return found ? wrapFrontmatterValue(v, "") : MISSING;
	}
	return MISSING;
}
