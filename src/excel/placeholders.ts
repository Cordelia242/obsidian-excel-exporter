import { childScope, type Runtime, type Scope } from "../graph/context";
import { parsePath, PathError, resolveHead, resolvePath, step } from "../graph/path";
import { isNoteRecord } from "../model/note-record";
import { MISSING, isEmptyValue, scalarToText, type NormalizeOptions } from "../values/normalize";
import { applyPipes, parseExpression, type PipeCall, type Rendered } from "../values/pipes";

export type Part =
	| { kind: "text"; text: string }
	| { kind: "ph"; src: string; path: string[]; pipes: PipeCall[]; error?: string };

export interface CellTemplate {
	parts: Part[];
	/** Collection expression of a leading `{{#each ...}}`, if any. */
	each?: { src: string; path: string[]; error?: string };
	/** Problems found while parsing (misplaced #each, unclosed braces...). */
	issues: string[];
}

/** Finds the `}}` closing a placeholder that starts at `from`, ignoring braces inside quotes. */
function findClose(text: string, from: number): number {
	let quote: string | null = null;
	for (let i = from; i < text.length - 1; i++) {
		const ch = text[i];
		if (quote) {
			if (ch === quote) quote = null;
		} else if (ch === '"' || ch === "'") {
			quote = ch;
		} else if (ch === "}" && text[i + 1] === "}") {
			return i;
		}
	}
	return -1;
}

function parsePlaceholder(src: string): Part {
	const { path, pipes } = parseExpression(src);
	try {
		return { kind: "ph", src, path: parsePath(path), pipes };
	} catch (e) {
		return { kind: "ph", src, path: [], pipes, error: e instanceof PathError ? e.message : String(e) };
	}
}

/** Parses cell text into literal text and placeholders. Returns null if it has no `{{`. */
export function parseCellText(text: string): CellTemplate | null {
	if (!text.includes("{{")) return null;
	const parts: Part[] = [];
	const issues: string[] = [];
	let each: CellTemplate["each"];
	let i = 0;
	let first = true;
	while (i < text.length) {
		const open = text.indexOf("{{", i);
		if (open < 0) {
			parts.push({ kind: "text", text: text.slice(i) });
			break;
		}
		if (open > i) parts.push({ kind: "text", text: text.slice(i, open) });
		const close = findClose(text, open + 2);
		if (close < 0) {
			issues.push(`Placeholder sin cerrar: "${text.slice(open)}"`);
			parts.push({ kind: "text", text: text.slice(open) });
			break;
		}
		const inner = text.slice(open + 2, close).trim();
		const leading = first && text.slice(0, open).trim() === "";
		if (/^#each\b/i.test(inner)) {
			const src = inner.replace(/^#each\s*/i, "").trim();
			if (!leading) issues.push(`{{#each}} debe ir al inicio de la celda: "${text}"`);
			else if (each) issues.push(`Solo se permite un {{#each}} por celda`);
			else {
				try {
					each = { src, path: parsePath(src) };
				} catch (e) {
					each = { src, path: [], error: (e as Error).message };
				}
			}
		} else if (/^\/each$/i.test(inner)) {
			// Optional closing tag: ignored (a #each always covers its whole row).
		} else {
			parts.push(parsePlaceholder(inner));
		}
		first = false;
		i = close + 2;
	}
	// Merge the leading whitespace left by a removed #each marker.
	const cleaned = parts.filter((p) => !(p.kind === "text" && p.text === ""));
	if (each && cleaned.length && cleaned[0].kind === "text" && cleaned[0].text.trim() === "") cleaned.shift();
	return { parts: cleaned, each, issues };
}

export interface RenderContext {
	rt: Runtime;
	opts: NormalizeOptions;
	/** "Sheet!A1" — for warnings. */
	location: string;
	/** When false, unresolved placeholders do not produce warnings (used by validation). */
	warnUnresolved?: boolean;
}

function noteOfHead(path: string[], scope: Scope): string | undefined {
	const head = resolveHead(path[0], scope);
	return isNoteRecord(head) ? head.path : undefined;
}

export function evaluatePlaceholder(part: Extract<Part, { kind: "ph" }>, scope: Scope, ctx: RenderContext): Rendered {
	const warn = (message: string) => ctx.rt.report.warn("other", message, { location: ctx.location });
	if (part.error) {
		ctx.rt.report.warn("unresolved", `{{${part.src}}}: ${part.error}`, { location: ctx.location });
		return applyPipes(MISSING, part.pipes, { opts: ctx.opts, rt: ctx.rt, warn });
	}
	const value = resolvePath(part.path, scope, ctx.rt);
	if (value === MISSING && ctx.warnUnresolved !== false && !part.pipes.some((p) => p.name === "default")) {
		if (resolveHead(part.path[0], scope) === MISSING) {
			ctx.rt.report.warn("unresolved", `{{${part.src}}}: "${part.path[0]}" no existe en este contexto`, {
				location: ctx.location,
			});
		} else {
			ctx.rt.report.warn("unresolved", `{{${part.src}}} sin valor`, {
				location: ctx.location,
				note: noteOfHead(part.path, scope),
			});
		}
	}
	return applyPipes(value, part.pipes, { opts: ctx.opts, rt: ctx.rt, warn });
}

/**
 * Renders a cell template. A cell that is exactly one placeholder keeps the
 * value's type (number, Date, boolean); mixed text produces a string.
 */
export function renderTemplate(tpl: CellTemplate, scope: Scope, ctx: RenderContext): Rendered {
	const phs = tpl.parts.filter((p) => p.kind === "ph");
	if (tpl.parts.length === 0) return { value: null };
	if (tpl.parts.length === 1 && phs.length === 1) {
		return evaluatePlaceholder(phs[0], scope, ctx);
	}
	let text = "";
	const links: string[] = [];
	for (const p of tpl.parts) {
		if (p.kind === "text") text += p.text;
		else {
			const r = evaluatePlaceholder(p, scope, ctx);
			text += scalarToText(r.value);
			if (r.hyperlink) links.push(r.hyperlink);
		}
	}
	return links.length === 1 ? { value: text, hyperlink: links[0] } : { value: text };
}

/** Renders a template string (file names, sheet names) to plain text. */
export function renderString(src: string, scope: Scope, ctx: RenderContext): string {
	const tpl = parseCellText(src);
	if (!tpl) return src;
	for (const issue of tpl.issues) ctx.rt.report.warn("other", issue, { location: ctx.location });
	return scalarToText(renderTemplate({ ...tpl, each: undefined }, scope, ctx).value);
}

function toList(v: unknown): unknown[] {
	if (v === MISSING || v === null || v === undefined) return [];
	if (Array.isArray(v)) return v.filter((x) => !isEmptyValue(x));
	return [v];
}

/**
 * Expands a `#each` collection into one scope per item. `a.b.c` iterates
 * nested (flatten) and binds `a`, `b` and `c` to the current element of each level.
 * Returns null when the collection name is unknown.
 */
export function iterateCollection(path: string[], scope: Scope, rt: Runtime): Scope[] | null {
	const head = resolveHead(path[0], scope);
	if (head === MISSING) return null;
	type Frame = { vars: Record<string, unknown>; value: unknown };
	const bind = (vars: Record<string, unknown>, name: string, value: unknown): Record<string, unknown> => {
		const rels = isNoteRecord(value) ? rt.relationsOf(value) : undefined;
		const relVars = rels ? Object.fromEntries(rels) : {};
		return { ...relVars, ...vars, [name]: value };
	};
	let frames: Frame[] = toList(head).map((x) => ({ vars: bind({}, path[0], x), value: x }));
	for (let i = 1; i < path.length; i++) {
		const seg = path[i];
		frames = frames.flatMap((f) => toList(step(f.value, seg, rt)).map((x) => ({ vars: bind(f.vars, seg, x), value: x })));
	}
	const count = frames.length;
	return frames.map((f, idx) => childScope(scope, f.vars, { index: idx + 1, count }));
}
