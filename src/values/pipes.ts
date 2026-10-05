import type { Runtime } from "../graph/context";
import { isNoteRecord, type NoteRecord } from "../model/note-record";
import { LinkRef } from "./link";
import {
	formatDate,
	isEmptyValue,
	MISSING,
	normalizeValue,
	PlainText,
	rawValue,
	scalarToText,
	starsToNumber,
	toDate,
	type CellScalar,
	type NormalizeOptions,
} from "./normalize";

export interface PipeCall {
	name: string;
	arg?: string;
}

export interface ParsedExpression {
	path: string;
	pipes: PipeCall[];
}

export const KNOWN_PIPES = ["raw", "target", "join", "first", "count", "stars", "date", "default", "upper", "lower", "link"];

/** Splits on `sep` outside of single/double quotes. */
export function splitOutsideQuotes(src: string, sep: string): string[] {
	const parts: string[] = [];
	let buf = "";
	let quote: string | null = null;
	for (let i = 0; i < src.length; i++) {
		const ch = src[i];
		if (quote) {
			if (ch === "\\" && i + 1 < src.length) {
				buf += ch + src[++i];
				continue;
			}
			if (ch === quote) quote = null;
			buf += ch;
		} else if (ch === '"' || ch === "'") {
			quote = ch;
			buf += ch;
		} else if (ch === sep) {
			parts.push(buf);
			buf = "";
		} else {
			buf += ch;
		}
	}
	parts.push(buf);
	return parts;
}

function unquote(s: string): string {
	const t = s.trim();
	if (t.length >= 2 && (t[0] === '"' || t[0] === "'") && t[t.length - 1] === t[0]) {
		return t.slice(1, -1).replace(/\\(.)/g, "$1");
	}
	return t;
}

/** `path | pipe:arg | pipe` → path + pipes. */
export function parseExpression(src: string): ParsedExpression {
	const [path, ...rest] = splitOutsideQuotes(src, "|");
	const pipes = rest.map((p) => {
		const t = p.trim();
		const colon = t.indexOf(":");
		if (colon < 0) return { name: t.toLowerCase() };
		return { name: t.slice(0, colon).trim().toLowerCase(), arg: unquote(t.slice(colon + 1)) };
	});
	return { path: path.trim(), pipes };
}

export interface Rendered {
	value: CellScalar;
	hyperlink?: string;
}

export interface PipeContext {
	opts: NormalizeOptions;
	rt: Runtime;
	warn: (message: string) => void;
}

function mapDeep(v: unknown, fn: (x: unknown) => unknown): unknown {
	return Array.isArray(v) ? v.map((x) => mapDeep(x, fn)) : fn(v);
}

function firstNote(v: unknown, rt: Runtime): NoteRecord | null {
	if (Array.isArray(v)) {
		for (const x of v) {
			const n = firstNote(x, rt);
			if (n) return n;
		}
		return null;
	}
	if (v instanceof LinkRef) return rt.resolveLink(v);
	if (isNoteRecord(v)) return v;
	return null;
}

/** Applies pipes (§6.5) to a resolved value and normalizes the result for a cell. */
export function applyPipes(value: unknown, pipes: PipeCall[], ctx: PipeContext): Rendered {
	let v = value;
	let raw = false;
	let hyperlink: string | undefined;
	const asText = (x: unknown) => scalarToText(raw ? rawValue(x, ctx.opts) : normalizeValue(x, ctx.opts, ctx.warn));

	for (const pipe of pipes) {
		switch (pipe.name) {
			case "raw":
				raw = true;
				break;
			case "target":
				v = mapDeep(v, (x) =>
					x instanceof LinkRef ? new PlainText(x.targetName) : isNoteRecord(x) ? new PlainText(x.basename) : x,
				);
				break;
			case "join":
				if (Array.isArray(v)) {
					const sep = pipe.arg ?? ctx.opts.listSeparator;
					v = new PlainText(v.filter((x) => !isEmptyValue(x)).map(asText).join(sep));
				}
				break;
			case "first":
				if (Array.isArray(v)) v = v.find((x) => !isEmptyValue(x)) ?? MISSING;
				break;
			case "count":
				v = Array.isArray(v) ? v.filter((x) => !isEmptyValue(x)).length : isEmptyValue(v) ? 0 : 1;
				break;
			case "stars": {
				const src = Array.isArray(v) ? v.find((x) => !isEmptyValue(x)) : v;
				const n = src instanceof PlainText ? starsToNumber(src.text) : starsToNumber(src);
				if (n === null && !isEmptyValue(src)) ctx.warn(`"${String(src)}" no es una calificación con estrellas`);
				v = n ?? MISSING;
				break;
			}
			case "date": {
				const d = toDate(v instanceof PlainText ? v.text : v);
				if (d) v = new PlainText(formatDate(d, pipe.arg || "yyyy-MM-dd"));
				else if (!isEmptyValue(v)) ctx.warn(`"${asText(v)}" no es una fecha`);
				break;
			}
			case "default":
				if (isEmptyValue(v)) v = new PlainText(pipe.arg ?? "");
				break;
			case "upper":
				v = new PlainText(asText(v).toUpperCase());
				break;
			case "lower":
				v = new PlainText(asText(v).toLowerCase());
				break;
			case "link": {
				const note = firstNote(v, ctx.rt);
				if (note) hyperlink = ctx.rt.fileUri(note);
				break;
			}
			default:
				ctx.rt.report.warn("pipe", `Pipe desconocido "${pipe.name}" (disponibles: ${KNOWN_PIPES.join(", ")})`);
		}
	}
	const final = raw ? rawValue(v, ctx.opts) : normalizeValue(v, ctx.opts, ctx.warn);
	return hyperlink ? { value: final, hyperlink } : { value: final };
}
