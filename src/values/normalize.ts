import { FileInfo, isNoteRecord } from "../model/note-record";
import { INLINE_LINK_RE, LinkRef } from "./link";

export interface NormalizeOptions {
	links: "display" | "target" | "raw";
	listSeparator: string;
	stars: "number" | "raw";
	emptyValue: string;
}

export const DEFAULT_NORMALIZE: NormalizeOptions = {
	links: "display",
	listSeparator: ", ",
	stars: "number",
	emptyValue: "",
};

/** Value types that can be written to an Excel cell. */
export type CellScalar = string | number | boolean | Date | null;

/** Marker for "the path did not resolve" (distinct from an explicit null). */
export const MISSING: unique symbol = Symbol("missing");
export type Missing = typeof MISSING;

/** Text produced by a pipe; written as-is (never re-parsed as a date, stars or link). */
export class PlainText {
	constructor(readonly text: string) {}
}

export function isEmptyValue(v: unknown): boolean {
	return (
		v === MISSING ||
		(v instanceof PlainText && v.text === "") ||
		v === null ||
		v === undefined ||
		v === "" ||
		(Array.isArray(v) && v.every(isEmptyValue))
	);
}

const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

/**
 * Parses an ISO date (`YYYY-MM-DD`, optional time) into a Date built from
 * local components, so `2026-03-01` is local midnight, not UTC midnight.
 * A trailing timezone designator is ignored on purpose (wall-clock time is kept).
 */
export function parseIsoDate(text: string): Date | null {
	const m = ISO_DATE_RE.exec(text.trim());
	if (!m) return null;
	const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
	const [h, mi, s] = [Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0)];
	const date = new Date(y, mo - 1, d, h, mi, s);
	if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
	return date;
}

export function isDateOnly(d: Date): boolean {
	return d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0 && d.getMilliseconds() === 0;
}

export function startOfToday(now: Date = new Date()): Date {
	return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

const STAR = "⭐";

/** `⭐⭐⭐` → 3; `⭐⭐½` → 2.5. Returns null when the text is not a star rating. */
export function countStars(text: string): number | null {
	const t = text.replace(/️/g, "").replace(/\s+/g, "");
	if (!t) return null;
	let n = 0;
	for (const ch of t) {
		if (ch === STAR) n += 1;
		else if (ch === "½") n += 0.5;
		else return null;
	}
	return n;
}

/** Number from a star string, a number, or a numeric string. */
export function starsToNumber(v: unknown): number | null {
	if (typeof v === "number") return v;
	if (typeof v === "string") {
		const s = countStars(v);
		if (s !== null) return s;
		const n = Number(v.trim());
		return v.trim() !== "" && Number.isFinite(n) ? n : null;
	}
	return null;
}

const pad = (n: number, w = 2) => String(n).padStart(w, "0");

/** Formats with tokens yyyy, yy, MM, M, dd, d, HH, H, mm, ss. */
export function formatDate(d: Date, fmt: string): string {
	return fmt.replace(/yyyy|yy|MM|M|dd|d|HH|H|mm|ss/g, (tok) => {
		switch (tok) {
			case "yyyy": return String(d.getFullYear());
			case "yy": return pad(d.getFullYear() % 100);
			case "MM": return pad(d.getMonth() + 1);
			case "M": return String(d.getMonth() + 1);
			case "dd": return pad(d.getDate());
			case "d": return String(d.getDate());
			case "HH": return pad(d.getHours());
			case "H": return String(d.getHours());
			case "mm": return pad(d.getMinutes());
			case "ss": return pad(d.getSeconds());
		}
		return tok;
	});
}

/** Coerces a value into a Date if it is one or looks like an ISO date. */
export function toDate(v: unknown): Date | null {
	if (v instanceof Date) return v;
	if (typeof v === "string") return parseIsoDate(v);
	if (Array.isArray(v) && v.length === 1) return toDate(v[0]);
	return null;
}

export function linkText(link: LinkRef, mode: NormalizeOptions["links"]): string {
	if (mode === "raw") return link.raw;
	if (mode === "target") return link.targetName;
	return link.display;
}

/** Replaces wikilinks embedded in free text with their display text. */
export function replaceInlineLinks(text: string, mode: NormalizeOptions["links"]): string {
	if (mode === "raw" || !text.includes("[[")) return text;
	return text.replace(INLINE_LINK_RE, (_m, linkpath: string, _sub: string | undefined, alias: string | undefined) => {
		const target = (linkpath.split("/").pop() ?? linkpath).replace(/\.md$/i, "");
		return mode === "display" && alias ? alias : target;
	});
}

export type WarnFn = (message: string) => void;

/**
 * Converts a resolved value into something writable to a cell (§8).
 * Lists of one element keep the element's type; longer lists are joined.
 */
export function normalizeValue(v: unknown, opts: NormalizeOptions, warn?: WarnFn): CellScalar {
	const empty = (): CellScalar => (opts.emptyValue === "" ? null : opts.emptyValue);
	if (isEmptyValue(v)) return empty();
	if (Array.isArray(v)) {
		const items = v.filter((x) => !isEmptyValue(x)).map((x) => normalizeValue(x, opts, warn));
		const nonEmpty = items.filter((x) => x !== null);
		if (nonEmpty.length === 0) return empty();
		if (nonEmpty.length === 1) return nonEmpty[0];
		return nonEmpty.map(scalarToText).join(opts.listSeparator);
	}
	if (v instanceof PlainText) return v.text;
	if (v instanceof LinkRef) return linkText(v, opts.links);
	if (isNoteRecord(v)) return v.basename;
	if (v instanceof FileInfo) return v.basename;
	if (v instanceof Date) return v;
	if (typeof v === "number" || typeof v === "boolean") return v;
	if (typeof v === "string") {
		if (opts.stars === "number") {
			const stars = countStars(v);
			if (stars !== null) return stars;
		}
		const date = parseIsoDate(v);
		if (date) return date;
		return replaceInlineLinks(v, opts.links);
	}
	if (typeof v === "object") {
		warn?.(`Valor de tipo objeto exportado como JSON: ${JSON.stringify(v)}`);
		return JSON.stringify(v);
	}
	return anyToText(v);
}

/** Like {@link normalizeValue} but without link cleanup or star/date parsing (`| raw`). */
export function rawValue(v: unknown, opts: NormalizeOptions): CellScalar {
	if (isEmptyValue(v)) return opts.emptyValue === "" ? null : opts.emptyValue;
	if (Array.isArray(v)) {
		return v.filter((x) => !isEmptyValue(x)).map((x) => scalarToText(rawValue(x, opts))).join(opts.listSeparator);
	}
	if (v instanceof PlainText) return v.text;
	if (v instanceof LinkRef) return v.raw;
	if (isNoteRecord(v)) return v.basename;
	if (v instanceof FileInfo) return v.basename;
	if (v instanceof Date || typeof v === "number" || typeof v === "boolean" || typeof v === "string") return v;
	return JSON.stringify(v);
}

/** Text for any value: primitives as-is, objects as JSON (never "[object Object]"). */
export function anyToText(v: unknown): string {
	if (typeof v === "string") return v;
	if (typeof v === "number" || typeof v === "boolean" || typeof v === "bigint") return String(v);
	if (v === null || v === undefined) return "";
	try {
		return JSON.stringify(v) ?? "";
	} catch {
		return "";
	}
}

/** Text form of a cell scalar, used when a placeholder is mixed with other text. */
export function scalarToText(v: CellScalar): string {
	if (v === null) return "";
	if (v instanceof Date) return isDateOnly(v) ? formatDate(v, "yyyy-MM-dd") : formatDate(v, "yyyy-MM-dd HH:mm");
	return String(v);
}
