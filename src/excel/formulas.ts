/**
 * Minimal A1-reference rewriting for formulas, used when `#each` rows are
 * inserted or removed (ExcelJS does not update formulas on spliceRows).
 */

// optional sheet prefix, then A1 or A1:B2 (with optional $). Not part of a longer token and not a function call.
// The leading group replaces a lookbehind (not supported on older iOS).
const REF_RE =
	/(^|[^A-Za-z0-9_.$'!])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(!])/g;

export interface RowRef {
	row: number;
	absolute: boolean;
}

/** Return the new row, or null to turn the whole reference into `#REF!`. */
export type RefMapper = (start: RowRef, end: RowRef | null) => { start: number; end: number | null } | null;

function sheetOf(prefix: string | undefined): string | null {
	if (!prefix) return null;
	const name = prefix.slice(0, -1);
	return name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
}

/**
 * Rewrites the references of `formula` that point to `targetSheet`.
 * `formulaSheet` is the sheet the formula lives in (unprefixed refs belong to it).
 */
export function mapFormulaRefs(formula: string, formulaSheet: string, targetSheet: string, mapper: RefMapper): string {
	// Split around string literals so their contents are never rewritten.
	const pieces = formula.split(/("(?:[^"]|"")*")/);
	return pieces
		.map((piece, idx) => {
			if (idx % 2 === 1) return piece;
			return piece.replace(
				REF_RE,
				(m, lead: string, prefix: string | undefined, c1abs: string, c1: string, r1abs: string, r1: string, c2abs?: string, c2?: string, r2abs?: string, r2?: string) => {
					const sheet = sheetOf(prefix) ?? formulaSheet;
					if (sheet.toLowerCase() !== targetSheet.toLowerCase()) return m;
					const start = { row: Number(r1), absolute: r1abs === "$" };
					const end = r2 !== undefined ? { row: Number(r2), absolute: r2abs === "$" } : null;
					const res = mapper(start, end);
					if (!res) return `${lead}#REF!`;
					let out = `${lead}${prefix ?? ""}${c1abs}${c1}${r1abs}${res.start}`;
					if (end && res.end !== null) out += `:${c2abs}${c2}${r2abs}${res.end}`;
					return out;
				},
			);
		})
		.join("");
}

/** Mapper for rows inserted (`delta > 0`) or removed (`delta < 0`) at the template row `at`. */
export function insertionMapper(at: number, delta: number): RefMapper {
	return (start, end) => {
		if (!end) {
			if (start.row > at) return { start: start.row + delta, end: null };
			if (start.row === at && delta < 0) return null;
			return { start: start.row, end: null };
		}
		const s = start.row > at ? start.row + delta : start.row;
		const e = end.row >= at ? end.row + delta : end.row;
		if (e < s) return null;
		return { start: s, end: e };
	};
}

/**
 * Mapper for the i-th copy (0-based) of a template row `at`: relative refs move
 * with the copy like an Excel copy/paste; refs below the block move by `delta`.
 */
export function copyMapper(at: number, delta: number, i: number): RefMapper {
	const one = (r: { row: number; absolute: boolean }) => {
		let row = r.row > at ? r.row + delta : r.row;
		if (!r.absolute) row += i;
		return row;
	};
	return (start, end) => ({ start: one(start), end: end ? one(end) : null });
}
