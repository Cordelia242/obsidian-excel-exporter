import { parseFilter, type Expr } from "../query/parser";

/** Operators offered by the visual filter builder. */
export type CondOp = "==" | "!=" | "contains" | "!contains" | ">" | ">=" | "<" | "<=" | "exists" | "!exists";

/** One row of the visual filter builder: `<field> <op> <value>`. */
export interface Condition {
	field: string;
	op: CondOp;
	value: string;
}

export const OP_LABELS: Record<CondOp, string> = {
	"==": "es",
	"!=": "no es",
	contains: "incluye",
	"!contains": "no incluye",
	">": "es mayor que",
	">=": "es desde",
	"<": "es menor que",
	"<=": "es hasta",
	exists: "tiene valor",
	"!exists": "está vacío",
};

export const VALUELESS_OPS: CondOp[] = ["exists", "!exists"];

const IDENT_RE = /^[A-Za-z_@À-￿][A-Za-z0-9_\-@À-￿]*$/;

/** `tech rating` → `["tech rating"]`; `file.name` stays as is. */
export function fieldToText(field: string): string {
	return field
		.split(".")
		.map((seg, i) => (IDENT_RE.test(seg) ? (i ? "." : "") + seg : `["${seg.replace(/"/g, '\\"')}"]`))
		.join("");
}

function valueToText(value: string): string {
	const v = value.trim();
	if (/^-?\d+(\.\d+)?$/.test(v)) return v;
	if (v === "true" || v === "false") return v;
	return `"${v.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

/** Builds the filter text for a list of conditions joined with `and`. */
export function conditionsToExpr(conds: Condition[]): string {
	return conds
		.filter((c) => c.field.trim())
		.map((c) => {
			const f = fieldToText(c.field.trim());
			if (VALUELESS_OPS.includes(c.op)) return `${f} ${c.op}`;
			return `${f} ${c.op} ${valueToText(c.value)}`;
		})
		.join(" and ");
}

function flattenAnd(e: Expr, out: Expr[]): void {
	if (e.type === "and") {
		flattenAnd(e.left, out);
		flattenAnd(e.right, out);
	} else out.push(e);
}

/**
 * Converts filter text back into builder rows. Returns null when the
 * expression uses something the builder cannot show (or, not, dates…).
 */
export function exprToConditions(src: string | undefined): Condition[] | null {
	if (!src || !src.trim()) return [];
	let ast: Expr;
	try {
		ast = parseFilter(src);
	} catch {
		return null;
	}
	const parts: Expr[] = [];
	flattenAnd(ast, parts);
	const out: Condition[] = [];
	for (const p of parts) {
		if (p.type === "exists") {
			out.push({ field: p.path.join("."), op: p.negate ? "!exists" : "exists", value: "" });
		} else if (p.type === "cmp") {
			let value: string;
			if (p.value.kind === "literal") {
				const v = p.value.value;
				if (v === null || v instanceof Date) return null;
				value = String(v);
			} else if (p.value.kind === "ref") {
				value = p.value.text;
			} else {
				return null;
			}
			out.push({ field: p.path.join("."), op: p.op, value });
		} else {
			return null;
		}
	}
	return out;
}
