import { parseIsoDate } from "../values/normalize";
import { QueryError, tokenize, type Token } from "./lexer";

export type CmpOp = "==" | "!=" | ">" | ">=" | "<" | "<=" | "contains" | "!contains";

export type ValueNode =
	| { kind: "literal"; value: string | number | boolean | null | Date }
	| { kind: "today" }
	/** A bare identifier: an alias (e.g. the root) if one exists, else the identifier as text. */
	| { kind: "ref"; path: string[]; text: string };

export type Expr =
	| { type: "and" | "or"; left: Expr; right: Expr }
	| { type: "not"; expr: Expr }
	| { type: "cmp"; path: string[]; op: CmpOp; value: ValueNode }
	| { type: "exists"; path: string[]; negate: boolean };

const KEYWORDS = new Set(["and", "or", "not", "contains", "exists", "true", "false", "null", "today", "date"]);

/** Parses a filter expression (§7). Throws {@link QueryError} with the column and token. */
export function parseFilter(src: string): Expr {
	const p = new Parser(src);
	const expr = p.parseOr();
	p.expectEnd();
	return expr;
}

class Parser {
	private tokens: Token[];
	private i = 0;

	constructor(private src: string) {
		this.tokens = tokenize(src);
	}

	private peek(): Token {
		return this.tokens[this.i];
	}

	private next(): Token {
		return this.tokens[this.i++];
	}

	private fail(message: string, tok = this.peek()): never {
		const shown = tok.type === "eof" ? "fin de la expresión" : `"${tok.value}"`;
		throw new QueryError(`${message}; se encontró ${shown}`, tok.pos, this.src);
	}

	private isWord(tok: Token, word: string): boolean {
		return tok.type === "ident" && tok.value.toLowerCase() === word;
	}

	expectEnd(): void {
		if (this.peek().type !== "eof") this.fail("Se esperaba el fin de la expresión");
	}

	parseOr(): Expr {
		let left = this.parseAnd();
		while (this.isWord(this.peek(), "or") || (this.peek().type === "op" && this.peek().value === "||")) {
			this.next();
			left = { type: "or", left, right: this.parseAnd() };
		}
		return left;
	}

	private parseAnd(): Expr {
		let left = this.parseUnary();
		while (this.isWord(this.peek(), "and") || (this.peek().type === "op" && this.peek().value === "&&")) {
			this.next();
			left = { type: "and", left, right: this.parseUnary() };
		}
		return left;
	}

	private parseUnary(): Expr {
		const tok = this.peek();
		if (this.isWord(tok, "not") || (tok.type === "op" && tok.value === "!")) {
			this.next();
			return { type: "not", expr: this.parseUnary() };
		}
		if (tok.type === "lparen") {
			this.next();
			const e = this.parseOr();
			if (this.peek().type !== "rparen") this.fail('Se esperaba ")"');
			this.next();
			return e;
		}
		return this.parseCmp();
	}

	private parsePath(): string[] {
		const segs: string[] = [];
		const readSeg = () => {
			const tok = this.peek();
			if (tok.type === "ident") {
				this.next();
				segs.push(tok.value);
			} else if (tok.type === "lbracket") {
				this.next();
				const s = this.next();
				if (s.type !== "string") this.fail("Se esperaba un nombre entre comillas dentro de [ ]", s);
				if (this.peek().type !== "rbracket") this.fail('Se esperaba "]"');
				this.next();
				segs.push(s.value);
			} else {
				this.fail("Se esperaba un campo");
			}
		};
		readSeg();
		for (;;) {
			const tok = this.peek();
			if (tok.type === "dot") {
				this.next();
				readSeg();
			} else if (tok.type === "lbracket") {
				readSeg();
			} else {
				break;
			}
		}
		return segs;
	}

	private parseCmp(): Expr {
		const first = this.peek();
		if (first.type === "ident" && KEYWORDS.has(first.value.toLowerCase()) && first.value.toLowerCase() !== "date") {
			this.fail("Se esperaba un campo");
		}
		const path = this.parsePath();
		const tok = this.peek();
		if (this.isWord(tok, "exists")) {
			this.next();
			return { type: "exists", path, negate: false };
		}
		if (tok.type === "op" && tok.value === "!exists") {
			this.next();
			return { type: "exists", path, negate: true };
		}
		let op: CmpOp;
		if (tok.type === "op" && ["==", "!=", ">", ">=", "<", "<=", "!contains"].includes(tok.value)) {
			op = tok.value as CmpOp;
		} else if (this.isWord(tok, "contains")) {
			op = "contains";
		} else if (this.isWord(tok, "not") && this.isWord(this.tokens[this.i + 1], "contains")) {
			this.next();
			op = "!contains";
		} else {
			this.fail("Se esperaba un operador (==, !=, >, >=, <, <=, contains, !contains, exists, !exists)");
		}
		this.next();
		return { type: "cmp", path, op, value: this.parseValue() };
	}

	private parseValue(): ValueNode {
		const tok = this.peek();
		if (tok.type === "string") {
			this.next();
			return { kind: "literal", value: tok.value };
		}
		if (tok.type === "number") {
			this.next();
			const n = Number(tok.value);
			if (!Number.isFinite(n)) this.fail("Número inválido", tok);
			return { kind: "literal", value: n };
		}
		if (tok.type === "ident" || tok.type === "lbracket") {
			const word = tok.type === "ident" ? tok.value.toLowerCase() : "";
			if (word === "true" || word === "false") {
				this.next();
				return { kind: "literal", value: word === "true" };
			}
			if (word === "null") {
				this.next();
				return { kind: "literal", value: null };
			}
			if (word === "today") {
				this.next();
				return { kind: "today" };
			}
			if (word === "date" && this.tokens[this.i + 1]?.type === "lparen") {
				this.next();
				this.next();
				const s = this.next();
				if (s.type !== "string") this.fail('Se esperaba una fecha entre comillas: date("2026-03-01")', s);
				const d = parseIsoDate(s.value);
				if (!d) this.fail(`Fecha inválida "${s.value}"`, s);
				if (this.peek().type !== "rparen") this.fail('Se esperaba ")"');
				this.next();
				return { kind: "literal", value: d };
			}
			const start = this.peek().pos;
			const path = this.parsePath();
			const end = this.peek().pos;
			return { kind: "ref", path, text: this.src.slice(start, end).trim() };
		}
		this.fail("Se esperaba un valor");
	}
}
