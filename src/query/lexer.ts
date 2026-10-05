export type TokenType =
	| "string"
	| "number"
	| "ident"
	| "op"
	| "lparen"
	| "rparen"
	| "lbracket"
	| "rbracket"
	| "dot"
	| "comma"
	| "eof";

export interface Token {
	type: TokenType;
	value: string;
	/** 0-based offset in the source. */
	pos: number;
}

export class QueryError extends Error {
	constructor(message: string, readonly pos: number, readonly source: string) {
		super(`${message} (columna ${pos + 1})`);
	}
}

const IDENT_START = /[A-Za-z_@À-￿]/;
const IDENT_PART = /[A-Za-z0-9_\-@À-￿]/;

export function tokenize(src: string): Token[] {
	const tokens: Token[] = [];
	let i = 0;
	while (i < src.length) {
		const ch = src[i];
		if (/\s/.test(ch)) {
			i++;
			continue;
		}
		const start = i;
		if (ch === '"' || ch === "'") {
			let value = "";
			i++;
			while (i < src.length && src[i] !== ch) {
				if (src[i] === "\\" && i + 1 < src.length) {
					value += src[i + 1];
					i += 2;
				} else {
					value += src[i++];
				}
			}
			if (i >= src.length) throw new QueryError("String sin cerrar", start, src);
			i++;
			tokens.push({ type: "string", value, pos: start });
			continue;
		}
		if (/[0-9]/.test(ch) || (ch === "-" && /[0-9]/.test(src[i + 1] ?? ""))) {
			let j = i + 1;
			while (j < src.length && /[0-9.]/.test(src[j])) j++;
			tokens.push({ type: "number", value: src.slice(i, j), pos: start });
			i = j;
			continue;
		}
		const two = src.slice(i, i + 2);
		if (["==", "!=", ">=", "<=", "&&", "||"].includes(two)) {
			tokens.push({ type: "op", value: two, pos: start });
			i += 2;
			continue;
		}
		if (ch === "!" && IDENT_START.test(src[i + 1] ?? "")) {
			let j = i + 1;
			while (j < src.length && IDENT_PART.test(src[j])) j++;
			const word = src.slice(i, j).toLowerCase();
			if (word !== "!contains" && word !== "!exists") {
				throw new QueryError(`Operador desconocido "${src.slice(i, j)}"`, start, src);
			}
			tokens.push({ type: "op", value: word, pos: start });
			i = j;
			continue;
		}
		if (ch === "!") {
			tokens.push({ type: "op", value: "!", pos: start });
			i++;
			continue;
		}
		if (ch === ">" || ch === "<" || ch === "=") {
			tokens.push({ type: "op", value: ch === "=" ? "==" : ch, pos: start });
			i++;
			continue;
		}
		const single: Record<string, TokenType> = { "(": "lparen", ")": "rparen", "[": "lbracket", "]": "rbracket", ".": "dot", ",": "comma" };
		if (single[ch]) {
			tokens.push({ type: single[ch], value: ch, pos: start });
			i++;
			continue;
		}
		if (IDENT_START.test(ch)) {
			let j = i + 1;
			while (j < src.length && IDENT_PART.test(src[j])) j++;
			tokens.push({ type: "ident", value: src.slice(i, j), pos: start });
			i = j;
			continue;
		}
		throw new QueryError(`Carácter inesperado "${ch}"`, start, src);
	}
	tokens.push({ type: "eof", value: "", pos: src.length });
	return tokens;
}
