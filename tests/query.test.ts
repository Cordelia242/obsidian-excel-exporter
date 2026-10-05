import { describe, expect, it } from "vitest";
import { makeScope } from "../src/graph/context";
import { makeNote } from "../src/model/note-record";
import { evaluate } from "../src/query/evaluator";
import { QueryError, tokenize } from "../src/query/lexer";
import { parseFilter } from "../src/query/parser";
import { parseSort, sortNotes } from "../src/query/sort";
import { note, runtime, vault } from "./helpers";

const rt = runtime();
const n = (fm: Record<string, unknown>) => makeNote("x/Nota.md", fm);
const check = (expr: string, fm: Record<string, unknown>) => evaluate(parseFilter(expr), n(fm), rt);

describe("lexer/parser", () => {
	it("tokenizes operators and strings", () => {
		const types = tokenize('a.b != "x" and not c !contains 3').map((t) => t.type);
		expect(types).toEqual(["ident", "dot", "ident", "op", "string", "ident", "ident", "ident", "op", "number", "eof"]);
	});

	it("parses precedence: not > and > or", () => {
		const e = parseFilter("a == 1 or b == 2 and not c exists");
		expect(e.type).toBe("or");
		if (e.type === "or") expect(e.right.type).toBe("and");
	});

	it("parses bracketed paths with spaces", () => {
		const e = parseFilter('["tech rating"] == "⭐"');
		expect(e).toMatchObject({ type: "cmp", path: ["tech rating"] });
		expect(parseFilter('file.name == "x"')).toMatchObject({ path: ["file", "name"] });
	});

	it("reports errors with column and token", () => {
		expect(() => parseFilter('status = = "x"')).toThrow(QueryError);
		try {
			parseFilter('status == "Closed" and');
		} catch (e) {
			expect((e as Error).message).toMatch(/columna 23/);
			expect((e as Error).message).toMatch(/fin de la expresión/);
		}
		try {
			parseFilter("status ~ 1");
		} catch (e) {
			expect((e as Error).message).toMatch(/Carácter inesperado "~".*columna 8/);
		}
		expect(() => parseFilter('a == "sin cerrar')).toThrow(/String sin cerrar/);
		expect(() => parseFilter("(a == 1")).toThrow(/Se esperaba "\)"/);
		expect(() => parseFilter("a")).toThrow(/operador/);
	});
});

describe("semantics (§7)", () => {
	it("links compare by destination basename, case-insensitive", () => {
		expect(check('role == "Backend"', { role: "[[Backend]]" })).toBe(true);
		expect(check('role == "backend"', { role: "[[Roles/Backend|Back]]" })).toBe(true);
		expect(check('role == "Back"', { role: "[[Roles/Backend|Back]]" })).toBe(false);
		expect(check('role == "[[Backend]]"', { role: "[[Roles/Backend]]" })).toBe(true);
		expect(check('role != "Backend"', { role: "[[Frontend]]" })).toBe(true);
	});

	it("lists: contains = membership, == any element", () => {
		const fm = { categories: ["[[Talent Acquisition Process]]", "[[TH]]"] };
		expect(check('categories contains "Talent Acquisition Process"', fm)).toBe(true);
		expect(check('categories contains "Talent"', fm)).toBe(false);
		expect(check('categories == "TH"', fm)).toBe(true);
		expect(check('categories !contains "People"', fm)).toBe(true);
		expect(check('tags contains "a"', { tags: ["a", "b"] })).toBe(true);
	});

	it("strings: contains is a case-insensitive substring", () => {
		expect(check('status contains "clo"', { status: "Closed" })).toBe(true);
		expect(check('status == "closed"', { status: "Closed" })).toBe(true);
		expect(check('status !contains "open"', { status: "Closed" })).toBe(true);
	});

	it("numbers and booleans", () => {
		expect(check("priority >= 5", { priority: 5 })).toBe(true);
		expect(check("priority > 5", { priority: 5 })).toBe(false);
		expect(check("priority < 10 and priority > -1", { priority: 5 })).toBe(true);
		expect(check("IsTechLead == true", { IsTechLead: true })).toBe(true);
		expect(check("IsTechLead == false", { IsTechLead: true })).toBe(false);
	});

	it("dates: ISO strings, date(...) and today", () => {
		expect(check('created >= "2026-01-01"', { created: "2026-03-01" })).toBe(true);
		expect(check('created < date("2026-03-01")', { created: "2026-02-28" })).toBe(true);
		expect(check('created == "2026-03-01"', { created: "2026-03-01" })).toBe(true);
		expect(check("created < today", { created: "2026-10-04" })).toBe(true);
		expect(check("created == today", { created: "2026-10-05" })).toBe(true);
		expect(check("created > today", { created: "2026-10-05" })).toBe(false);
	});

	it("missing fields: false except != and !exists", () => {
		expect(check('status == "x"', {})).toBe(false);
		expect(check('status contains "x"', {})).toBe(false);
		expect(check("priority > 1", {})).toBe(false);
		expect(check('status != "Closed"', {})).toBe(true);
		expect(check("status !exists", {})).toBe(true);
		expect(check("status exists", {})).toBe(false);
		expect(check("status exists", { status: "" })).toBe(false);
		expect(check("status == null", {})).toBe(true);
		expect(check("status != null", { status: "x" })).toBe(true);
	});

	it("properties with spaces, keys case-insensitive, file pseudo-props", () => {
		expect(check('["tech rating"] == "⭐⭐"', { "tech rating": "⭐⭐" })).toBe(true);
		expect(check('Status == "Open"', { status: "Open" })).toBe(true);
		expect(check('file.name == "Nota"', {})).toBe(true);
		expect(check('file.folder == "x"', {})).toBe(true);
	});

	it("unquoted words are text unless they are aliases", () => {
		expect(check("status == Closed", { status: "Closed" })).toBe(true);
	});

	it("alias references compare by resolved file", () => {
		const adapter = vault();
		const r = runtime(adapter);
		const proceso = note(adapter, "Proceso Backend Sr");
		const interview = note(adapter, "Interview Ana Torres - Proceso Backend Sr");
		const scope = makeScope({ proceso });
		expect(evaluate(parseFilter("process == proceso"), interview, r, scope)).toBe(true);
		const other = note(adapter, "Proceso Frontend");
		expect(evaluate(parseFilter("process == proceso"), interview, r, makeScope({ proceso: other }))).toBe(false);
	});

	it("traverses links in filter paths", () => {
		const adapter = vault();
		const r = runtime(adapter);
		const interview = note(adapter, "Interview Ana Torres - Proceso Backend Sr");
		expect(evaluate(parseFilter('interviewed.seniority == "Experimentado"'), interview, r)).toBe(true);
	});
});

describe("sort", () => {
	it("sorts by multiple keys with empties last", () => {
		const notes = [
			makeNote("a.md", { p: 1, c: "2026-01-02" }),
			makeNote("b.md", { p: 3 }),
			makeNote("c.md", { c: "2026-01-01" }),
			makeNote("d.md", { p: 3, c: "2026-01-01" }),
		];
		const sorted = sortNotes(notes, parseSort("p desc, c asc"), rt).map((x) => x.basename);
		expect(sorted).toEqual(["d", "b", "a", "c"]);
		expect(parseSort('["tech rating"] desc')).toEqual([{ path: ["tech rating"], desc: true }]);
		expect(parseSort("tech rating")).toEqual([{ path: ["tech rating"], desc: false }]);
	});
});
