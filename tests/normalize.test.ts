import { describe, expect, it } from "vitest";
import { LinkRef, parseWikilink, toLinkRef, wrapFrontmatterValue } from "../src/values/link";
import {
	countStars,
	DEFAULT_NORMALIZE,
	formatDate,
	MISSING,
	normalizeValue,
	parseIsoDate,
	type NormalizeOptions,
} from "../src/values/normalize";

const opts = (o: Partial<NormalizeOptions> = {}): NormalizeOptions => ({ ...DEFAULT_NORMALIZE, ...o });
const norm = (v: unknown, o: Partial<NormalizeOptions> = {}) => normalizeValue(wrapFrontmatterValue(v, "src.md"), opts(o));

describe("wikilinks", () => {
	it("parses links with alias, folder and subpath", () => {
		expect(parseWikilink("[[Backend]]")).toEqual({ linkpath: "Backend", subpath: null, alias: null });
		expect(parseWikilink("[[Celula X|Alias]]")).toEqual({ linkpath: "Celula X", subpath: null, alias: "Alias" });
		expect(parseWikilink("[[Carpeta/Nota#Sección|A]]")).toEqual({ linkpath: "Carpeta/Nota", subpath: "#Sección", alias: "A" });
		expect(parseWikilink("texto [[X]]")).toBeNull();
		expect(parseWikilink("[[]]")).toBeNull();
	});

	it("wraps frontmatter values, including the unquoted YAML [[X]] quirk", () => {
		expect(wrapFrontmatterValue("[[X]]", "a.md")).toBeInstanceOf(LinkRef);
		const list = wrapFrontmatterValue(["[[A]]", "b"], "a.md") as unknown[];
		expect(list[0]).toBeInstanceOf(LinkRef);
		expect(list[1]).toBe("b");
		const quirk = wrapFrontmatterValue([["Backend"]], "a.md") as LinkRef;
		expect(quirk.targetName).toBe("Backend");
	});
});

describe("normalizeValue (§8)", () => {
	it("links", () => {
		expect(norm("[[Backend]]")).toBe("Backend");
		expect(norm("[[Celula X|Alias]]")).toBe("Alias");
		expect(norm("[[Celula X|Alias]]", { links: "target" })).toBe("Celula X");
		expect(norm("[[Celula X|Alias]]", { links: "raw" })).toBe("[[Celula X|Alias]]");
		expect(norm("[[Carpeta/Nota]]")).toBe("Nota");
	});

	it("inline links inside free text", () => {
		expect(norm("Trabajó con [[Ana Torres]] y [[Equipo/Pagos|Pagos]]")).toBe("Trabajó con Ana Torres y Pagos");
	});

	it("lists", () => {
		expect(norm(["[[TypeScript]]", "[[Node]]"])).toBe("TypeScript, Node");
		expect(norm(["a", "b"], { listSeparator: "; " })).toBe("a; b");
		expect(norm(["[[Solo]]"])).toBe("Solo");
		expect(norm([3])).toBe(3);
		expect(norm([])).toBeNull();
		expect(norm(["", null])).toBeNull();
	});

	it("dates are local, without timezone shift", () => {
		const d = norm("2026-03-01") as Date;
		expect(d).toBeInstanceOf(Date);
		expect([d.getFullYear(), d.getMonth(), d.getDate(), d.getHours()]).toEqual([2026, 2, 1, 0]);
		const dt = parseIsoDate("2026-03-01T14:05:09")!;
		expect([dt.getHours(), dt.getMinutes(), dt.getSeconds()]).toEqual([14, 5, 9]);
		expect(parseIsoDate("2026-02-30")).toBeNull();
		expect(parseIsoDate("hola")).toBeNull();
		expect(norm(["2026-03-01", "2026-03-02"])).toBe("2026-03-01, 2026-03-02");
	});

	it("stars", () => {
		expect(norm("⭐⭐⭐")).toBe(3);
		expect(norm("⭐️⭐️")).toBe(2);
		expect(norm("⭐⭐⭐", { stars: "raw" })).toBe("⭐⭐⭐");
		expect(countStars("⭐⭐½")).toBe(2.5);
		expect(countStars("3 estrellas")).toBeNull();
	});

	it("native types and empties", () => {
		expect(norm(5)).toBe(5);
		expect(norm(true)).toBe(true);
		expect(norm(null)).toBeNull();
		expect(norm(undefined)).toBeNull();
		expect(norm("")).toBeNull();
		expect(normalizeValue(MISSING, opts())).toBeNull();
		expect(norm(null, { emptyValue: "-" })).toBe("-");
	});

	it("objects become compact JSON with a warning", () => {
		const warnings: string[] = [];
		expect(normalizeValue({ a: 1 }, opts(), (m) => warnings.push(m))).toBe('{"a":1}');
		expect(warnings).toHaveLength(1);
	});
});

describe("formatDate", () => {
	it("supports common tokens", () => {
		const d = new Date(2026, 2, 1, 9, 5, 7);
		expect(formatDate(d, "dd/MM/yyyy")).toBe("01/03/2026");
		expect(formatDate(d, "yyyy-MM-dd HH:mm:ss")).toBe("2026-03-01 09:05:07");
		expect(formatDate(d, "d/M/yy")).toBe("1/3/26");
	});
});

describe("toLinkRef", () => {
	it("keeps the source path for resolution", () => {
		expect(toLinkRef("[[X]]", "a/b.md")?.sourcePath).toBe("a/b.md");
	});
});
