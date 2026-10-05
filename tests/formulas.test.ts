import { describe, expect, it } from "vitest";
import { copyMapper, insertionMapper, mapFormulaRefs } from "../src/excel/formulas";

const ins = (f: string, at: number, d: number, sheet = "S") => mapFormulaRefs(f, sheet, "S", insertionMapper(at, d));

describe("insertionMapper", () => {
	it("expands ranges that contain the template row", () => {
		expect(ins("AVERAGE(G7:G7)", 7, 2)).toBe("AVERAGE(G7:G9)");
		expect(ins("SUM(F5:F11)", 7, 3)).toBe("SUM(F5:F14)");
		expect(ins("SUM(F7:F11)", 7, 3)).toBe("SUM(F7:F14)");
	});

	it("shifts references below the template row, keeps the ones above", () => {
		expect(ins("A8+B6+$C$9", 7, 2)).toBe("A10+B6+$C$11");
		expect(ins("A7*2", 7, 2)).toBe("A7*2");
		expect(ins("SUM(A1:A5)", 7, 2)).toBe("SUM(A1:A5)");
	});

	it("handles removal of the template row", () => {
		expect(ins("AVERAGE(G7:G7)", 7, -1)).toBe("AVERAGE(#REF!)");
		expect(ins("SUM(G6:G8)", 7, -1)).toBe("SUM(G6:G7)");
		expect(ins("A9", 7, -1)).toBe("A8");
		expect(ins("A7", 7, -1)).toBe("#REF!");
	});

	it("ignores strings, function names and other sheets", () => {
		expect(ins('IF(A8>0,"A8","")', 7, 1)).toBe('IF(A9>0,"A8","")');
		expect(ins("LOG10(A8)", 7, 1)).toBe("LOG10(A9)");
		expect(ins("Otra!A8+A8", 7, 1)).toBe("Otra!A8+A9");
		expect(ins("S!A8+'S'!B8", 7, 1, "Otra")).toBe("S!A9+'S'!B9");
		expect(ins("A8", 7, 1, "Otra")).toBe("A8");
		expect(ins("'Mi hoja'!A8", 7, 1)).toBe("'Mi hoja'!A8");
	});
});

describe("copyMapper", () => {
	it("moves relative refs with the copy, like Excel copy/paste", () => {
		const m = (f: string, i: number) => mapFormulaRefs(f, "S", "S", copyMapper(7, 2, i));
		expect(m("F7*G7", 0)).toBe("F7*G7");
		expect(m("F7*G7", 2)).toBe("F9*G9");
		expect(m("F7*$G$7", 1)).toBe("F8*$G$7");
		expect(m("H6+F7", 1)).toBe("H7+F8");
		expect(m("$A$10", 1)).toBe("$A$12");
	});
});
