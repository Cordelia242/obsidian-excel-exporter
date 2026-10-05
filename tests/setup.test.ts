import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { compileMapping, toBareExpression, toTemplateText } from "../src/config/mapping";
import { replaceCodeBlock, extractCodeBlocks, orderDefinitionKeys } from "../src/config/parse";
import { worksheetToGrid } from "../src/excel/grid";
import { loadWorkbook } from "../src/excel/workbook-io";
import { buildExport, prepareRun, runExport, writeExport } from "../src/export/runner";
import { validateTemplate } from "../src/export/validate";
import {
	baseScope,
	cellSources,
	collectionOptions,
	composeExpr,
	fieldsAt,
	formatOptions,
	friendlyLabel,
	parseExprParts,
	pathExpr,
	previewExpression,
	rowSampleScope,
	type FieldEntry,
} from "../src/setup/fields";
import { definition, MAPPING_YAML, NOW, procesoBlankTemplate, procesoTemplate, PROCESO_DEF, readXlsx, vault } from "./helpers";

const BLANK_DEF = PROCESO_DEF.replace("Templates/Excel/Proceso.xlsx", "Templates/Excel/Blank.xlsx") + MAPPING_YAML;

async function blankVault() {
	const adapter = vault();
	adapter.files.set("Templates/Excel/Blank.xlsx", await procesoBlankTemplate());
	return adapter;
}

describe("mapping (Setup)", () => {
	it("compiles cells and rows, wrapping bare paths", () => {
		const errors: string[] = [];
		const m = compileMapping({ "Hoja 1!b2": "proceso.team", "C3": "Rol: {{proceso.role}}" }, { "Hoja 1!7": "interviews" }, errors);
		expect(errors).toEqual([]);
		expect(m.cells[0]).toMatchObject({ sheet: "Hoja 1", address: "B2", row: 2, col: 2, expr: "{{proceso.team}}" });
		expect(m.cells[1]).toMatchObject({ sheet: null, address: "C3", expr: "Rol: {{proceso.role}}" });
		expect(m.rows[0]).toMatchObject({ sheet: "Hoja 1", row: 7, src: "interviews", path: ["interviews"] });
		expect(toTemplateText("a.b | count")).toBe("{{a.b | count}}");
		expect(toBareExpression("{{ a.b }}")).toBe("a.b");
		expect(toBareExpression("x {{a}}")).toBe("x {{a}}");
	});

	it("reports invalid keys and expressions", () => {
		const errors: string[] = [];
		compileMapping({ "Hoja!": "a", "A1": "a..b" }, { "Hoja!B": "x" }, errors);
		expect(errors).toHaveLength(3);
	});

	it("fills a blank template from the mapping, same result as placeholders", async () => {
		const adapter = await blankVault();
		const res = await runExport(definition(BLANK_DEF), adapter, { activeNotePath: "TH/Procesos/Proceso Backend Sr.md", now: NOW });
		const ws = (await readXlsx(adapter, res.files[0])).getWorksheet("Proceso")!;
		expect(ws.getCell("B1").value).toBe("Proceso Backend Sr");
		expect(ws.getCell("B2").value).toBe("Célula Pagos");
		expect(ws.getCell("E4").value).toBe(1);
		expect([7, 8, 9].map((r) => ws.getCell(r, 1).value)).toEqual(["Bruno Díaz", "Ana Torres", "Persona Inexistente"]);
		expect([7, 8, 9].map((r) => ws.getCell(r, 7).value)).toEqual([3, 5, 1]);
		expect((ws.getCell("C9").fill as ExcelJS.FillPattern).fgColor?.argb).toBe("FFDDEBF7");
		expect(ws.getCell("A10").value).toBe("Promedio tech");
		expect((ws.getCell("G10").value as ExcelJS.CellFormulaValue).formula).toBe("AVERAGE(G7:G9)");
		expect(ws.getCell("E8").numFmt).toBe("dd/mm/yyyy");
	});

	it("works in single and sheet-per-root modes", async () => {
		const adapter = await blankVault();
		const single = definition(
			BLANK_DEF.replace("mode: file-per-root", "mode: single")
				.replace("Proceso!7: interviews", "Proceso!7: proceso.interviews")
				.replace("proceso.hires | count", "proceso | count")
				.replace(/filename: .*/, "filename: Todo.xlsx"),
		);
		const s = await runExport(single, adapter, { now: NOW });
		const ws = (await readXlsx(adapter, s.files[0])).getWorksheet("Proceso")!;
		expect(ws.getCell("E4").value).toBe(2);
		expect([7, 8, 9].map((r) => ws.getCell(r, 1).value)).toEqual(["Bruno Díaz", "Ana Torres", "Persona Inexistente"]);

		const sheets = definition(BLANK_DEF.replace("mode: file-per-root", "mode: sheet-per-root").replace(/filename: .*/, "filename: Hojas.xlsx"));
		const h = await runExport(sheets, adapter, { now: NOW });
		const wb = await readXlsx(adapter, h.files[0]);
		expect(wb.worksheets.map((w) => w.name)).toEqual(["Proceso Backend Sr", "Proceso Data"]);
		expect(wb.getWorksheet("Proceso Backend Sr")!.getCell("A9").value).toBe("Persona Inexistente");
	});

	it("a {{#each}} in the cell wins over a row repeat (with warning)", async () => {
		const adapter = vault();
		adapter.files.set("Templates/Excel/Proceso.xlsx", await procesoTemplate());
		const res = await runExport(definition(PROCESO_DEF + "rows:\n  Proceso!7: hires\n"), adapter, {
			activeNotePath: "TH/Procesos/Proceso Backend Sr.md",
			now: NOW,
		});
		expect(res.warnings.some((w) => w.message.includes("se ignora la repetición"))).toBe(true);
		const ws = (await readXlsx(adapter, res.files[0])).getWorksheet("Proceso")!;
		expect(ws.getCell("A9").value).toBe("Persona Inexistente");
	});

	it("validates mapped cells", async () => {
		const adapter = await blankVault();
		const v = await validateTemplate(definition(BLANK_DEF), adapter, { now: NOW });
		expect(v.eachRows.map((e) => [e.collection, e.status, e.count])).toEqual([["interviews", "ok", 3]]);
		expect(v.placeholders.every((p) => p.status === "ok")).toBe(true);
	});
});

describe("buildExport + writeExport (preview before writing)", () => {
	it("builds in memory without writing, then writes", async () => {
		const adapter = await blankVault();
		const before = adapter.files.size;
		const built = await buildExport(definition(BLANK_DEF), adapter, { now: NOW });
		expect(adapter.files.size).toBe(before);
		expect(built.outputs.map((o) => o.fileName)).toEqual(["Proceso Backend Sr - 2026-10-05.xlsx", "Proceso Data - 2026-10-05.xlsx"]);
		const grid = worksheetToGrid(built.outputs[0].workbook.getWorksheet("Proceso")!);
		expect(grid.rows[6].cells[0].text).toBe("Bruno Díaz");
		const res = await writeExport(built, adapter, {});
		expect(res.files).toHaveLength(2);
	});
});

describe("worksheetToGrid", () => {
	it("renders values, styles, widths, merges and formulas", async () => {
		const adapter = await blankVault();
		const built = await buildExport(definition(BLANK_DEF), adapter, { activeNotePath: "TH/Procesos/Proceso Backend Sr.md", now: NOW });
		const g = worksheetToGrid(built.outputs[0].workbook.worksheets[0], { minRows: 20, minCols: 10 });
		expect(g.sheet).toBe("Proceso");
		expect(g.rows).toHaveLength(20);
		expect(g.cols).toHaveLength(10);
		expect(g.cols[3]).toMatchObject({ letter: "D", width: Math.round(34 * 7 + 5) });
		const cell = (a: string) => g.rows.flatMap((r) => r.cells).find((c) => c.address === a)!;
		expect(cell("B1")).toMatchObject({ text: "Proceso Backend Sr", colSpan: 3, hidden: false });
		expect(cell("C1").hidden).toBe(true);
		expect(cell("A1").style).toMatchObject({ bold: true, size: 14 });
		expect(cell("A6").style).toMatchObject({ bold: true, color: "#FFFFFF", bg: "#1F4E78" });
		expect(cell("E7")).toMatchObject({ text: "03/03/2026", kind: "date" });
		expect(cell("G7")).toMatchObject({ text: "3", kind: "number" });
		expect(cell("G10")).toMatchObject({ text: "=AVERAGE(G7:G9)", kind: "formula" });
		expect(cell("C7").style.borders?.bottom).toBe(true);
		expect(g.rows[6].height).toBe(Math.round(22 * 4 / 3));
	});
});

describe("field browser (Setup)", () => {
	const run = () => prepareRun(definition(BLANK_DEF), vault(), { now: NOW });
	const byLabel = (entries: FieldEntry[], label: string) => entries.find((e) => e.label === label);

	it("lists sources: the root, each relation, and dates", () => {
		const r = run();
		const sources = cellSources({ run: r, perRoot: true });
		expect(sources.map((x) => [x.label, x.base.join("."), x.isList])).toEqual([
			["Proceso", "proceso", false],
			["Interviews (todas)", "proceso.interviews", true],
			["Hires (todas)", "proceso.hires", true],
			["Fecha y otros", "", false],
		]);
	});

	it("in a repeated row the row element comes first", () => {
		const r = run();
		const sources = cellSources({ run: r, perRoot: true, rowPath: ["interviews"] });
		expect(sources[0]).toMatchObject({ label: "Interviews de esta fila", base: ["interviews"], isList: false, kind: "row" });
		expect(sources[0].notes).toHaveLength(3);
		const fields = fieldsAt(sources[0], [], { run: r, perRoot: true, rowPath: ["interviews"] });
		expect(fields[0]).toMatchObject({ label: "Nombre de la nota", segs: ["interviews"], kind: "note" });
		expect(byLabel(fields, "decision")).toMatchObject({ segs: ["interviews", "decision"], kind: "text", isList: false });
		expect(byLabel(fields, "tech rating")?.kind).toBe("stars");
		const interviewed = byLabel(fields, "interviewed")!;
		expect(interviewed).toMatchObject({ kind: "link", navigable: true });
		const persona = fieldsAt(sources[0], ["interviewed"], { run: r, perRoot: true, rowPath: ["interviews"] });
		expect(byLabel(persona, "linkedin")).toMatchObject({ segs: ["interviews", "interviewed", "linkedin"], kind: "text" });
		expect(byLabel(persona, "stack")?.kind).toBe("list");
		const special = fieldsAt(sources[sources.length - 1], [], { run: r, perRoot: true, rowPath: ["interviews"] });
		expect(special.map((e) => e.segs[0])).toEqual(["@index", "@today", "@now", "@exportName", "@count"]);
	});

	it("root fields include relations, properties with examples and file info", () => {
		const r = run();
		const [root] = cellSources({ run: r, perRoot: true });
		const fields = fieldsAt(root, [], { run: r, perRoot: true });
		expect(byLabel(fields, "Interviews")).toMatchObject({ kind: "collection", segs: ["proceso", "interviews"], isList: true });
		expect(byLabel(fields, "status")).toMatchObject({ kind: "text", sample: "Open" });
		expect(byLabel(fields, "created")?.kind).toBe("date");
		expect(byLabel(fields, "Fecha de creación")?.segs).toEqual(["proceso", "file", "ctime"]);
	});

	it("single mode: all roots are a list, rows bind the root", () => {
		const r = prepareRun(definition(BLANK_DEF.replace("mode: file-per-root", "mode: single")), vault(), { now: NOW });
		expect(cellSources({ run: r, perRoot: false })[0]).toMatchObject({ label: "Todos los Proceso", isList: true });
		const inRow = cellSources({ run: r, perRoot: false, rowPath: ["proceso", "interviews"] });
		expect(inRow.map((x) => x.label)).toEqual(["Interviews de esta fila", "Proceso", "Interviews (todas)", "Hires (todas)", "Fecha y otros"]);
	});

	it("format options depend on the kind of field", () => {
		expect(formatOptions("collection", true).map((f) => f.pipe)).toEqual(["", "count", "first", 'join:"; "']);
		expect(formatOptions("date", false).map((f) => f.pipe)).toEqual(["", 'date:"dd/MM/yyyy"', 'date:"yyyy-MM-dd"']);
		expect(formatOptions("stars", false).map((f) => f.pipe)).toEqual(["", "raw"]);
		expect(formatOptions("link", false).map((f) => f.pipe)).toEqual(["", "target", "link"]);
		expect(formatOptions("number", false)).toHaveLength(1);
	});

	it("expressions <-> parts and friendly labels", () => {
		const def = definition(BLANK_DEF).def;
		expect(parseExprParts('proceso.hires | count | default:"-"')).toEqual({ segs: ["proceso", "hires"], format: "count", fallback: "-" });
		expect(parseExprParts("Rol: {{proceso.role}}")).toBeNull();
		expect(composeExpr({ segs: ["interviews", "tech rating"], format: "", fallback: "" })).toBe("interviews.tech rating");
		expect(friendlyLabel("proceso.hires | count", def)).toBe("Proceso › Hires (cantidad)");
		expect(friendlyLabel("interviews.interviewed.linkedin", def)).toBe("Interviews › interviewed › linkedin");
		expect(friendlyLabel("proceso.file.ctime", def)).toBe("Proceso › Fecha de creación");
		expect(friendlyLabel("@today", def)).toBe("Fecha de hoy");
		expect(friendlyLabel("Rol: {{proceso.role}}", def)).toBe("Rol: [Proceso › role]");
	});

	it("collection options with counts for the sample", () => {
		const r = run();
		expect(collectionOptions(r, true)).toEqual([
			{ value: "interviews", label: "Interviews", count: 3 },
			{ value: "hires", label: "Hires", count: 1 },
		]);
	});

	it("previews an expression against the sample root and row", () => {
		const r = run();
		const base = baseScope(r, true)!;
		expect(previewExpression("proceso.team", base, r)).toBe("Célula Pagos");
		expect(previewExpression("Rol: {{proceso.role | upper}}", base, r)).toBe("Rol: BACK");
		const row = rowSampleScope(base, ["interviews"], r.rt)!;
		expect(previewExpression("interviews.interviewed.seniority", row, r)).toBe("Semi Senior");
		expect(pathExpr(["a", "tech rating", "x.y"])).toBe('a.tech rating["x.y"]');
	});
});

describe("replaceCodeBlock", () => {
	it("replaces only the requested block", () => {
		const md = "intro\n\n```excel-export\nname: a\n```\n\n```excel-export\nname: b\n```\nfin\n";
		const out = replaceCodeBlock(md, 1, "name: c\ncells: {}");
		expect(extractCodeBlocks(out)).toEqual(["name: a\n", "name: c\ncells: {}\n"]);
		expect(out.startsWith("intro\n")).toBe(true);
		expect(out.endsWith("fin\n")).toBe(true);
		expect(() => replaceCodeBlock(md, 5, "x")).toThrow();
	});

	it("orders definition keys", () => {
		expect(Object.keys(orderDefinitionKeys({ rows: {}, cells: {}, name: "x", extra: 1, template: "t" }))).toEqual([
			"name", "template", "cells", "rows", "extra",
		]);
	});
});

describe("blank template round-trip", () => {
	it("blank template has no placeholders", async () => {
		const wb = await loadWorkbook(await procesoBlankTemplate());
		const texts = worksheetToGrid(wb.worksheets[0]).rows.flatMap((r) => r.cells.map((c) => c.text));
		expect(texts.some((t) => t.includes("{{"))).toBe(false);
	});
});
