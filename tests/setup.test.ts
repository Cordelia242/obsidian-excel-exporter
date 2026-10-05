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
	collectionOptions,
	contextFields,
	pathExpr,
	previewExpression,
	rowSampleScope,
	type FieldNode,
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

describe("field discovery", () => {
	const find = (nodes: FieldNode[], label: string) => nodes.find((n) => n.label === label);

	it("lists root properties, relations and navigable links with samples", () => {
		const cd = definition(BLANK_DEF);
		const run = prepareRun(cd, vault(), { now: NOW });
		const groups = contextFields({ run, perRoot: true });
		expect(groups.map((g) => g.label)).toEqual(["proceso (nota raíz)", "Variables especiales"]);
		const root = groups[0].children!();
		expect(find(root, "status")).toMatchObject({ expr: "proceso.status", kind: "text", sample: "Open" });
		expect(find(root, "quantity")).toMatchObject({ kind: "number" });
		expect(find(root, "created")).toMatchObject({ kind: "date", sample: "2026-03-01" });
		expect(find(root, "role")).toMatchObject({ kind: "link", sample: "Back" });
		const interviews = find(root, "interviews")!;
		expect(interviews).toMatchObject({ expr: "proceso.interviews", kind: "collection", sample: "3 elementos en la muestra" });
		const itemFields = interviews.children!();
		expect(find(itemFields, "tech rating")).toMatchObject({ expr: "proceso.interviews.tech rating", kind: "stars" });
		const interviewed = find(itemFields, "interviewed")!;
		expect(interviewed.kind).toBe("link");
		expect(find(interviewed.children!(), "linkedin")?.expr).toBe("proceso.interviews.interviewed.linkedin");
	});

	it("offers the row element when the row repeats", () => {
		const run = prepareRun(definition(BLANK_DEF), vault(), { now: NOW });
		const groups = contextFields({ run, perRoot: true, rowPath: ["interviews"] });
		expect(groups[0].label).toBe("Fila: cada interviews");
		const fields = groups[0].children!();
		expect(fields[0]).toMatchObject({ expr: "interviews", kind: "link" });
		expect(find(fields, "decision")?.expr).toBe("interviews.decision");
		expect(groups[groups.length - 1].children!()[0].expr).toBe("@index");
	});

	it("collection options depend on the mode", () => {
		const run = prepareRun(definition(BLANK_DEF), vault(), { now: NOW });
		expect(collectionOptions(run, true).map((o) => o.value)).toEqual(["interviews", "hires"]);
		expect(collectionOptions(run, false).map((o) => o.value)).toEqual(["proceso", "proceso.interviews", "proceso.hires"]);
	});

	it("previews an expression against the sample root and row", () => {
		const run = prepareRun(definition(BLANK_DEF), vault(), { now: NOW });
		const base = baseScope(run, true)!;
		expect(previewExpression("proceso.team", base, run)).toBe("Célula Pagos");
		expect(previewExpression("Rol: {{proceso.role | upper}}", base, run)).toBe("Rol: BACK");
		const row = rowSampleScope(base, ["interviews"], run.rt)!;
		expect(previewExpression("interviews.interviewed.seniority", row, run)).toBe("Semi Senior");
		expect(previewExpression("interviews.date", row, run)).toBe("2026-03-03");
	});

	it("pathExpr brackets segments with dots", () => {
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
