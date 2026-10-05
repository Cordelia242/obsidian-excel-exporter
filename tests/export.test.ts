import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { runExport } from "../src/export/runner";
import { validateTemplate } from "../src/export/validate";
import { consolidadoTemplate, definition, NOW, procesoTemplate, PROCESO_DEF, readXlsx, toBuffer, vault } from "./helpers";

async function procesoVault() {
	const adapter = vault();
	adapter.files.set("Templates/Excel/Proceso.xlsx", await procesoTemplate());
	return adapter;
}

const merges = (ws: ExcelJS.Worksheet) => ((ws.model as unknown as { merges: string[] }).merges ?? []).slice().sort();
const formula = (c: ExcelJS.Cell) => (c.value as ExcelJS.CellFormulaValue).formula;
const utcDay = (d: Date) => [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];

describe("file-per-root (§14.1)", () => {
	it("exports the active note with scalars, one row per interview and adjusted formulas", async () => {
		const adapter = await procesoVault();
		const cd = definition(PROCESO_DEF);
		const res = await runExport(cd, adapter, { activeNotePath: "TH/Procesos/Proceso Backend Sr.md", now: NOW });
		expect(res.files).toEqual(["Exports/out/Proceso Backend Sr - 2026-10-05.xlsx"]);

		const wb = await readXlsx(adapter, res.files[0]);
		const ws = wb.getWorksheet("Proceso")!;
		expect(ws.getCell("B1").value).toBe("Proceso Backend Sr");
		expect(ws.getCell("B2").value).toBe("Célula Pagos");
		expect(ws.getCell("E2").value).toBe("Back");
		expect(ws.getCell("B3").value).toBe("Experimentado");
		expect(ws.getCell("E3").value).toBe(2);
		expect(ws.getCell("B4").value).toBe("Open");
		expect(ws.getCell("E4").value).toBe(1);

		// rows 7-9: interviews sorted by date asc
		expect([7, 8, 9].map((r) => ws.getCell(r, 1).value)).toEqual(["Bruno Díaz", "Ana Torres", "Persona Inexistente"]);
		expect(ws.getCell("B8").value).toBe("Experimentado");
		expect(ws.getCell("C8").value).toBe("TypeScript, Node");
		expect(ws.getCell("C7").value).toBe("Java");
		expect(ws.getCell("D8").value).toBe("https://linkedin.com/in/ana-torres");
		expect(ws.getCell("B9").value).toBeNull();
		const date = ws.getCell("E8").value as Date;
		expect(date).toBeInstanceOf(Date);
		expect(utcDay(date)).toEqual([2026, 3, 5]);
		expect(ws.getCell("E8").numFmt).toBe("dd/mm/yyyy");
		expect([7, 8, 9].map((r) => ws.getCell(r, 6).value)).toEqual([2, 3, 1]);
		expect([7, 8, 9].map((r) => ws.getCell(r, 7).value)).toEqual([3, 5, 1]);
		expect([7, 8, 9].map((r) => ws.getCell(r, 8).value)).toEqual(["Pending", "Hire", "Discard"]);

		// formulas below the block
		expect(ws.getCell("A10").value).toBe("Promedio tech");
		expect(formula(ws.getCell("G10"))).toBe("AVERAGE(G7:G9)");
		expect(formula(ws.getCell("G11"))).toBe("COUNTA(H7:H9)+G10*0");
		expect(ws.getCell("A12").value).toBe("Generado 05/10/2026");

		// styles, heights, widths and merges preserved
		for (const r of [7, 8, 9]) {
			const c = ws.getCell(r, 3);
			expect((c.fill as ExcelJS.FillPattern).fgColor?.argb).toBe("FFDDEBF7");
			expect(c.border?.bottom?.style).toBe("thin");
			expect(ws.getRow(r).height).toBe(22);
		}
		expect(ws.getCell("A6").font?.bold).toBe(true);
		expect(ws.getColumn(4).width).toBe(34);
		expect(merges(ws)).toEqual(["A10:B10", "B1:D1"]);

		// warnings: broken link, persona without linkedin...
		expect(res.warnings.some((w) => w.kind === "broken-link" && w.message.includes("Persona Inexistente"))).toBe(true);
		expect(res.warnings.some((w) => w.kind === "unresolved")).toBe(true);
	});

	it("exports every root matching where+filter (Closed excluded), sorted", async () => {
		const adapter = await procesoVault();
		const res = await runExport(definition(PROCESO_DEF), adapter, { now: NOW });
		expect(res.roots).toBe(2);
		expect(res.files).toEqual([
			"Exports/out/Proceso Backend Sr - 2026-10-05.xlsx",
			"Exports/out/Proceso Data - 2026-10-05.xlsx",
		]);
		const ws = (await readXlsx(adapter, res.files[1])).getWorksheet("Proceso")!;
		// Proceso Data has no interviews: the #each row is removed and formulas collapse.
		expect(ws.getCell("A7").value).toBe("Promedio tech");
		expect(formula(ws.getCell("G7"))).toBe("AVERAGE(#REF!)");
		expect(ws.getCell("E4").value).toBe(0);
		expect(merges(ws)).toEqual(["A7:B7", "B1:D1"]);
		expect(res.warnings.some((w) => w.kind === "formula" && w.location === "Proceso!G7")).toBe(true);
	});

	it("keeps an empty #each row with emptyBlock: blank", async () => {
		const adapter = await procesoVault();
		const cd = definition(PROCESO_DEF + "emptyBlock: blank\n");
		const res = await runExport(cd, adapter, { activeNotePath: "TH/Procesos/Proceso Data.md", now: NOW });
		const ws = (await readXlsx(adapter, res.files[0])).getWorksheet("Proceso")!;
		expect(ws.getCell("A7").value).toBeNull();
		expect(formula(ws.getCell("G8"))).toBe("AVERAGE(G7:G7)");
	});

	it("handles existing files according to overwrite", async () => {
		const adapter = await procesoVault();
		const cd = definition(PROCESO_DEF);
		const opts = { activeNotePath: "TH/Procesos/Proceso Backend Sr.md", now: NOW };
		await runExport(cd, adapter, opts);
		const second = await runExport(cd, adapter, opts);
		expect(second.files).toEqual(["Exports/out/Proceso Backend Sr - 2026-10-05 (2).xlsx"]);
		const asked: string[] = [];
		const ask = definition(PROCESO_DEF.replace("overwrite: suffix", "overwrite: ask"));
		const third = await runExport(ask, adapter, {
			...opts,
			confirmOverwrite: async (p) => {
				asked.push(p);
				return "skip";
			},
		});
		expect(asked).toEqual(["Exports/out/Proceso Backend Sr - 2026-10-05.xlsx"]);
		expect(third.files).toEqual([]);
		const over = definition(PROCESO_DEF.replace("overwrite: suffix", "overwrite: overwrite"));
		expect((await runExport(over, adapter, opts)).files).toEqual(["Exports/out/Proceso Backend Sr - 2026-10-05.xlsx"]);
	});

	it("warns (but exports) when the active note does not match root.where", async () => {
		const adapter = await procesoVault();
		const res = await runExport(definition(PROCESO_DEF), adapter, { activeNotePath: "People/Ana Torres.md", now: NOW });
		expect(res.files).toHaveLength(1);
		expect(res.warnings.some((w) => w.kind === "skipped-note")).toBe(true);
	});

	it("fails clearly when the template is missing", async () => {
		await expect(runExport(definition(PROCESO_DEF), vault(), { now: NOW })).rejects.toThrow(/No se encontró el template/);
	});
});

describe("sheet-per-root (§6.4)", () => {
	it("clones the first sheet per root with valid, unique names and keeps the other sheets", async () => {
		const adapter = vault([
			{
				path: "TH/Procesos/Proceso: QA/Lead?.md",
				frontmatter: { categories: ["[[Talent Acquisition Process]]"], status: "Open", created: "2025-01-01" },
			},
		]);
		const wb = new ExcelJS.Workbook();
		wb.addWorksheet("Resumen").getCell("A1").value = "Procesos: {{proceso | count}}";
		await (async () => {
			const tpl = new ExcelJS.Workbook();
			await tpl.xlsx.load(await procesoTemplate());
			const src = tpl.getWorksheet("Proceso")!;
			const dst = wb.addWorksheet("Proceso");
			dst.model = { ...src.model, mergeCells: (src.model as unknown as { merges: string[] }).merges, name: "Proceso" } as never;
		})();
		// Template order: Proceso first, Resumen second.
		(wb.getWorksheet("Proceso") as unknown as { orderNo: number }).orderNo = 0;
		const buf = await toBuffer(wb);
		adapter.files.set("Templates/Excel/Proceso.xlsx", buf);

		const cd = definition(
			PROCESO_DEF.replace("mode: file-per-root", "mode: sheet-per-root").replace(
				"filename: '{{proceso.file.name}} - {{@today | date:\"yyyy-MM-dd\"}}.xlsx'",
				"filename: 'Procesos.xlsx'",
			),
		);
		const res = await runExport(cd, adapter, { now: NOW });
		expect(res.files).toEqual(["Exports/out/Procesos.xlsx"]);
		const out = await readXlsx(adapter, res.files[0]);
		const names = out.worksheets.map((w) => w.name);
		expect(names).toEqual(["Proceso Backend Sr", "Proceso Data", "Lead", "Resumen"]);
		const be = out.getWorksheet("Proceso Backend Sr")!;
		expect(be.getCell("B1").value).toBe("Proceso Backend Sr");
		expect(be.getCell("A9").value).toBe("Persona Inexistente");
		expect(formula(be.getCell("G10"))).toBe("AVERAGE(G7:G9)");
		expect((be.getCell("C7").fill as ExcelJS.FillPattern).fgColor?.argb).toBe("FFDDEBF7");
		expect(be.getColumn(4).width).toBe(34);
		expect(merges(be)).toEqual(["A10:B10", "B1:D1"]);
		const data = out.getWorksheet("Proceso Data")!;
		expect(data.getCell("A7").value).toBe("Promedio tech");
		expect(out.getWorksheet("Resumen")!.getCell("A1").value).toBe("Procesos: 3");
	});
});

describe("single + flatten (§14.2)", () => {
	it("produces one row per interview of active processes", async () => {
		const adapter = vault();
		adapter.files.set("Templates/Excel/Consolidado.xlsx", await consolidadoTemplate());
		const cd = definition(`
name: Postulantes en procesos activos
template: Templates/Excel/Consolidado.xlsx
root:
  alias: proceso
  where: 'categories contains "Talent Acquisition Process"'
  filter: 'status != "Closed"'
  sort: 'priority desc, created asc'
mode: single
relations:
  interviews:
    from: 'categories contains "Interview Evaluations"'
    on: 'process -> proceso'
output:
  folder: Exports/out
  filename: 'Consolidado {{@today | date:"yyyy-MM-dd"}}.xlsx'
`);
		const res = await runExport(cd, adapter, { now: NOW });
		expect(res.files).toEqual(["Exports/out/Consolidado 2026-10-05.xlsx"]);
		const ws = (await readXlsx(adapter, res.files[0])).getWorksheet("Consolidado")!;
		const rows = [2, 3, 4].map((r) => [1, 2, 3, 4, 5, 6, 7].map((c) => ws.getCell(r, c).value));
		// Proceso Data (priority 8) has no interviews; Backend Sr interviews in note order.
		expect(rows.map((r) => r[0])).toEqual(["Proceso Backend Sr", "Proceso Backend Sr", "Proceso Backend Sr"]);
		expect(rows.map((r) => r[1])).toEqual(["Back", "Back", "Back"]);
		expect(rows.map((r) => r[2])).toEqual([5, 5, 5]);
		expect(rows.map((r) => r[3]).sort()).toEqual(["Ana Torres", "Bruno Díaz", "Persona Inexistente"]);
		expect(rows.map((r) => r[6])).toEqual(["1/3", "2/3", "3/3"]);
		expect(ws.getCell("A5").value).toBe("Total");
		expect(formula(ws.getCell("E5"))).toBe("SUM(E2:E4)");
	});

	it("iterates roots with {{#each proceso}} and exposes their relations", async () => {
		const adapter = vault();
		const wb = new ExcelJS.Workbook();
		const ws = wb.addWorksheet("Lista");
		ws.getCell("A1").value = "{{#each proceso}}{{proceso.file.name}}";
		ws.getCell("B1").value = "{{interviews | count}}";
		ws.getCell("C1").value = "{{proceso.created}}";
		const buf = await toBuffer(wb);
		adapter.files.set("t.xlsx", buf);
		const cd = definition(`
name: Lista
template: t.xlsx
root: { alias: proceso, where: 'categories contains "Talent Acquisition Process"', sort: 'created asc' }
mode: single
relations:
  interviews: { from: 'categories contains "Interview Evaluations"', on: 'process -> proceso' }
`);
		const res = await runExport(cd, adapter, { now: NOW });
		expect(res.files).toEqual(["Exports/out/Lista.xlsx"]);
		const out = (await readXlsx(adapter, res.files[0])).getWorksheet("Lista")!;
		expect([1, 2, 3].map((r) => out.getCell(r, 1).value)).toEqual(["Proceso Frontend", "Proceso Data", "Proceso Backend Sr"]);
		expect([1, 2, 3].map((r) => out.getCell(r, 2).value)).toEqual([1, 0, 3]);
		expect(utcDay(out.getCell("C1").value as Date)).toEqual([2026, 1, 15]);
		expect(out.getCell("C1").numFmt).toBe("yyyy-mm-dd");
	});
});

describe("pipes (§6.5)", () => {
	it("applies raw, target, join, first, default, upper, lower, date and link", async () => {
		const adapter = vault();
		const wb = new ExcelJS.Workbook();
		const ws = wb.addWorksheet("P");
		const cells: Record<string, string> = {
			A1: "{{proceso.role | raw}}",
			A2: "{{proceso.role | target}}",
			A3: '{{interviews.interviewed | join:" / "}}',
			A4: "{{interviews.interviewed | first}}",
			A5: '{{proceso.nada | default:"-"}}',
			A6: "{{proceso.status | upper}}",
			A7: "{{proceso.team | lower}}",
			A8: '{{proceso.created | date:"dd/MM/yyyy"}}',
			A9: "{{proceso | link}}",
			A10: "{{interviews.date | first}}",
			A11: "Rol: {{proceso.role}} ({{proceso.quantity}})",
			A12: "{{proceso.categories}}",
			A13: "{{@exportName}}",
			A14: "{{proceso.status | nope}}",
		};
		for (const [a, v] of Object.entries(cells)) ws.getCell(a).value = v;
		const buf = await toBuffer(wb);
		adapter.files.set("t.xlsx", buf);
		const cd = definition(`
name: Pipes
template: t.xlsx
root: { alias: proceso, where: 'file.name == "Proceso Backend Sr"' }
relations:
  interviews: { from: 'categories contains "Interview Evaluations"', on: 'process -> proceso', sort: 'date asc' }
`);
		const res = await runExport(cd, adapter, { now: NOW });
		const out = (await readXlsx(adapter, res.files[0])).getWorksheet("P")!;
		const v = (a: string) => out.getCell(a).value;
		expect(v("A1")).toBe("[[Roles/Backend|Back]]");
		expect(v("A2")).toBe("Backend");
		expect(v("A3")).toBe("Bruno Díaz / Ana Torres / Persona Inexistente");
		expect(v("A4")).toBe("Bruno Díaz");
		expect(v("A5")).toBe("-");
		expect(v("A6")).toBe("OPEN");
		expect(v("A7")).toBe("célula pagos");
		expect(v("A8")).toBe("01/03/2026");
		expect(v("A9")).toEqual({
			text: "Proceso Backend Sr",
			hyperlink: "obsidian://open?vault=Contrataci%C3%B3n&file=TH%2FProcesos%2FProceso%20Backend%20Sr.md",
		});
		expect(utcDay(v("A10") as Date)).toEqual([2026, 3, 3]);
		expect(v("A11")).toBe("Rol: Back (2)");
		expect(v("A12")).toBe("Talent Acquisition Process, TH");
		expect(v("A13")).toBe("Pipes");
		expect(res.warnings.some((w) => w.kind === "pipe" && w.message.includes("nope"))).toBe(true);
	});
});

describe("validateTemplate (§9)", () => {
	it("reports placeholders, #each rows and unknown collections", async () => {
		const adapter = await procesoVault();
		const tpl = new ExcelJS.Workbook();
		await tpl.xlsx.load(await procesoTemplate());
		const ws = tpl.getWorksheet("Proceso")!;
		ws.getCell("A20").value = "{{#each candidatos}}{{candidatos.x}}";
		ws.getCell("A21").value = "{{proceso.inexistente}}";
		ws.getCell("A22").value = "{{proceso.status | bogus}}";
		const buf = await toBuffer(tpl);
		adapter.files.set("Templates/Excel/Proceso.xlsx", buf);
		const v = await validateTemplate(definition(PROCESO_DEF), adapter, { now: NOW });
		expect(v.roots).toBe(2);
		expect(v.sampleRoot).toBe("TH/Procesos/Proceso Backend Sr.md");
		expect(v.eachRows.map((e) => [e.address, e.collection, e.status, e.count])).toEqual([
			["A7", "interviews", "ok", 3],
			["A20", "candidatos", "unknown", undefined],
		]);
		const byAddr = Object.fromEntries(v.placeholders.map((p) => [p.address, p.status]));
		expect(byAddr.B1).toBe("ok");
		expect(byAddr.D7).toBe("ok");
		expect(byAddr.A21).toBe("unresolved");
		expect(byAddr.A22).toBe("error");
		expect(byAddr.A20).toBe("no-sample");
	});
});
