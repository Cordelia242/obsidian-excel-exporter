import { readFileSync } from "fs";
import { describe, expect, it } from "vitest";
import { runExport } from "../src/export/runner";
import { compileTemplate, templateStatus, type TemplateConfig } from "../src/store/templates";
import { NOW, readXlsx, vault } from "./helpers";

const read = (p: string) => {
	const b = readFileSync(p);
	return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};

const templates = JSON.parse(readFileSync("examples/templates.json", "utf8")) as TemplateConfig[];

describe("examples/templates.json", () => {
	for (const cfg of templates) {
		it(`${cfg.name} is ready and exports`, async () => {
			expect(templateStatus(cfg)).toEqual({ ready: true, message: "" });
			const adapter = vault();
			adapter.files.set("Templates/Excel/Proceso.xlsx", read("examples/Templates/Excel/Proceso.xlsx"));
			adapter.files.set("Templates/Excel/Consolidado.xlsx", read("examples/Templates/Excel/Consolidado.xlsx"));
			adapter.files.set(
				"Templates/Excel/Ficha de proceso completa.xlsx",
				read("examples/Templates/Excel/Ficha de proceso completa.xlsx"),
			);
			const { cd } = compileTemplate(cfg);
			const res = await runExport(cd!, adapter, { now: NOW });
			expect(res.files.length).toBeGreaterThan(0);
			// The first mapped cell of each template must get data.
			const [key] = Object.keys(cfg.cells).filter((k) => k !== "Proceso!C2");
			const [sheet, address] = key.split("!");
			const ws = (await readXlsx(adapter, res.files[0])).getWorksheet(sheet)!;
			expect(ws.getCell(address).value).not.toBeNull();
		});
	}
});

describe("Ficha de proceso completa", () => {
	const cfg = templates.find((t) => t.id === "ficha-completa") as TemplateConfig;
	const formula = (v: unknown) => (v as { formula: string }).formula;

	async function exportFor(path: string) {
		const adapter = vault();
		adapter.files.set(cfg.template, read(`examples/${cfg.template}`));
		const { cd } = compileTemplate(cfg);
		const res = await runExport(cd!, adapter, { now: NOW, rootPaths: [path] });
		return { wb: await readXlsx(adapter, res.files[0]), res };
	}

	it("fills the details sheet and one row per candidate, adjusting cross-sheet formulas", async () => {
		const { wb } = await exportFor("TH/Procesos/Proceso Backend Sr.md");
		const p = wb.getWorksheet("Proceso")!;
		expect(p.getCell("C5").value).toBe("Proceso Backend Sr");
		expect(p.getCell("F5").value).toBe("Open");
		expect(p.getCell("C6").value).toBe("Célula Pagos");
		expect(p.getCell("F6").value).toBe(5);
		expect(p.getCell("C7").value).toBe("Back");
		expect(p.getCell("F7").value).toBe(2);
		expect(p.getCell("F8").value).toBeInstanceOf(Date);
		expect(p.getCell("B15").value).toBeNull();
		expect(formula(p.getCell("C11").value)).toBe("COUNTA(Postulantes!B5:B7)");
		expect(formula(p.getCell("F11").value)).toBe('COUNTIF(Postulantes!L5:L7,"Hire")');
		expect(formula(p.getCell("C12").value)).toBe('IFERROR(AVERAGE(Postulantes!K5:K7),"-")');

		const s = wb.getWorksheet("Postulantes")!;
		expect(s.getCell("B2").value).toBe("Proceso:");
		expect(s.getCell("C2").value).toBe("Proceso Backend Sr");
		expect([5, 6, 7].map((r) => s.getCell(r, 1).value)).toEqual([1, 2, 3]);
		expect([5, 6, 7].map((r) => s.getCell(r, 2).value)).toEqual(["Bruno Díaz", "Ana Torres", "Persona Inexistente"]);
		expect(s.getCell("E6").value).toBe("TypeScript, Node");
		expect(s.getCell("F6").value).toBe("Bogotá");
		expect([5, 6, 7].map((r) => s.getCell(r, 11).value)).toEqual([3, 5, 1]);
		expect(formula(s.getCell("K8").value)).toBe('IFERROR(AVERAGE(K5:K7),"-")');
	});

	it("keeps an empty candidates row (and valid formulas) when there are no interviews", async () => {
		const { wb } = await exportFor("TH/Procesos/Proceso Data.md");
		expect(formula(wb.getWorksheet("Proceso")!.getCell("C11").value)).toBe("COUNTA(Postulantes!B5:B5)");
		expect(wb.getWorksheet("Postulantes")!.getCell("B5").value).toBeNull();
	});
});
