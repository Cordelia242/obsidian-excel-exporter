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
			const { cd } = compileTemplate(cfg);
			const res = await runExport(cd!, adapter, { now: NOW });
			expect(res.files.length).toBeGreaterThan(0);
			const ws = (await readXlsx(adapter, res.files[0])).worksheets[0];
			expect(ws.getCell("A2").value).not.toBeNull();
		});
	}
});
