import { readdirSync, readFileSync } from "fs";
import { parse as parseYaml } from "yaml";
import { describe, expect, it } from "vitest";
import { extractCodeBlocks, parseDefinition } from "../src/config/parse";
import { runExport } from "../src/export/runner";
import { NOW, readXlsx, vault } from "./helpers";

const read = (p: string) => {
	const b = readFileSync(p);
	return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};

describe("examples/", () => {
	for (const file of readdirSync("examples/Exports")) {
		it(`${file} runs against the example templates`, async () => {
			const adapter = vault();
			adapter.files.set("Templates/Excel/Proceso.xlsx", read("examples/Templates/Excel/Proceso.xlsx"));
			adapter.files.set("Templates/Excel/Consolidado.xlsx", read("examples/Templates/Excel/Consolidado.xlsx"));
			const [source] = extractCodeBlocks(readFileSync(`examples/Exports/${file}`, "utf8"));
			const cd = parseDefinition(source, (t) => parseYaml(t));
			const res = await runExport(cd, adapter, { now: NOW });
			expect(res.files.length).toBeGreaterThan(0);
			if (cd.mapping.cells.length) {
				const ws = (await readXlsx(adapter, res.files[0])).worksheets[0];
				expect(ws.getCell("A2").value).not.toBeNull();
			}
		});
	}
});
