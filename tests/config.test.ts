import { parse as parseYaml } from "yaml";
import { describe, expect, it } from "vitest";
import { EXAMPLE_DEFINITION, extractCodeBlocks, parseDefinition } from "../src/config/parse";
import { ConfigError } from "../src/config/schema";
import { definition, PROCESO_DEF } from "./helpers";

const errorsOf = (yaml: string): string[] => {
	try {
		parseDefinition(yaml, (t) => parseYaml(t));
	} catch (e) {
		if (e instanceof ConfigError) return e.errors;
		throw e;
	}
	return [];
};

describe("definition parsing (§5)", () => {
	it("compiles the reference definition", () => {
		const cd = definition(PROCESO_DEF);
		expect(cd.def.mode).toBe("file-per-root");
		expect(cd.relations.map((r) => [r.name, r.kind])).toEqual([
			["interviews", "backlink"],
			["hires", "derived"],
		]);
		expect(cd.relations[0].field).toEqual(["process"]);
		expect(cd.sort).toEqual([{ path: ["created"], desc: true }]);
	});

	it("the bundled example is valid", () => {
		expect(() => definition(EXAMPLE_DEFINITION)).not.toThrow();
	});

	it("orders derived relations after their source", () => {
		const cd = definition(`
name: x
template: t.xlsx
root: { alias: p, where: 'a exists' }
relations:
  b: { source: a, filter: 'x == 1' }
  a: { from: 'y exists', on: 'process -> p' }
`);
		expect(cd.relations.map((r) => r.name)).toEqual(["a", "b"]);
	});

	it("reports missing keys and invalid values clearly", () => {
		expect(errorsOf("name: x")).toEqual([
			"Falta la clave requerida `template`",
			"Falta la clave requerida `root` (con `alias` y `where`)",
		]);
		const errs = errorsOf(`
name: x
template: t.xlsx
root: { alias: "mi alias", where: 'a exists' }
mode: per-root
relations:
  r1: { from: 'a exists' }
  r2: { source: nope }
output: { overwrite: maybe }
`);
		expect(errs).toContain('`root.alias` inválido: "mi alias" (usa letras, números, _ o -)');
		expect(errs).toContain("`mode` inválido: \"per-root\". Valores permitidos: file-per-root, sheet-per-root, single");
		expect(errs).toContain("`relations.r1.`: falta `from` y `on` (o `source` para una relación derivada)");
		expect(errs).toContain('`output.overwrite` inválido: "maybe". Valores permitidos: ask, overwrite, suffix');
	});

	it("reports filter syntax errors with the key and column", () => {
		const errs = errorsOf(`
name: x
template: t.xlsx
root: { alias: p, where: 'status == ' }
relations:
  r: { from: 'a exists', on: 'process -> otro' }
  d: { source: r, sort: 'a..b' }
`);
		expect(errs[0]).toMatch(/^`root.where`: Se esperaba un valor; se encontró fin de la expresión \(columna 11\)/);
		expect(errs.some((e) => e.includes("`relations.r.on`") && e.includes('"otro"'))).toBe(true);
		expect(errs.some((e) => e.startsWith("`relations.d.sort`"))).toBe(true);
	});

	it("detects cycles between derived relations", () => {
		const errs = errorsOf(`
name: x
template: t.xlsx
root: { alias: p, where: 'a exists' }
relations:
  a: { source: b }
  b: { source: a }
`);
		expect(errs.some((e) => e.startsWith("Ciclo entre relaciones derivadas"))).toBe(true);
	});

	it("reports invalid YAML", () => {
		expect(errorsOf("name: [unclosed")[0]).toMatch(/^YAML inválido/);
		expect(errorsOf("- a\n- b")).toEqual(["La definición debe ser un objeto YAML"]);
	});
});

describe("extractCodeBlocks", () => {
	it("finds excel-export fenced blocks only", () => {
		const md = "# Hola\n\n```excel-export\nname: a\n```\n\n```yaml\nx: 1\n```\n\n````excel-export\nname: b\n````\n";
		expect(extractCodeBlocks(md)).toEqual(["name: a\n", "name: b\n"]);
	});
});
