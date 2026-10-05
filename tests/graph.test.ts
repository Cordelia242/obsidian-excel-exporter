import { describe, expect, it } from "vitest";
import { makeScope } from "../src/graph/context";
import { parsePath, resolvePath } from "../src/graph/path";
import { RelationBuilder } from "../src/graph/relations";
import { LinkRef } from "../src/values/link";
import { MISSING, normalizeValue, DEFAULT_NORMALIZE } from "../src/values/normalize";
import { definition, note, PROCESO_DEF, runtime, vault } from "./helpers";

describe("parsePath (§6.2)", () => {
	it("splits on dots and keeps spaces", () => {
		expect(parsePath("interviews.tech rating")).toEqual(["interviews", "tech rating"]);
		expect(parsePath(" proceso . file.name ")).toEqual(["proceso", "file", "name"]);
		expect(parsePath('proceso["mi.prop"]')).toEqual(["proceso", "mi.prop"]);
		expect(parsePath('proceso["mi.prop"].x')).toEqual(["proceso", "mi.prop", "x"]);
		expect(parsePath("@index")).toEqual(["@index"]);
		expect(() => parsePath("a..b")).toThrow();
		expect(() => parsePath("")).toThrow();
		expect(() => parsePath('a["x')).toThrow();
	});
});

function setup() {
	const adapter = vault();
	const rt = runtime(adapter);
	const cd = definition(PROCESO_DEF);
	const builder = new RelationBuilder(cd.relations, adapter.listNotes(), "proceso", rt);
	const proceso = note(adapter, "Proceso Backend Sr");
	const rels = builder.build(proceso);
	const scope = makeScope({ ...Object.fromEntries(rels), proceso });
	const get = (p: string) => resolvePath(parsePath(p), scope, rt);
	const show = (p: string) => normalizeValue(get(p), DEFAULT_NORMALIZE);
	return { adapter, rt, builder, rels, get, show };
}

describe("relations (§5, §3)", () => {
	it("backlink relation returns the process interviews sorted by date", () => {
		const { rels } = setup();
		expect(rels.get("interviews")!.map((n) => n.basename)).toEqual([
			"Interview Bruno Díaz - Proceso Backend Sr",
			"Interview Ana Torres - Proceso Backend Sr",
			"Interview Fantasma - Proceso Backend Sr",
		]);
	});

	it("derived relation filters its source", () => {
		const { rels } = setup();
		expect(rels.get("hires")!.map((n) => n.basename)).toEqual(["Interview Ana Torres - Proceso Backend Sr"]);
	});

	it("each process gets only its own interviews", () => {
		const { adapter, builder } = setup();
		const fe = builder.build(note(adapter, "Proceso Frontend"));
		expect(fe.get("interviews")!.map((n) => n.basename)).toEqual(["Interview Carla Ruiz - Proceso Frontend"]);
		expect(builder.build(note(adapter, "Proceso Data")).get("interviews")).toEqual([]);
	});

	it("traverses forward links: interviews.interviewed.<prop>", () => {
		const { show, rt } = setup();
		expect(show("interviews.interviewed")).toBe("Bruno Díaz, Ana Torres, Persona Inexistente");
		expect(show("interviews.interviewed.seniority")).toBe("Semi Senior, Experimentado");
		expect(show("interviews.interviewed.linkedin")).toBe(
			"https://linkedin.com/in/bruno-diaz, https://linkedin.com/in/ana-torres",
		);
		expect(rt.report.warnings.some((w) => w.kind === "broken-link" && w.message.includes("Persona Inexistente"))).toBe(true);
	});

	it("maps over lists of links and properties with spaces", () => {
		const { show } = setup();
		expect(show("interviews.people")).toBe("Luis Pérez, Luis Pérez, Marta Gómez");
		expect(show("interviews.tech rating")).toBe("3, 5, 1");
		expect(show("hires.interviewed.past workplaces")).toBe("Acme, Globex");
		expect(show("hires.interviewed.stack")).toBe("TypeScript, Node");
	});

	it("exposes file pseudo-properties and relations on the root", () => {
		const { show, get } = setup();
		expect(show("proceso.file.name")).toBe("Proceso Backend Sr");
		expect(show("proceso.file.path")).toBe("TH/Procesos/Proceso Backend Sr.md");
		expect(show("proceso.file.folder")).toBe("TH/Procesos");
		expect(String(show("proceso.file.link"))).toBe(
			"obsidian://open?vault=Contrataci%C3%B3n&file=TH%2FProcesos%2FProceso%20Backend%20Sr.md",
		);
		expect((get("proceso.interviews") as unknown[]).length).toBe(3);
		expect(show("proceso.TEAM")).toBe("Célula Pagos");
		expect(get("proceso.role")).toBeInstanceOf(LinkRef);
	});

	it("returns MISSING for unknown paths", () => {
		const { get } = setup();
		expect(get("nada")).toBe(MISSING);
		expect(get("proceso.nada")).toBe(MISSING);
		expect(get("proceso.status.x")).toBe(MISSING);
	});
});
