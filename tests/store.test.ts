import { describe, expect, it } from "vitest";
import { conditionsToExpr, exprToConditions, fieldToText } from "../src/setup/conditions";
import { noteTypeSuggestions, propertiesOf, relationSuggestions, slugify, valuesOfProperty } from "../src/setup/discovery";
import { compileTemplate, fromRawDefinition, newTemplateConfig, renameVariable, templateStatus, uniqueName } from "../src/store/templates";
import { runExport } from "../src/export/runner";
import { parse as parseYaml } from "yaml";
import { MAPPING_YAML, NOW, PROCESO_DEF, procesoBlankTemplate, readXlsx, runtime, vault } from "./helpers";

describe("conditions <-> filter text", () => {
	it("builds filter text", () => {
		expect(
			conditionsToExpr([
				{ field: "categories", op: "contains", value: "Talent Acquisition Process" },
				{ field: "tech rating", op: ">=", value: "3" },
				{ field: "status", op: "!=", value: 'Cl"osed' },
				{ field: "linkedin", op: "exists", value: "ignored" },
				{ field: "", op: "==", value: "x" },
			]),
		).toBe('categories contains "Talent Acquisition Process" and ["tech rating"] >= 3 and status != "Cl\\"osed" and linkedin exists');
		expect(fieldToText("file.name")).toBe("file.name");
	});

	it("parses simple filters back, refuses complex ones", () => {
		expect(exprToConditions('categories contains "TH" and ["tech rating"] >= 3 and x !exists')).toEqual([
			{ field: "categories", op: "contains", value: "TH" },
			{ field: "tech rating", op: ">=", value: "3" },
			{ field: "x", op: "!exists", value: "" },
		]);
		expect(exprToConditions("status == Closed")).toEqual([{ field: "status", op: "==", value: "Closed" }]);
		expect(exprToConditions("")).toEqual([]);
		expect(exprToConditions("a == 1 or b == 2")).toBeNull();
		expect(exprToConditions("created > today")).toBeNull();
		expect(exprToConditions("((")).toBeNull();
	});

	it("round-trips", () => {
		const src = 'categories contains "Interview Evaluations" and decision == "Hire"';
		expect(conditionsToExpr(exprToConditions(src)!)).toBe(src);
	});
});

describe("discovery", () => {
	const notes = vault().listNotes();

	it("suggests note types by categories and folders", () => {
		const s = noteTypeSuggestions(notes);
		const tap = s.find((x) => x.label === "Talent Acquisition Process")!;
		expect(tap).toMatchObject({ count: 3, condition: { field: "categories", op: "contains", value: "Talent Acquisition Process" } });
		expect(tap.examples).toContain("Proceso Backend Sr");
		expect(s.find((x) => x.label === "People")?.count).toBe(3);
	});

	it("describes properties and their values", () => {
		const procesos = notes.filter((n) => n.path.startsWith("TH/Procesos"));
		const props = propertiesOf(procesos);
		expect(props.find((p) => p.key === "status")).toMatchObject({ kind: "text", count: 3 });
		expect(props.find((p) => p.key === "role")).toMatchObject({ kind: "link", sample: "Backend" });
		expect(props.find((p) => p.key === "categories")?.kind).toBe("list");
		expect(props.find((p) => p.key === "created")?.kind).toBe("date");
		expect(valuesOfProperty(procesos, "status").map((v) => v.value).sort()).toEqual(["Closed", "On hold", "Open"]);
		const interviews = notes.filter((n) => n.path.includes("Interviews"));
		expect(propertiesOf(interviews).find((p) => p.key === "tech rating")?.kind).toBe("stars");
	});

	it("detects notes pointing to the roots (relation suggestions)", () => {
		const rt = runtime();
		const roots = notes.filter((n) => n.path.startsWith("TH/Procesos"));
		const [s] = relationSuggestions(notes, roots, rt);
		expect(s).toMatchObject({
			field: "process",
			count: 4,
			label: "Interview Evaluations",
			name: "interview_evaluations",
			from: { field: "categories", op: "contains", value: "Interview Evaluations" },
		});
		expect(slugify("Contratados ñandú 2")).toBe("contratados_nandu_2");
		expect(slugify("2x")).toBe("r_2x");
	});
});

describe("template store", () => {
	it("new templates report what is missing", () => {
		const cfg = newTemplateConfig("Ficha", "Templates/Excel/Blank.xlsx");
		expect(templateStatus(cfg)).toEqual({ ready: false, message: "Falta elegir qué notas exporta (paso 1)" });
		cfg.root.where = 'categories contains "Talent Acquisition Process"';
		expect(templateStatus(cfg).message).toMatch(/paso 3/);
		cfg.cells["Proceso!B1"] = "nota.file.name";
		expect(templateStatus(cfg).ready).toBe(true);
	});

	it("renames the root alias and relations everywhere", () => {
		const cfg = fromRawDefinition(parseYaml(PROCESO_DEF + MAPPING_YAML));
		renameVariable(cfg, "proceso", "p");
		renameVariable(cfg, "interviews", "entrevistas");
		expect(cfg.root.alias).toBe("p");
		expect(cfg.cells["Proceso!B2"]).toBe("p.team");
		expect(cfg.cells["Proceso!E4"]).toBe("p.hires | count");
		expect(cfg.cells["Proceso!A7"]).toBe("entrevistas.interviewed");
		expect(cfg.cells["Proceso!B1"]).toBe("{{p.file.name}}");
		expect(cfg.rows["Proceso!7"]).toBe("entrevistas");
		expect(cfg.output.filename).toBe('{{p.file.name}} - {{@today | date:"yyyy-MM-dd"}}.xlsx');
		expect(Object.keys(cfg.relations)).toEqual(["entrevistas", "hires"]);
		expect(cfg.relations.entrevistas.on).toBe("process -> p");
		expect(cfg.relations.hires.source).toBe("entrevistas");
		expect(uniqueName(cfg, "Hires", slugify)).toBe("hires_2");
		expect(compileTemplate(cfg).errors).toEqual([]);
	});

	it("a stored template exports exactly the picked notes", async () => {
		const adapter = vault();
		adapter.files.set("Templates/Excel/Blank.xlsx", await procesoBlankTemplate());
		const cfg = fromRawDefinition(parseYaml(PROCESO_DEF.replace("Templates/Excel/Proceso.xlsx", "Templates/Excel/Blank.xlsx") + MAPPING_YAML));
		const { cd } = compileTemplate(cfg);
		const res = await runExport(cd!, adapter, {
			now: NOW,
			rootPaths: ["TH/Procesos/Proceso Frontend.md", "TH/Procesos/Proceso Backend Sr.md"],
		});
		expect(res.files).toEqual(["Exports/out/Proceso Frontend - 2026-10-05.xlsx", "Exports/out/Proceso Backend Sr - 2026-10-05.xlsx"]);
		const ws = (await readXlsx(adapter, res.files[0])).worksheets[0];
		expect(ws.getCell("A7").value).toBe("Carla Ruiz");
	});
});
