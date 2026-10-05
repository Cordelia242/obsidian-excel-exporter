import { parsePath, PathError } from "../graph/path";
import { QueryError } from "../query/lexer";
import { parseFilter, type Expr } from "../query/parser";
import { parseSort, type SortKey } from "../query/sort";
import type { NormalizeOptions } from "../values/normalize";
import { compileMapping, type CompiledMapping } from "./mapping";

export type ExportMode = "file-per-root" | "sheet-per-root" | "single";
export type OverwriteMode = "ask" | "overwrite" | "suffix";
export type EmptyBlockMode = "remove" | "blank";

export interface RootDef {
	alias: string;
	where: string;
	filter?: string;
	sort?: string;
}

export interface RelationDef {
	/** Backlink relation: candidate notes. */
	from?: string;
	/** Backlink relation: `<field> -> <rootAlias>`. */
	on?: string;
	/** Derived relation: name of another relation. */
	source?: string;
	filter?: string;
	sort?: string;
}

export interface OutputDef {
	folder?: string;
	filename?: string;
	sheetName?: string;
	overwrite?: OverwriteMode;
}

export interface ExportDefinition {
	name: string;
	template: string;
	root: RootDef;
	mode: ExportMode;
	relations: Record<string, RelationDef>;
	output: OutputDef;
	normalize?: Partial<NormalizeOptions>;
	emptyBlock?: EmptyBlockMode;
	/** Setup mapping: `"Hoja!B2": "{{proceso.team}}"` (a bare path is wrapped in braces). */
	cells: Record<string, string>;
	/** Setup mapping: `"Hoja!7": "interviews"` repeats row 7 once per element. */
	rows: Record<string, string>;
}

export interface CompiledRelation {
	name: string;
	kind: "backlink" | "derived";
	from?: Expr;
	/** Backlink field path on the candidate note. */
	field?: string[];
	source?: string;
	filter?: Expr;
	sort: SortKey[];
}

export interface CompiledDefinition {
	def: ExportDefinition;
	where: Expr;
	filter?: Expr;
	sort: SortKey[];
	/** Topologically ordered: sources before derived relations. */
	relations: CompiledRelation[];
	mapping: CompiledMapping;
}

export class ConfigError extends Error {
	constructor(readonly errors: string[]) {
		super(errors.join("\n"));
	}
}

const NAME_RE = /^[A-Za-z_À-￿][A-Za-z0-9_\-À-￿]*$/;
const MODES: ExportMode[] = ["file-per-root", "sheet-per-root", "single"];
const OVERWRITE: OverwriteMode[] = ["ask", "overwrite", "suffix"];

function isObj(v: unknown): v is Record<string, unknown> {
	return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** Validates raw YAML data (§5). Collects every problem instead of stopping at the first. */
export function validateDefinition(raw: unknown): ExportDefinition {
	const errors: string[] = [];
	if (!isObj(raw)) throw new ConfigError(["La definición debe ser un objeto YAML"]);

	const str = (obj: Record<string, unknown>, key: string, where: string, required: boolean): string | undefined => {
		const v = obj[key];
		if (v === undefined || v === null || v === "") {
			if (required) errors.push(`Falta la clave requerida \`${where}${key}\``);
			return undefined;
		}
		if (typeof v !== "string") {
			errors.push(`\`${where}${key}\` debe ser texto`);
			return undefined;
		}
		return v;
	};
	const oneOf = <T extends string>(v: unknown, allowed: T[], key: string): T | undefined => {
		if (v === undefined || v === null) return undefined;
		if (typeof v !== "string" || !allowed.includes(v as T)) {
			errors.push(`\`${key}\` inválido: "${String(v)}". Valores permitidos: ${allowed.join(", ")}`);
			return undefined;
		}
		return v as T;
	};

	const name = str(raw, "name", "", true) ?? "";
	const template = str(raw, "template", "", true) ?? "";

	let root: RootDef = { alias: "", where: "" };
	if (!isObj(raw.root)) {
		errors.push("Falta la clave requerida `root` (con `alias` y `where`)");
	} else {
		const alias = str(raw.root, "alias", "root.", true) ?? "";
		if (alias && !NAME_RE.test(alias)) errors.push(`\`root.alias\` inválido: "${alias}" (usa letras, números, _ o -)`);
		root = {
			alias,
			where: str(raw.root, "where", "root.", true) ?? "",
			filter: str(raw.root, "filter", "root.", false),
			sort: str(raw.root, "sort", "root.", false),
		};
	}

	const mode = oneOf(raw.mode ?? "file-per-root", MODES, "mode") ?? "file-per-root";

	const relations: Record<string, RelationDef> = {};
	if (raw.relations !== undefined && raw.relations !== null) {
		if (!isObj(raw.relations)) {
			errors.push("`relations` debe ser un mapa nombre → relación");
		} else {
			for (const [rname, rdef] of Object.entries(raw.relations)) {
				const where = `relations.${rname}.`;
				if (!NAME_RE.test(rname)) errors.push(`Nombre de relación inválido: "${rname}"`);
				if (rname === root.alias) errors.push(`La relación "${rname}" tiene el mismo nombre que el alias de la raíz`);
				if (!isObj(rdef)) {
					errors.push(`\`relations.${rname}\` debe ser un objeto`);
					continue;
				}
				const r: RelationDef = {
					from: str(rdef, "from", where, false),
					on: str(rdef, "on", where, false),
					source: str(rdef, "source", where, false),
					filter: str(rdef, "filter", where, false),
					sort: str(rdef, "sort", where, false),
				};
				if (r.source) {
					if (r.from || r.on) errors.push(`\`${where}\`: usa \`source\` o \`from\`+\`on\`, no ambos`);
				} else if (!r.from || !r.on) {
					errors.push(`\`${where}\`: falta \`from\` y \`on\` (o \`source\` para una relación derivada)`);
				}
				relations[rname] = r;
			}
		}
	}

	const output: OutputDef = {};
	if (raw.output !== undefined && raw.output !== null) {
		if (!isObj(raw.output)) errors.push("`output` debe ser un objeto");
		else {
			output.folder = str(raw.output, "folder", "output.", false);
			output.filename = str(raw.output, "filename", "output.", false);
			output.sheetName = str(raw.output, "sheetName", "output.", false);
			output.overwrite = oneOf(raw.output.overwrite, OVERWRITE, "output.overwrite");
		}
	}

	let normalize: Partial<NormalizeOptions> | undefined;
	if (raw.normalize !== undefined && raw.normalize !== null) {
		if (!isObj(raw.normalize)) errors.push("`normalize` debe ser un objeto");
		else {
			normalize = {};
			const n = raw.normalize;
			const links = oneOf(n.links, ["display", "target", "raw"], "normalize.links");
			if (links) normalize.links = links;
			const stars = oneOf(n.stars, ["number", "raw"], "normalize.stars");
			if (stars) normalize.stars = stars;
			if (n.listSeparator !== undefined) normalize.listSeparator = String(n.listSeparator);
			if (n.emptyValue !== undefined && n.emptyValue !== null) normalize.emptyValue = String(n.emptyValue);
		}
	}

	const emptyBlock = oneOf(raw.emptyBlock, ["remove", "blank"], "emptyBlock");

	const stringMap = (key: "cells" | "rows"): Record<string, string> => {
		const v = raw[key];
		if (v === undefined || v === null) return {};
		if (!isObj(v)) {
			errors.push(`\`${key}\` debe ser un mapa celda → valor`);
			return {};
		}
		const out: Record<string, string> = {};
		for (const [k, val] of Object.entries(v)) {
			if (val === null || val === undefined || val === "") continue;
			if (typeof val !== "string") errors.push(`\`${key}.${k}\` debe ser texto`);
			else out[k] = val;
		}
		return out;
	};
	const cells = stringMap("cells");
	const rows = stringMap("rows");

	if (errors.length) throw new ConfigError(errors);
	return { name, template, root, mode, relations, output, normalize, emptyBlock, cells, rows };
}

/** Parses all filters/sorts/paths of a validated definition. */
export function compileDefinition(def: ExportDefinition): CompiledDefinition {
	const errors: string[] = [];
	const filter = (src: string | undefined, key: string): Expr | undefined => {
		if (!src) return undefined;
		try {
			return parseFilter(src);
		} catch (e) {
			errors.push(`\`${key}\`: ${e instanceof QueryError ? e.message : String(e)}`);
			return undefined;
		}
	};
	const sort = (src: string | undefined, key: string): SortKey[] => {
		if (!src) return [];
		try {
			return parseSort(src);
		} catch (e) {
			errors.push(`\`${key}\`: ${(e as Error).message}`);
			return [];
		}
	};

	const where = filter(def.root.where, "root.where");
	const rootFilter = filter(def.root.filter, "root.filter");
	const rootSort = sort(def.root.sort, "root.sort");

	const compiled = new Map<string, CompiledRelation>();
	for (const [name, r] of Object.entries(def.relations)) {
		const key = `relations.${name}`;
		if (r.source) {
			if (!def.relations[r.source]) errors.push(`\`${key}.source\`: la relación "${r.source}" no existe`);
			compiled.set(name, {
				name,
				kind: "derived",
				source: r.source,
				filter: filter(r.filter, `${key}.filter`),
				sort: sort(r.sort, `${key}.sort`),
			});
			continue;
		}
		let field: string[] | undefined;
		const m = /^(.+?)\s*->\s*(.+)$/.exec(r.on ?? "");
		if (!m) {
			errors.push(`\`${key}.on\`: formato esperado "<campo> -> ${def.root.alias}"`);
		} else {
			try {
				field = parsePath(m[1].trim());
			} catch (e) {
				errors.push(`\`${key}.on\`: ${e instanceof PathError ? e.message : String(e)}`);
			}
			if (m[2].trim() !== def.root.alias) {
				errors.push(`\`${key}.on\`: en v1 el destino debe ser el alias de la raíz ("${def.root.alias}"), no "${m[2].trim()}"`);
			}
		}
		compiled.set(name, {
			name,
			kind: "backlink",
			from: filter(r.from, `${key}.from`),
			field,
			filter: filter(r.filter, `${key}.filter`),
			sort: sort(r.sort, `${key}.sort`),
		});
	}

	// Topological order, detecting cycles between derived relations.
	const ordered: CompiledRelation[] = [];
	const state = new Map<string, "visiting" | "done">();
	const visit = (name: string, chain: string[]) => {
		const rel = compiled.get(name);
		if (!rel || state.get(name) === "done") return;
		if (state.get(name) === "visiting") {
			errors.push(`Ciclo entre relaciones derivadas: ${[...chain, name].join(" → ")}`);
			return;
		}
		state.set(name, "visiting");
		if (rel.source) visit(rel.source, [...chain, name]);
		state.set(name, "done");
		ordered.push(rel);
	};
	for (const name of compiled.keys()) visit(name, []);

	const mapping = compileMapping(def.cells, def.rows, errors);

	if (errors.length || !where) throw new ConfigError(errors.length ? errors : ["`root.where` es inválido"]);
	return { def, where, filter: rootFilter, sort: rootSort, relations: ordered, mapping };
}
