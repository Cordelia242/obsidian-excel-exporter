import { compileDefinition, ConfigError, validateDefinition, type CompiledDefinition } from "../config/schema";

/**
 * A template as stored in the plugin data: the same shape as the YAML
 * definition (so the engine is shared), plus an id.
 */
export interface TemplateConfig {
	id: string;
	name: string;
	description?: string;
	template: string;
	root: { alias: string; label?: string; where: string; filter?: string; sort?: string };
	mode: "file-per-root" | "sheet-per-root" | "single";
	relations: Record<string, { label?: string; from?: string; on?: string; source?: string; filter?: string; sort?: string }>;
	output: { folder?: string; filename?: string; sheetName?: string; repeatSheet?: string; overwrite?: "ask" | "overwrite" | "suffix" };
	cells: Record<string, string>;
	rows: Record<string, string>;
	normalize?: Record<string, unknown>;
	emptyBlock?: "remove" | "blank";
	createdAt?: number;
	updatedAt?: number;
}

export function newId(): string {
	return `t${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
}

export function newTemplateConfig(name: string, templatePath: string, outputFolder = "Exports/out"): TemplateConfig {
	const now = Date.now();
	return {
		id: newId(),
		name,
		template: templatePath,
		root: { alias: "nota", label: "Nota", where: "" },
		mode: "file-per-root",
		relations: {},
		output: { folder: outputFolder, filename: "{{nota.file.name}}.xlsx", overwrite: "ask" },
		cells: {},
		rows: {},
		createdAt: now,
		updatedAt: now,
	};
}

/** Raw definition object (what validateDefinition expects). */
export function toRawDefinition(cfg: TemplateConfig): Record<string, unknown> {
	const { id: _id, createdAt: _c, updatedAt: _u, ...raw } = cfg;
	return JSON.parse(JSON.stringify(raw)) as Record<string, unknown>;
}

export interface CompileResult {
	cd?: CompiledDefinition;
	errors: string[];
}

export function compileTemplate(cfg: TemplateConfig): CompileResult {
	try {
		return { cd: compileDefinition(validateDefinition(toRawDefinition(cfg))), errors: [] };
	} catch (e) {
		return { errors: e instanceof ConfigError ? e.errors : [(e as Error).message] };
	}
}

/** Builds a stored template from a YAML definition found in a note. */
export function fromRawDefinition(raw: Record<string, unknown>): TemplateConfig {
	const def = validateDefinition(raw);
	const now = Date.now();
	return JSON.parse(
		JSON.stringify({
			id: newId(),
			name: def.name,
			description: def.description,
			template: def.template,
			root: def.root,
			mode: def.mode,
			relations: def.relations,
			output: def.output,
			cells: def.cells,
			rows: def.rows,
			normalize: def.normalize,
			emptyBlock: def.emptyBlock,
			createdAt: now,
			updatedAt: now,
		}),
	) as TemplateConfig;
}

export function duplicateTemplate(cfg: TemplateConfig): TemplateConfig {
	const copy = JSON.parse(JSON.stringify(cfg)) as TemplateConfig;
	copy.id = newId();
	copy.name = `${cfg.name} (copia)`;
	copy.createdAt = copy.updatedAt = Date.now();
	return copy;
}

/** Problems to finish before a template can export, phrased for users. */
export function templateStatus(cfg: TemplateConfig): { ready: boolean; message: string } {
	if (!cfg.template) return { ready: false, message: "Falta elegir el archivo Excel" };
	if (!cfg.root.where.trim()) return { ready: false, message: "Falta elegir qué notas exporta (paso 1)" };
	if (!Object.keys(cfg.cells).length) return { ready: false, message: "Falta indicar qué dato va en cada celda (paso 3)" };
	const { errors } = compileTemplate(cfg);
	if (errors.length) return { ready: false, message: errors[0] };
	return { ready: true, message: "" };
}

function escapeRe(s: string): string {
	return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const END = "(?=$|[.\\[\\s|}])";

/** Renames a path variable (root alias or relation) in every expression of the template. */
export function renameVariable(cfg: TemplateConfig, oldName: string, newName: string): void {
	if (!oldName || oldName === newName) return;
	const head = new RegExp(`(^\\s*|\\{\\{\\s*|#each\\s+)${escapeRe(oldName)}${END}`, "g");
	const seg = new RegExp(`(\\.)${escapeRe(oldName)}${END}`, "g");
	const fix = (s: string | undefined) => (s === undefined ? s : s.replace(head, `$1${newName}`).replace(seg, `$1${newName}`));
	for (const k of Object.keys(cfg.cells)) cfg.cells[k] = fix(cfg.cells[k]) as string;
	for (const k of Object.keys(cfg.rows)) cfg.rows[k] = fix(cfg.rows[k]) as string;
	cfg.output.filename = fix(cfg.output.filename);
	cfg.output.sheetName = fix(cfg.output.sheetName);
	for (const r of Object.values(cfg.relations)) {
		if (r.on) r.on = r.on.replace(new RegExp(`->\\s*${escapeRe(oldName)}\\s*$`), `-> ${newName}`);
		if (r.source === oldName) r.source = newName;
	}
	if (cfg.root.alias === oldName) cfg.root.alias = newName;
	if (cfg.relations[oldName]) {
		const entries = Object.entries(cfg.relations).map(([k, v]) => [k === oldName ? newName : k, v] as const);
		cfg.relations = Object.fromEntries(entries);
	}
}

/** A free identifier based on a label (`Contratados` → `contratados`, `contratados_2`…). */
export function uniqueName(cfg: TemplateConfig, label: string, slug: (s: string) => string): string {
	const base = slug(label) || "datos";
	const taken = new Set([cfg.root.alias, ...Object.keys(cfg.relations)].map((x) => x.toLowerCase()));
	let name = base;
	for (let i = 2; taken.has(name.toLowerCase()); i++) name = `${base}_${i}`;
	return name;
}
