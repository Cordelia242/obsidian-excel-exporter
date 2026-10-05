import { compileDefinition, ConfigError, validateDefinition, type CompiledDefinition } from "./schema";

export type YamlParser = (text: string) => unknown;

export const CODEBLOCK_LANG = "excel-export";

/** Parses + validates + compiles the YAML of an `excel-export` block. Throws {@link ConfigError}. */
export function parseDefinition(text: string, parseYaml: YamlParser): CompiledDefinition {
	let raw: unknown;
	try {
		raw = parseYaml(text);
	} catch (e) {
		throw new ConfigError([`YAML inválido: ${(e as Error).message}`]);
	}
	return compileDefinition(validateDefinition(raw));
}

/** Returns the bodies of every ```excel-export fenced block in a markdown document. */
export function extractCodeBlocks(markdown: string, lang = CODEBLOCK_LANG): string[] {
	const blocks: string[] = [];
	const re = /^([ \t]*)(`{3,}|~{3,})[ \t]*([\w-]+)[^\n]*\n([\s\S]*?)^\1\2[ \t]*$/gm;
	let m: RegExpExecArray | null;
	while ((m = re.exec(markdown))) {
		if (m[3] === lang) blocks.push(m[4]);
	}
	return blocks;
}

export const EXAMPLE_DEFINITION = `name: Procesos activos
# Excel con títulos y formato; los datos se configuran con el botón Setup.
template: Templates/Excel/Proceso.xlsx

root:
  alias: proceso
  where: 'categories contains "Talent Acquisition Process"'
  filter: 'status != "Closed"'
  sort: 'created desc'

mode: file-per-root

relations:
  interviews:
    from: 'categories contains "Interview Evaluations"'
    on: 'process -> proceso'
    sort: 'date asc'
  hires:
    source: interviews
    filter: 'decision == "Hire"'

output:
  folder: Exports/out
  filename: '{{proceso.file.name}} - {{@today | date:"yyyy-MM-dd"}}.xlsx'
  overwrite: ask
`;

/** Replaces the body of the `index`-th ```excel-export block of a markdown document. */
export function replaceCodeBlock(markdown: string, index: number, body: string, lang = CODEBLOCK_LANG): string {
	const re = /^([ \t]*)(`{3,}|~{3,})[ \t]*([\w-]+)[^\n]*\n([\s\S]*?)^\1\2[ \t]*$/gm;
	let m: RegExpExecArray | null;
	let i = 0;
	while ((m = re.exec(markdown))) {
		if (m[3] !== lang) continue;
		if (i++ === index) {
			const bodyStart = m.index + m[0].indexOf("\n") + 1;
			const bodyEnd = bodyStart + m[4].length;
			const text = body.endsWith("\n") ? body : body + "\n";
			return markdown.slice(0, bodyStart) + text + markdown.slice(bodyEnd);
		}
	}
	throw new Error(`No se encontró el bloque ${lang} #${index + 1}`);
}

/** Canonical key order when writing a definition back to YAML. */
export const DEFINITION_KEY_ORDER = ["name", "template", "root", "mode", "relations", "output", "normalize", "emptyBlock", "cells", "rows"];

export function orderDefinitionKeys(raw: Record<string, unknown>): Record<string, unknown> {
	const out: Record<string, unknown> = {};
	for (const k of DEFINITION_KEY_ORDER) if (raw[k] !== undefined) out[k] = raw[k];
	for (const k of Object.keys(raw)) if (!(k in out)) out[k] = raw[k];
	return out;
}
