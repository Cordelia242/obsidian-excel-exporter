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
