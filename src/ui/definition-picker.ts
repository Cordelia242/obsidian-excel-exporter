import { FuzzySuggestModal, normalizePath, parseYaml, TFile, type App, type FuzzyMatch } from "obsidian";
import { extractCodeBlocks, parseDefinition } from "../config/parse";
import { ConfigError, type CompiledDefinition } from "../config/schema";

export interface DefinitionEntry {
	file: TFile;
	/** Index of the block inside the note (a note may hold several). */
	index: number;
	name: string;
	compiled?: CompiledDefinition;
	error?: string;
}

export function errorText(e: unknown): string {
	return e instanceof ConfigError ? e.errors.join("\n") : e instanceof Error ? e.message : String(e);
}

export function compileSource(source: string): { compiled?: CompiledDefinition; error?: string } {
	try {
		return { compiled: parseDefinition(source, parseYaml) };
	} catch (e) {
		return { error: errorText(e) };
	}
}

/** Finds every `excel-export` block in the definitions folder (or the whole vault if empty). */
export async function findDefinitions(app: App, folder: string): Promise<DefinitionEntry[]> {
	const prefix = folder.trim() ? normalizePath(folder.trim()) + "/" : "";
	const files = app.vault.getMarkdownFiles().filter((f) => !prefix || f.path.startsWith(prefix));
	const entries: DefinitionEntry[] = [];
	for (const file of files) {
		const content = await app.vault.cachedRead(file);
		if (!content.includes("excel-export")) continue;
		extractCodeBlocks(content).forEach((source, index) => {
			const { compiled, error } = compileSource(source);
			entries.push({ file, index, name: compiled?.def.name ?? file.basename, compiled, error });
		});
	}
	return entries.sort((a, b) => a.name.localeCompare(b.name));
}

export class DefinitionPicker extends FuzzySuggestModal<DefinitionEntry> {
	constructor(
		app: App,
		private entries: DefinitionEntry[],
		private onPick: (entry: DefinitionEntry) => void,
		placeholder = "Elige una definición de export…",
	) {
		super(app);
		this.setPlaceholder(placeholder);
	}

	getItems(): DefinitionEntry[] {
		return this.entries;
	}

	getItemText(item: DefinitionEntry): string {
		return `${item.name} ${item.file.path}`;
	}

	renderSuggestion(match: FuzzyMatch<DefinitionEntry>, el: HTMLElement): void {
		const item = match.item;
		el.createDiv({ text: item.name });
		const sub = el.createEl("small", { cls: "xte-muted" });
		sub.setText(item.error ? `⚠ YAML inválido — ${item.file.path}` : `${item.compiled?.def.mode} · ${item.file.path}`);
	}

	onChooseItem(item: DefinitionEntry): void {
		this.onPick(item);
	}
}
