import { Notice, normalizePath, Plugin, TFile, type TAbstractFile } from "obsidian";
import { ObsidianVaultAdapter } from "./adapter/obsidian-adapter";
import { CODEBLOCK_LANG, EXAMPLE_DEFINITION } from "./config/parse";
import type { CompiledDefinition } from "./config/schema";
import { buildExport, writeExport, type RunOptions } from "./export/runner";
import { validateTemplate } from "./export/validate";
import { Runtime } from "./graph/context";
import { Report } from "./model/report";
import { evaluate } from "./query/evaluator";
import { DEFAULT_SETTINGS, ExcelExportSettingTab, type ExportSettings } from "./settings";
import { renderCodeBlock } from "./ui/codeblock";
import { DefinitionPicker, errorText, findDefinitions, type DefinitionEntry } from "./ui/definition-picker";
import { PreviewModal } from "./ui/preview-modal";
import { askOverwrite, openFile, showResultNotice, ValidationModal } from "./ui/report-modal";
import { SETUP_VIEW_TYPE, SetupView } from "./ui/setup-view";

export default class ExcelTemplateExportPlugin extends Plugin {
	settings: ExportSettings = { ...DEFAULT_SETTINGS };

	async onload(): Promise<void> {
		await this.loadSettings();
		this.addSettingTab(new ExcelExportSettingTab(this.app, this));
		this.registerView(SETUP_VIEW_TYPE, (leaf) => new SetupView(leaf, this));

		this.addCommand({
			id: "setup",
			name: "Setup (configurar celdas del Excel)…",
			callback: () => this.pickDefinition((entry) => void this.openSetup(entry.file.path, entry.index)),
		});

		this.addCommand({
			id: "run-export",
			name: "Ejecutar export…",
			callback: () => this.pickDefinition((entry) => this.withCompiled(entry, (cd) => this.runDefinition(cd))),
		});

		this.addCommand({
			id: "export-active-note",
			name: "Exportar nota activa con…",
			checkCallback: (checking) => {
				const file = this.app.workspace.getActiveFile();
				if (!file || file.extension !== "md") return false;
				if (!checking) void this.exportNote(file);
				return true;
			},
		});

		this.addCommand({
			id: "validate-template",
			name: "Validar template…",
			callback: () => this.pickDefinition((entry) => this.withCompiled(entry, (cd) => this.validateDefinition(cd))),
		});

		this.addCommand({
			id: "new-definition",
			name: "Nueva definición",
			callback: () => void this.createDefinition(),
		});

		this.registerMarkdownCodeBlockProcessor(CODEBLOCK_LANG, (source, el, ctx) => renderCodeBlock(this, source, el, ctx));

		this.registerEvent(
			this.app.workspace.on("file-menu", (menu, file: TAbstractFile) => {
				if (!(file instanceof TFile) || file.extension !== "md") return;
				menu.addItem((item) =>
					item
						.setTitle("Exportar a Excel con…")
						.setIcon("sheet")
						.onClick(() => void this.exportNote(file)),
				);
			}),
		);
	}

	async loadSettings(): Promise<void> {
		this.settings = { ...DEFAULT_SETTINGS, ...((await this.loadData()) as Partial<ExportSettings> | null) };
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.settings);
	}

	private async pickDefinition(onPick: (entry: DefinitionEntry) => void, entries?: DefinitionEntry[], placeholder?: string) {
		const list = entries ?? (await findDefinitions(this.app, this.settings.definitionsFolder));
		if (!list.length) {
			new Notice(
				`Excel Export: no hay definiciones en "${this.settings.definitionsFolder || "el vault"}". Usa "Nueva definición".`,
			);
			return;
		}
		new DefinitionPicker(this.app, list, onPick, placeholder).open();
	}

	private withCompiled(entry: DefinitionEntry, fn: (cd: CompiledDefinition) => void): void {
		if (!entry.compiled) {
			new Notice(`Excel Export: "${entry.file.path}" tiene una definición inválida:\n${entry.error}`, 10000);
			return;
		}
		fn(entry.compiled);
	}

	/** Command 2 / file menu: picks among definitions whose root.where matches the note. */
	async exportNote(file: TFile): Promise<void> {
		const all = await findDefinitions(this.app, this.settings.definitionsFolder);
		const adapter = new ObsidianVaultAdapter(this.app);
		const note = adapter.listNotes().find((n) => n.path === file.path);
		const rt = new Runtime(adapter, new Report());
		const matching = note ? all.filter((e) => e.compiled && evaluate(e.compiled.where, note, rt)) : [];
		let entries = matching;
		let placeholder = `Exportar "${file.basename}" con…`;
		if (!matching.length) {
			if (all.length) new Notice("Excel Export: ninguna definición coincide con esta nota (root.where); se muestran todas.");
			entries = all;
			placeholder = `Exportar "${file.basename}" con… (ninguna coincide con root.where)`;
		}
		await this.pickDefinition(
			(entry) => this.withCompiled(entry, (cd) => void this.runDefinition(cd, file.path)),
			entries,
			placeholder,
		);
	}

	async runDefinition(cd: CompiledDefinition, activeNotePath?: string): Promise<void> {
		const adapter = new ObsidianVaultAdapter(this.app);
		const options: RunOptions = {
			settings: this.settings,
			activeNotePath,
			confirmOverwrite: (path) => askOverwrite(this.app, path),
		};
		const progress = new Notice(`Excel Export: generando "${cd.def.name}"…`, 0);
		try {
			const built = await buildExport(cd, adapter, options);
			progress.hide();
			const write = async () => {
				try {
					const result = await writeExport(built, adapter, options);
					showResultNotice(this.app, cd.def.name, result.files, result.warnings);
					if (this.settings.openAfterExport && result.files.length === 1) openFile(this.app, result.files[0]);
				} catch (e) {
					console.error("Excel Template Export", e);
					new Notice(`Excel Export: error al guardar "${cd.def.name}": ${errorText(e)}`, 10000);
				}
			};
			if (this.settings.previewBeforeExport) {
				new PreviewModal(this.app, `Vista previa — ${cd.def.name}`, built, () => void write()).open();
			} else {
				await write();
			}
		} catch (e) {
			progress.hide();
			console.error("Excel Template Export", e);
			new Notice(`Excel Export: error en "${cd.def.name}": ${errorText(e)}`, 10000);
		}
	}

	async openSetup(path: string, index: number): Promise<void> {
		const existing = this.app.workspace
			.getLeavesOfType(SETUP_VIEW_TYPE)
			.find((l) => {
				const st = l.getViewState().state as { file?: string; index?: number } | undefined;
				return st?.file === path && (st?.index ?? 0) === index;
			});
		const leaf = existing ?? this.app.workspace.getLeaf("tab");
		if (!existing) await leaf.setViewState({ type: SETUP_VIEW_TYPE, active: true, state: { file: path, index } });
		this.app.workspace.revealLeaf(leaf);
	}

	async validateDefinition(cd: CompiledDefinition): Promise<void> {
		try {
			const result = await validateTemplate(cd, new ObsidianVaultAdapter(this.app), { settings: this.settings });
			new ValidationModal(this.app, cd.def.name, result).open();
		} catch (e) {
			new Notice(`Excel Export: no se pudo validar "${cd.def.name}": ${errorText(e)}`, 10000);
		}
	}

	private async createDefinition(): Promise<void> {
		const folder = normalizePath(this.settings.definitionsFolder || "/");
		if (folder !== "/" && !this.app.vault.getAbstractFileByPath(folder)) await this.app.vault.createFolder(folder);
		const base = folder === "/" ? "" : `${folder}/`;
		let path = `${base}Nueva definición.md`;
		for (let i = 2; this.app.vault.getAbstractFileByPath(path); i++) path = `${base}Nueva definición ${i}.md`;
		const content = `Definición de export a Excel. Ajusta \`root\` y \`relations\`, elige el template y usa **Setup** para indicar qué dato va en cada celda.\n\n\`\`\`${CODEBLOCK_LANG}\n${EXAMPLE_DEFINITION}\`\`\`\n`;
		const file = await this.app.vault.create(path, content);
		await this.app.workspace.getLeaf(true).openFile(file);
	}
}
