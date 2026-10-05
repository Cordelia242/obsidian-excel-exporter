import type { EmptyBlockMode, OverwriteMode } from "./schema";
import type { NormalizeOptions } from "../values/normalize";
import type { TemplateConfig } from "../store/templates";

/** Plugin settings (§10). Pure: no Obsidian imports. */
export interface ExportSettings {
	definitionsFolder: string;
	outputFolder: string;
	links: NormalizeOptions["links"];
	listSeparator: string;
	stars: NormalizeOptions["stars"];
	emptyValue: string;
	emptyBlock: EmptyBlockMode;
	openAfterExport: boolean;
	overwrite: OverwriteMode;
	previewBeforeExport: boolean;
	/** Where Excel files uploaded from the computer are stored. */
	templatesFolder: string;
	templates: TemplateConfig[];
}

export const DEFAULT_SETTINGS: ExportSettings = {
	definitionsFolder: "Exports/",
	outputFolder: "Exports/out/",
	links: "display",
	listSeparator: ", ",
	stars: "number",
	emptyValue: "",
	emptyBlock: "remove",
	openAfterExport: false,
	overwrite: "ask",
	previewBeforeExport: true,
	templatesFolder: "Templates/Excel",
	templates: [],
};
