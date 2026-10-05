import type { EmptyBlockMode, OverwriteMode } from "./schema";
import type { NormalizeOptions } from "../values/normalize";

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
};
