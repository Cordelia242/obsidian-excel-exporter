import { PluginSettingTab, Setting, type App } from "obsidian";
import type ExcelTemplateExportPlugin from "./main";

export { DEFAULT_SETTINGS, type ExportSettings } from "./config/defaults";

export class ExcelExportSettingTab extends PluginSettingTab {
	constructor(app: App, private plugin: ExcelTemplateExportPlugin) {
		super(app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		const s = this.plugin.settings;
		const save = () => this.plugin.saveSettings();
		containerEl.empty();

		new Setting(containerEl)
			.setName("Carpeta de definiciones")
			.setDesc("Notas con bloques `excel-export`. Vacío = todo el vault.")
			.addText((t) => t.setPlaceholder("Exports/").setValue(s.definitionsFolder).onChange(async (v) => {
				s.definitionsFolder = v.trim();
				await save();
			}));

		new Setting(containerEl)
			.setName("Carpeta de salida por defecto")
			.setDesc("Se usa cuando la definición no tiene `output.folder`.")
			.addText((t) => t.setPlaceholder("Exports/out/").setValue(s.outputFolder).onChange(async (v) => {
				s.outputFolder = v.trim();
				await save();
			}));

		containerEl.createEl("h3", { text: "Normalización" });

		new Setting(containerEl)
			.setName("Links")
			.setDesc("Cómo se escriben los wikilinks: alias visible, nombre de la nota o texto crudo.")
			.addDropdown((d) => d
				.addOptions({ display: "Alias / nombre (display)", target: "Nombre de la nota (target)", raw: "Crudo [[...]] (raw)" })
				.setValue(s.links)
				.onChange(async (v) => {
					s.links = v as typeof s.links;
					await save();
				}));

		new Setting(containerEl)
			.setName("Separador de listas")
			.addText((t) => t.setValue(s.listSeparator).onChange(async (v) => {
				s.listSeparator = v;
				await save();
			}));

		new Setting(containerEl)
			.setName("Estrellas")
			.setDesc("⭐⭐⭐ como número (3) o como texto.")
			.addDropdown((d) => d
				.addOptions({ number: "Número", raw: "Texto" })
				.setValue(s.stars)
				.onChange(async (v) => {
					s.stars = v as typeof s.stars;
					await save();
				}));

		new Setting(containerEl)
			.setName("Valor vacío")
			.setDesc("Texto para valores vacíos o sin resolver. Vacío = celda vacía.")
			.addText((t) => t.setValue(s.emptyValue).onChange(async (v) => {
				s.emptyValue = v;
				await save();
			}));

		containerEl.createEl("h3", { text: "Exportación" });

		new Setting(containerEl)
			.setName("Fila #each vacía")
			.setDesc("Qué hacer con una fila #each cuando la colección no tiene elementos.")
			.addDropdown((d) => d
				.addOptions({ remove: "Eliminar la fila", blank: "Dejarla en blanco" })
				.setValue(s.emptyBlock)
				.onChange(async (v) => {
					s.emptyBlock = v as typeof s.emptyBlock;
					await save();
				}));

		new Setting(containerEl)
			.setName("Sobrescritura por defecto")
			.setDesc("Si el archivo de salida ya existe y la definición no indica `output.overwrite`.")
			.addDropdown((d) => d
				.addOptions({ ask: "Preguntar", overwrite: "Sobrescribir", suffix: "Agregar sufijo" })
				.setValue(s.overwrite)
				.onChange(async (v) => {
					s.overwrite = v as typeof s.overwrite;
					await save();
				}));

		new Setting(containerEl)
			.setName("Previsualizar antes de exportar")
			.setDesc("Muestra el Excel generado dentro de Obsidian y pide confirmación antes de guardarlo.")
			.addToggle((t) => t.setValue(s.previewBeforeExport).onChange(async (v) => {
				s.previewBeforeExport = v;
				await save();
			}));

		new Setting(containerEl)
			.setName("Abrir archivo al terminar")
			.setDesc("Abre el .xlsx generado con la aplicación predeterminada (desktop).")
			.addToggle((t) => t.setValue(s.openAfterExport).onChange(async (v) => {
				s.openAfterExport = v;
				await save();
			}));
	}
}
