import { Modal, Notice, normalizePath, setIcon, type App, type TFile } from "obsidian";
import { button } from "./components";

/** Saves an .xlsx chosen on the computer into the vault templates folder. Returns its vault path. */
export async function importExcelFile(app: App, folder: string, file: File): Promise<string> {
	const dir = normalizePath(folder || "Templates/Excel");
	if (!app.vault.getAbstractFileByPath(dir)) await app.vault.createFolder(dir).catch(() => undefined);
	const base = file.name.replace(/\.xlsx$/i, "");
	let path = `${dir}/${base}.xlsx`;
	for (let i = 2; app.vault.getAbstractFileByPath(path); i++) path = `${dir}/${base} (${i}).xlsx`;
	await app.vault.createBinary(path, await file.arrayBuffer());
	return path;
}

export function excelFiles(app: App): TFile[] {
	return app.vault
		.getFiles()
		.filter((f) => f.extension.toLowerCase() === "xlsx" && !f.path.startsWith("Exports/out"))
		.sort((a, b) => a.basename.localeCompare(b.basename));
}

/** Chooses an Excel file of the vault, or uploads one from the computer. */
export class ExcelPickerModal extends Modal {
	constructor(
		app: App,
		private folder: string,
		private onPick: (path: string, name: string) => void,
		private title = "Elige el archivo Excel",
		private askName = false,
	) {
		super(app);
	}

	onOpen(): void {
		this.titleEl.setText(this.title);
		this.modalEl.addClass("xte-modal-mid");
		const { contentEl } = this;
		let chosen: string | null = null;
		let nameInput: HTMLInputElement | null = null;

		contentEl.createDiv({
			cls: "xte-help",
			text: "El Excel solo necesita los títulos y el formato (colores, anchos, logos). Las celdas de datos pueden quedar vacías: luego eliges qué va en cada una.",
		});

		const upload = contentEl.createDiv({ cls: "xte-upload" });
		setIcon(upload.createSpan({ cls: "xte-upload-icon" }), "upload");
		upload.createSpan({ text: "Subir un Excel desde tu computador" });
		const fileInput = upload.createEl("input", { type: "file", attr: { accept: ".xlsx" } });
		fileInput.hide();
		upload.onclick = () => fileInput.click();
		fileInput.onchange = async () => {
			const f = fileInput.files?.[0];
			if (!f) return;
			try {
				const path = await importExcelFile(this.app, this.folder, f);
				new Notice(`Excel guardado en ${path}`);
				chosen = path;
				if (!this.askName) this.finish(path, f.name.replace(/\.xlsx$/i, ""));
				else {
					if (nameInput && !nameInput.value) nameInput.value = f.name.replace(/\.xlsx$/i, "");
					this.drawList(list, () => chosen, (p) => (chosen = p), nameInput);
				}
			} catch (e) {
				new Notice(`No se pudo guardar el archivo: ${(e as Error).message}`);
			}
		};

		const files = excelFiles(this.app);
		contentEl.createDiv({ cls: "xte-q", text: files.length ? "O elige uno que ya está en tu vault:" : "No hay archivos .xlsx en tu vault todavía." });
		const list = contentEl.createDiv({ cls: "xte-file-cards" });

		if (this.askName) {
			const nb = contentEl.createDiv({ cls: "xte-inline xte-name-row" });
			nb.createSpan({ text: "Nombre del template:" });
			nameInput = nb.createEl("input", { type: "text", cls: "xte-input-md", placeholder: "Ej.: Ficha de proceso" });
			const footer = contentEl.createDiv({ cls: "modal-button-container" });
			button(footer, "Cancelar", () => this.close());
			const create = button(footer, "Crear template", () => {
				if (!chosen) return new Notice("Elige o sube un archivo Excel");
				this.finish(chosen, nameInput?.value.trim() || (chosen.split("/").pop() ?? "").replace(/\.xlsx$/i, ""));
			}, { cta: true });
			create.addClass("xte-create");
		}
		this.drawList(list, () => chosen, (p) => {
			chosen = p;
			if (!this.askName) this.finish(p, (p.split("/").pop() ?? p).replace(/\.xlsx$/i, ""));
			else if (nameInput && !nameInput.value) nameInput.value = (p.split("/").pop() ?? p).replace(/\.xlsx$/i, "");
		}, nameInput);
	}

	private drawList(list: HTMLElement, chosen: () => string | null, choose: (p: string) => void, nameInput: HTMLInputElement | null): void {
		list.empty();
		for (const f of excelFiles(this.app)) {
			const card = list.createDiv({ cls: `xte-file-card${chosen() === f.path ? " is-selected" : ""}` });
			setIcon(card.createSpan({ cls: "xte-file-icon" }), "sheet");
			const t = card.createDiv();
			t.createDiv({ cls: "xte-file-name", text: f.basename });
			t.createDiv({ cls: "xte-help", text: f.parent?.path && f.parent.path !== "/" ? f.parent.path : "raíz del vault" });
			card.onclick = () => {
				choose(f.path);
				this.drawList(list, chosen, choose, nameInput);
			};
		}
	}

	private finish(path: string, name: string): void {
		this.close();
		this.onPick(path, name);
	}

	onClose(): void {
		this.contentEl.empty();
	}
}
