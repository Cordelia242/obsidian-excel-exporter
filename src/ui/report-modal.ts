import { Modal, Notice, Platform, type App } from "obsidian";
import type { OverwriteAnswer } from "../export/runner";
import type { TemplateValidation } from "../export/validate";
import type { ExportWarning, WarningKind } from "../model/report";

const KIND_LABEL: Record<WarningKind, string> = {
	unresolved: "Placeholders sin resolver",
	"broken-link": "Links rotos",
	"skipped-note": "Notas omitidas",
	formula: "Fórmulas",
	merge: "Celdas combinadas",
	table: "Tablas de Excel",
	pipe: "Pipes",
	other: "Otros",
};

export class ReportModal extends Modal {
	constructor(app: App, private title: string, private warnings: ExportWarning[]) {
		super(app);
	}

	onOpen(): void {
		const { contentEl } = this;
		this.titleEl.setText(this.title);
		const groups = new Map<WarningKind, ExportWarning[]>();
		for (const w of this.warnings) groups.set(w.kind, [...(groups.get(w.kind) ?? []), w]);
		for (const [kind, list] of groups) {
			contentEl.createEl("h4", { text: `${KIND_LABEL[kind]} (${list.length})` });
			const ul = contentEl.createEl("ul", { cls: "xte-report-list" });
			for (const w of list) {
				const li = ul.createEl("li");
				li.createSpan({ text: w.message });
				const where = [w.location, w.note].filter(Boolean).join(" · ");
				if (where) li.createEl("small", { text: ` — ${where}`, cls: "xte-muted" });
			}
		}
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

const STATUS_LABEL: Record<string, string> = {
	ok: "✓",
	unknown: "✗ desconocido",
	unresolved: "⚠ no resuelve",
	"no-sample": "· sin muestra",
	error: "✗ error",
};

export class ValidationModal extends Modal {
	constructor(app: App, private name: string, private result: TemplateValidation) {
		super(app);
	}

	onOpen(): void {
		const { contentEl } = this;
		const r = this.result;
		this.titleEl.setText(`Validar template — ${this.name}`);
		contentEl.createEl("p", {
			text: `${r.roots} raíces. ${r.sampleRoot ? `Muestra: ${r.sampleRoot}` : ""}`,
		});
		if (r.issues.length) {
			contentEl.createEl("h4", { text: "Problemas" });
			const ul = contentEl.createEl("ul");
			for (const i of r.issues) ul.createEl("li", { text: i });
		}
		contentEl.createEl("h4", { text: `Filas #each (${r.eachRows.length})` });
		const eachTable = contentEl.createEl("table", { cls: "xte-table" });
		for (const e of r.eachRows) {
			const tr = eachTable.createEl("tr");
			tr.createEl("td", { text: `${e.sheet}!${e.address}` });
			tr.createEl("td", { text: e.collection });
			tr.createEl("td", { text: `${STATUS_LABEL[e.status]}${e.count !== undefined ? ` (${e.count})` : ""}` });
		}
		contentEl.createEl("h4", { text: `Placeholders (${r.placeholders.length})` });
		const table = contentEl.createEl("table", { cls: "xte-table" });
		for (const p of r.placeholders) {
			const tr = table.createEl("tr", { cls: `xte-${p.status}` });
			tr.createEl("td", { text: `${p.sheet}!${p.address}` });
			tr.createEl("td").createEl("code", { text: p.placeholder });
			tr.createEl("td", { text: STATUS_LABEL[p.status] + (p.detail ? ` — ${p.detail}` : "") });
		}
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

/** Asks what to do with an existing output file. Closing the modal means "skip". */
export function askOverwrite(app: App, path: string): Promise<OverwriteAnswer> {
	return new Promise((resolve) => {
		let answered = false;
		const modal = new Modal(app);
		const answer = (a: OverwriteAnswer) => {
			answered = true;
			resolve(a);
			modal.close();
		};
		modal.titleEl.setText("El archivo ya existe");
		modal.contentEl.createEl("p", { text: path });
		const buttons = modal.contentEl.createDiv({ cls: "modal-button-container" });
		buttons.createEl("button", { text: "Sobrescribir", cls: "mod-warning" }).onclick = () => answer("overwrite");
		buttons.createEl("button", { text: "Agregar sufijo", cls: "mod-cta" }).onclick = () => answer("suffix");
		buttons.createEl("button", { text: "Omitir" }).onclick = () => answer("skip");
		modal.onClose = () => {
			if (!answered) resolve("skip");
		};
		modal.open();
	});
}

type AppWithShell = App & { showInFolder?: (path: string) => void; openWithDefaultApp?: (path: string) => void };

export function openFile(app: App, path: string): void {
	(app as AppWithShell).openWithDefaultApp?.(path);
}

/** Short notice after an export, with "Ver detalle" and "Abrir carpeta" when relevant. */
export function showResultNotice(app: App, name: string, files: string[], warnings: ExportWarning[]): void {
	const frag = createFragment();
	const n = files.length;
	const w = warnings.length;
	frag.createSpan({
		text: `${name}: ${n} ${n === 1 ? "archivo generado" : "archivos generados"}${w ? `, ${w} ${w === 1 ? "warning" : "warnings"}` : ""}`,
	});
	const actions = frag.createDiv({ cls: "xte-notice-actions" });
	let notice: Notice | null = null;
	if (w) {
		const link = actions.createEl("a", { text: "Ver detalle", href: "#" });
		link.onclick = (ev) => {
			ev.preventDefault();
			notice?.hide();
			new ReportModal(app, `Reporte — ${name}`, warnings).open();
		};
	}
	const shell = app as AppWithShell;
	if (n && Platform.isDesktopApp && shell.showInFolder) {
		const open = actions.createEl("a", { text: "Abrir carpeta", href: "#" });
		open.onclick = (ev) => {
			ev.preventDefault();
			shell.showInFolder?.(files[0]);
		};
	}
	notice = new Notice(frag, w ? 12000 : 6000);
}
