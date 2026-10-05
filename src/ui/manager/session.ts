import type { App } from "obsidian";
import { ObsidianVaultAdapter } from "../../adapter/obsidian-adapter";
import { loadWorkbook, type Workbook } from "../../excel/workbook-io";
import { prepareRun, type PreparedRun } from "../../export/runner";
import { Runtime } from "../../graph/context";
import type { NoteRecord } from "../../model/note-record";
import { Report } from "../../model/report";
import { evaluate } from "../../query/evaluator";
import { parseFilter } from "../../query/parser";
import { compileTemplate, type TemplateConfig } from "../../store/templates";
import type ExcelTemplateExportPlugin from "../../main";

/** Everything the editor needs about one template: vault notes, sample run, workbook, autosave. */
export class TemplateSession {
	readonly adapter: ObsidianVaultAdapter;
	readonly rt: Runtime;
	private runCache: PreparedRun | null | undefined;
	private wbCache: { path: string; wb: Workbook | null } | null = null;
	private saveTimer: number | null = null;
	onSaved?: (ok: boolean) => void;

	constructor(readonly app: App, readonly plugin: ExcelTemplateExportPlugin, readonly cfg: TemplateConfig) {
		this.adapter = new ObsidianVaultAdapter(app);
		this.rt = new Runtime(this.adapter, new Report());
	}

	notes(): NoteRecord[] {
		return this.adapter.listNotes();
	}

	private filterNotes(notes: NoteRecord[], src: string | undefined): NoteRecord[] {
		if (!src || !src.trim()) return notes;
		try {
			const expr = parseFilter(src);
			return notes.filter((n) => evaluate(expr, n, this.rt));
		} catch {
			return [];
		}
	}

	/** Notes of the chosen type (root.where), ignoring the default filter. */
	candidates(): NoteRecord[] {
		if (!this.cfg.root.where.trim()) return [];
		return this.filterNotes(this.notes(), this.cfg.root.where);
	}

	/** Candidates that pass the default filter. */
	filtered(): NoteRecord[] {
		return this.filterNotes(this.candidates(), this.cfg.root.filter);
	}

	/** Notes matching an arbitrary filter (relations step). */
	matching(src: string | undefined): NoteRecord[] {
		return this.filterNotes(this.notes(), src);
	}

	/** Sample run (roots + relations) ignoring the cell mapping; null if the definition is incomplete. */
	run(): PreparedRun | null {
		if (this.runCache !== undefined) return this.runCache;
		// Ignore relations that are still being configured, so the rest of the editor keeps working.
		const rels = Object.entries(this.cfg.relations).filter(([, r]) => (r.source ? true : /^\s*\S.*->/.test(r.on ?? "")));
		const valid = Object.fromEntries(rels.filter(([, r]) => !r.source || rels.some(([k]) => k === r.source)));
		const { cd } = compileTemplate({ ...this.cfg, relations: valid, cells: {}, rows: {} });
		if (!cd) return (this.runCache = null);
		const filtered = this.filtered();
		const roots = (filtered.length ? filtered : this.candidates()).map((n) => n.path);
		try {
			this.runCache = prepareRun(cd, this.adapter, { settings: this.plugin.settings, rootPaths: roots });
		} catch {
			this.runCache = null;
		}
		return this.runCache;
	}

	/** Call after changing root/relations. */
	invalidate(): void {
		this.runCache = undefined;
	}

	async workbook(): Promise<Workbook | null> {
		const path = this.cfg.template;
		if (this.wbCache && this.wbCache.path === path) return this.wbCache.wb;
		let wb: Workbook | null = null;
		try {
			if (path && (await this.adapter.exists(path))) wb = await loadWorkbook(await this.adapter.readBinary(path));
		} catch {
			wb = null;
		}
		this.wbCache = { path, wb };
		return wb;
	}

	/** Marks the template as changed; saves after a short pause. */
	changed(opts: { structure?: boolean } = {}): void {
		if (opts.structure) this.invalidate();
		this.cfg.updatedAt = Date.now();
		if (this.saveTimer !== null) window.clearTimeout(this.saveTimer);
		this.saveTimer = window.setTimeout(() => void this.flush(), 500);
		this.onSaved?.(false);
	}

	async flush(): Promise<void> {
		if (this.saveTimer === null) return;
		window.clearTimeout(this.saveTimer);
		this.saveTimer = null;
		await this.plugin.saveTemplate(this.cfg);
		this.onSaved?.(true);
	}
}
