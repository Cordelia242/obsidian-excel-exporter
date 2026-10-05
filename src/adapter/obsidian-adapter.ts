import { normalizePath, TFile, type App } from "obsidian";
import type { NoteRecord } from "../model/note-record";
import type { VaultAdapter } from "./vault-adapter";

/**
 * VaultAdapter backed by the Obsidian API. Create one per export run:
 * the note list is read once and cached for the run.
 * Frontmatter values keep their raw `[[...]]` text in Obsidian's metadata cache,
 * so links are parsed by the core and resolved with `getFirstLinkpathDest`.
 */
export class ObsidianVaultAdapter implements VaultAdapter {
	private notes: NoteRecord[] | null = null;
	private byPath = new Map<string, NoteRecord>();

	constructor(private app: App) {}

	private toRecord(file: TFile): NoteRecord {
		const fm = this.app.metadataCache.getFileCache(file)?.frontmatter ?? {};
		const { position: _position, ...frontmatter } = fm as Record<string, unknown>;
		return { path: file.path, basename: file.basename, frontmatter, ctime: file.stat.ctime, mtime: file.stat.mtime };
	}

	listNotes(): NoteRecord[] {
		if (!this.notes) {
			this.notes = this.app.vault.getMarkdownFiles().map((f) => this.toRecord(f));
			for (const n of this.notes) this.byPath.set(n.path, n);
		}
		return this.notes;
	}

	resolveLink(linkpath: string, sourcePath: string): NoteRecord | null {
		const file = this.app.metadataCache.getFirstLinkpathDest(linkpath, sourcePath);
		if (!file || file.extension !== "md") return null;
		this.listNotes();
		return this.byPath.get(file.path) ?? this.toRecord(file);
	}

	async readBinary(path: string): Promise<ArrayBuffer> {
		const file = this.app.vault.getFileByPath(normalizePath(path));
		if (!file) throw new Error(`No se encontró el archivo «${path}»`);
		return this.app.vault.readBinary(file);
	}

	async writeBinary(path: string, data: ArrayBuffer, overwrite: boolean): Promise<void> {
		const p = normalizePath(path);
		const folder = p.includes("/") ? p.slice(0, p.lastIndexOf("/")) : "";
		if (folder && !this.app.vault.getAbstractFileByPath(folder)) {
			await this.app.vault.createFolder(folder).catch(() => undefined);
		}
		const existing = this.app.vault.getAbstractFileByPath(p);
		if (existing instanceof TFile) {
			if (!overwrite) throw new Error(`El archivo ya existe: ${p}`);
			await this.app.vault.modifyBinary(existing, data);
		} else {
			await this.app.vault.createBinary(p, data);
		}
	}

	async exists(path: string): Promise<boolean> {
		return this.app.vault.getAbstractFileByPath(normalizePath(path)) !== null;
	}

	vaultName(): string {
		return this.app.vault.getName();
	}
}
