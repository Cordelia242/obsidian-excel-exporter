import { basenameOf, makeNote, type NoteRecord } from "../model/note-record";
import type { VaultAdapter } from "./vault-adapter";

export interface MockNote {
	path: string;
	frontmatter?: Record<string, unknown>;
	ctime?: number;
	mtime?: number;
}

/** In-memory vault, used by tests and handy for scripting. Resolves links like Obsidian (path, then basename). */
export class MemoryVaultAdapter implements VaultAdapter {
	readonly files = new Map<string, ArrayBuffer>();
	private notes: NoteRecord[];

	constructor(notes: MockNote[], private name = "Vault") {
		this.notes = notes.map((n) => makeNote(n.path, n.frontmatter ?? {}, n.ctime ?? 0, n.mtime ?? 0));
	}

	listNotes(): NoteRecord[] {
		return this.notes;
	}

	resolveLink(linkpath: string, sourcePath: string): NoteRecord | null {
		const lp = linkpath.replace(/\.md$/i, "").toLowerCase();
		const exact = this.notes.find((n) => n.path.replace(/\.md$/i, "").toLowerCase() === lp);
		if (exact) return exact;
		const folder = sourcePath.includes("/") ? sourcePath.slice(0, sourcePath.lastIndexOf("/") + 1) : "";
		const relative = this.notes.find((n) => n.path.replace(/\.md$/i, "").toLowerCase() === (folder + lp).toLowerCase());
		if (relative) return relative;
		if (lp.includes("/")) {
			return this.notes.find((n) => n.path.replace(/\.md$/i, "").toLowerCase().endsWith("/" + lp)) ?? null;
		}
		return this.notes.find((n) => basenameOf(n.path).toLowerCase() === lp) ?? null;
	}

	async readBinary(path: string): Promise<ArrayBuffer> {
		const f = this.files.get(path);
		if (!f) throw new Error(`Archivo no encontrado: ${path}`);
		return f;
	}

	async writeBinary(path: string, data: ArrayBuffer, overwrite: boolean): Promise<void> {
		if (!overwrite && this.files.has(path)) throw new Error(`El archivo ya existe: ${path}`);
		this.files.set(path, data);
	}

	async exists(path: string): Promise<boolean> {
		return this.files.has(path) || this.notes.some((n) => n.path === path);
	}

	vaultName(): string {
		return this.name;
	}
}
