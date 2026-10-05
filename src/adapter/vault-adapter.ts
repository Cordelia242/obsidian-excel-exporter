import type { NoteRecord } from "../model/note-record";

/** Everything the core needs from the vault. Implemented for Obsidian and in-memory (tests). */
export interface VaultAdapter {
	/** All markdown notes (with their frontmatter, possibly empty). */
	listNotes(): NoteRecord[];
	/** Resolves a link path (without `[[ ]]`, alias or subpath) from a source note. */
	resolveLink(linkpath: string, sourcePath: string): NoteRecord | null;
	readBinary(path: string): Promise<ArrayBuffer>;
	/** Writes a file, creating parent folders as needed. */
	writeBinary(path: string, data: ArrayBuffer, overwrite: boolean): Promise<void>;
	exists(path: string): Promise<boolean>;
	/** Vault name, used to build `obsidian://open` URIs. */
	vaultName?(): string;
}
