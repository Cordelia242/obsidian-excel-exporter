/** A markdown note as seen by the exporter: path + parsed frontmatter. */
export interface NoteRecord {
	path: string;
	/** File name without extension (what Obsidian calls `basename`). */
	basename: string;
	frontmatter: Record<string, unknown>;
	ctime: number;
	mtime: number;
}

export function isNoteRecord(value: unknown): value is NoteRecord {
	if (typeof value !== "object" || value === null) return false;
	const v = value as Partial<NoteRecord>;
	return (
		typeof v.path === "string" &&
		typeof v.basename === "string" &&
		typeof v.frontmatter === "object" &&
		typeof v.mtime === "number"
	);
}

/** Pseudo-properties exposed as `<note>.file.<prop>`. */
export class FileInfo {
	constructor(
		readonly name: string,
		readonly basename: string,
		readonly path: string,
		readonly folder: string,
		readonly ext: string,
		readonly ctime: Date,
		readonly mtime: Date,
		readonly link: string,
	) {}
}

export function basenameOf(path: string): string {
	const name = path.split("/").pop() ?? path;
	return name.replace(/\.md$/i, "");
}

export function makeNote(
	path: string,
	frontmatter: Record<string, unknown> = {},
	ctime = 0,
	mtime = 0,
): NoteRecord {
	return { path, basename: basenameOf(path), frontmatter, ctime, mtime };
}
