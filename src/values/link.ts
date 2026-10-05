/** A wikilink found in a frontmatter value, remembering which note it came from. */
export class LinkRef {
	constructor(
		readonly raw: string,
		readonly linkpath: string,
		readonly subpath: string | null,
		readonly alias: string | null,
		readonly sourcePath: string,
	) {}

	/** Basename of the link destination: `[[Roles/Backend|Back]]` → `Backend`. */
	get targetName(): string {
		const last = this.linkpath.split("/").pop() ?? this.linkpath;
		return last.replace(/\.md$/i, "");
	}

	/** What Obsidian would show: the alias if present, else the destination basename. */
	get display(): string {
		return this.alias ?? this.targetName;
	}
}

export interface ParsedLink {
	linkpath: string;
	subpath: string | null;
	alias: string | null;
}

const FULL_LINK_RE = /^\s*!?\[\[([^\]|#^]*)([#^][^\]|]*)?(?:\|([^\]]*))?\]\]\s*$/;
export const INLINE_LINK_RE = /!?\[\[([^\]|#^]*)([#^][^\]|]*)?(?:\|([^\]]*))?\]\]/g;

/** Parses a string that is exactly one wikilink. */
export function parseWikilink(text: string): ParsedLink | null {
	const m = FULL_LINK_RE.exec(text);
	if (!m) return null;
	const linkpath = m[1].trim();
	if (!linkpath && !m[2]) return null;
	return {
		linkpath,
		subpath: m[2] ? m[2].trim() : null,
		alias: m[3] !== undefined && m[3].trim() !== "" ? m[3].trim() : null,
	};
}

export function toLinkRef(text: string, sourcePath: string): LinkRef | null {
	const p = parseWikilink(text);
	return p ? new LinkRef(text.trim(), p.linkpath, p.subpath, p.alias, sourcePath) : null;
}

/**
 * Converts raw frontmatter values into values the exporter understands:
 * wikilink strings become {@link LinkRef}s, recursively inside lists.
 * Also accepts the YAML quirk where an unquoted `[[X]]` parses as `[["X"]]`.
 */
export function wrapFrontmatterValue(value: unknown, sourcePath: string): unknown {
	if (typeof value === "string") return toLinkRef(value, sourcePath) ?? value;
	if (Array.isArray(value)) {
		if (value.length === 1 && Array.isArray(value[0]) && value[0].length === 1 && typeof value[0][0] === "string") {
			return toLinkRef(`[[${value[0][0]}]]`, sourcePath) ?? value;
		}
		return value.map((v) => {
			if (Array.isArray(v) && v.length === 1 && typeof v[0] === "string") {
				return toLinkRef(`[[${v[0]}]]`, sourcePath) ?? v;
			}
			return wrapFrontmatterValue(v, sourcePath);
		});
	}
	return value;
}
