import type { VaultAdapter } from "../adapter/vault-adapter";
import { FileInfo, type NoteRecord } from "../model/note-record";
import type { Report } from "../model/report";
import type { LinkRef } from "../values/link";
import { startOfToday } from "../values/normalize";

/** Shared state for one export run: link resolution cache, relations per root, warnings. */
export class Runtime {
	private linkCache = new Map<string, NoteRecord | null>();
	private relations = new Map<string, Map<string, NoteRecord[]>>();
	private fileInfos = new Map<string, FileInfo>();

	constructor(
		readonly adapter: VaultAdapter,
		readonly report: Report,
		readonly now: Date = new Date(),
		readonly exportName = "",
	) {}

	get today(): Date {
		return startOfToday(this.now);
	}

	resolveLink(link: LinkRef, warnIfBroken = true): NoteRecord | null {
		const key = `${link.sourcePath}\u0000${link.linkpath}`;
		let note = this.linkCache.get(key);
		if (note === undefined) {
			note = link.linkpath ? this.adapter.resolveLink(link.linkpath, link.sourcePath) : null;
			this.linkCache.set(key, note);
		}
		if (!note && warnIfBroken) {
			this.report.warn("broken-link", `Link roto ${link.raw}`, { note: link.sourcePath });
		}
		return note;
	}

	setRelations(root: NoteRecord, relations: Map<string, NoteRecord[]>): void {
		this.relations.set(root.path, relations);
	}

	relationsOf(note: NoteRecord): Map<string, NoteRecord[]> | undefined {
		return this.relations.get(note.path);
	}

	fileUri(note: NoteRecord): string {
		const vault = this.adapter.vaultName?.() ?? "";
		const params = vault ? `vault=${encodeURIComponent(vault)}&` : "";
		return `obsidian://open?${params}file=${encodeURIComponent(note.path)}`;
	}

	fileInfo(note: NoteRecord): FileInfo {
		let info = this.fileInfos.get(note.path);
		if (!info) {
			const name = note.path.split("/").pop() ?? note.path;
			const dot = name.lastIndexOf(".");
			const folder = note.path.includes("/") ? note.path.slice(0, note.path.lastIndexOf("/")) : "";
			info = new FileInfo(
				note.basename,
				note.basename,
				note.path,
				folder,
				dot >= 0 ? name.slice(dot + 1) : "",
				new Date(note.ctime),
				new Date(note.mtime),
				this.fileUri(note),
			);
			this.fileInfos.set(note.path, info);
		}
		return info;
	}
}

/** Variables visible to a placeholder or filter: aliases, relation names, `@` specials. */
export interface Scope {
	vars: Record<string, unknown>;
	specials: Record<string, unknown>;
}

export function makeScope(vars: Record<string, unknown>, specials: Record<string, unknown> = {}): Scope {
	return { vars, specials };
}

export function childScope(parent: Scope, vars: Record<string, unknown>, specials: Record<string, unknown> = {}): Scope {
	return { vars: { ...parent.vars, ...vars }, specials: { ...parent.specials, ...specials } };
}

/** Exact key lookup, then case-insensitive. Returns `[found, value]`. */
export function lookupKey(obj: Record<string, unknown>, key: string): [boolean, unknown] {
	if (Object.prototype.hasOwnProperty.call(obj, key)) return [true, obj[key]];
	const lower = key.toLowerCase();
	for (const k of Object.keys(obj)) {
		if (k.toLowerCase() === lower) return [true, obj[k]];
	}
	return [false, undefined];
}
