import type { CompiledRelation } from "../config/schema";
import { isNoteRecord, type NoteRecord } from "../model/note-record";
import { evaluate } from "../query/evaluator";
import { sortNotes } from "../query/sort";
import { LinkRef } from "../values/link";
import { makeScope, type Runtime } from "./context";
import { resolveOnValue } from "./path";

/**
 * Builds relations for a set of roots. Backlink relations use an inverted
 * index (resolved target path → candidates) built once per run, so the join
 * is O(n + m) instead of O(n·m).
 */
export class RelationBuilder {
	private indexes = new Map<string, Map<string, NoteRecord[]>>();

	constructor(
		private relations: CompiledRelation[],
		private notes: NoteRecord[],
		private rootAlias: string,
		private rt: Runtime,
	) {}

	private index(rel: CompiledRelation): Map<string, NoteRecord[]> {
		let idx = this.indexes.get(rel.name);
		if (idx) return idx;
		idx = new Map();
		for (const cand of this.notes) {
			if (rel.from && !evaluate(rel.from, cand, this.rt)) continue;
			const value = resolveOnValue(cand, rel.field ?? [], this.rt);
			const targets = new Set<string>();
			for (const v of Array.isArray(value) ? value.flat(Infinity) : [value]) {
				if (v instanceof LinkRef) {
					const note = this.rt.resolveLink(v);
					if (note) targets.add(note.path);
				} else if (isNoteRecord(v)) {
					targets.add(v.path);
				}
			}
			for (const t of targets) {
				const list = idx.get(t);
				if (list) list.push(cand);
				else idx.set(t, [cand]);
			}
		}
		this.indexes.set(rel.name, idx);
		return idx;
	}

	/** Computes every relation for `root` and registers them on the runtime. */
	build(root: NoteRecord): Map<string, NoteRecord[]> {
		const result = new Map<string, NoteRecord[]>();
		const scope = makeScope({ [this.rootAlias]: root });
		for (const rel of this.relations) {
			let items =
				rel.kind === "derived" ? [...(result.get(rel.source ?? "") ?? [])] : [...(this.index(rel).get(root.path) ?? [])];
			if (rel.filter) {
				const f = rel.filter;
				items = items.filter((n) => evaluate(f, n, this.rt, scope));
			}
			items = sortNotes(items, rel.sort, this.rt);
			result.set(rel.name, items);
		}
		this.rt.setRelations(root, result);
		return result;
	}
}
