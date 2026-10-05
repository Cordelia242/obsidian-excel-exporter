import type { Runtime } from "../graph/context";
import { parsePath, resolveOnValue } from "../graph/path";
import type { NoteRecord } from "../model/note-record";
import { isEmptyValue } from "../values/normalize";
import { order } from "./evaluator";

export interface SortKey {
	path: string[];
	desc: boolean;
}

/** Parses `created desc, ["tech rating"] asc, priority`. */
export function parseSort(src: string): SortKey[] {
	return src
		.split(",")
		.map((s) => s.trim())
		.filter(Boolean)
		.map((part) => {
			const m = /^(.*?)\s+(asc|desc)$/i.exec(part);
			const path = m ? m[1] : part;
			return { path: parsePath(path), desc: m ? m[2].toLowerCase() === "desc" : false };
		});
}

function firstValue(v: unknown): unknown {
	return Array.isArray(v) ? v.find((x) => !isEmptyValue(x)) : v;
}

/** Stable sort; empty values always go last. */
export function sortNotes(notes: NoteRecord[], keys: SortKey[], rt: Runtime): NoteRecord[] {
	if (!keys.length) return notes;
	const decorated = notes.map((note, idx) => ({
		note,
		idx,
		vals: keys.map((k) => firstValue(resolveOnValue(note, k.path, rt))),
	}));
	decorated.sort((x, y) => {
		for (let i = 0; i < keys.length; i++) {
			const a = x.vals[i];
			const b = y.vals[i];
			const ea = isEmptyValue(a);
			const eb = isEmptyValue(b);
			if (ea && eb) continue;
			if (ea) return 1;
			if (eb) return -1;
			let c = order(a, b);
			if (c === null) c = typeof a === "number" ? -1 : typeof b === "number" ? 1 : 0;
			if (c !== 0) return keys[i].desc ? -c : c;
		}
		return x.idx - y.idx;
	});
	return decorated.map((d) => d.note);
}
