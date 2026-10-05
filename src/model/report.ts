export type WarningKind =
	| "unresolved"
	| "broken-link"
	| "skipped-note"
	| "formula"
	| "merge"
	| "table"
	| "pipe"
	| "other";

export interface ExportWarning {
	kind: WarningKind;
	message: string;
	/** Note the warning refers to, when there is one. */
	note?: string;
	/** Sheet!Cell of the template, when there is one. */
	location?: string;
}

/** Collects warnings for one run, ignoring exact duplicates. */
export class Report {
	readonly warnings: ExportWarning[] = [];
	private seen = new Set<string>();

	warn(kind: WarningKind, message: string, extra: { note?: string; location?: string } = {}): void {
		const key = `${kind}|${message}|${extra.note ?? ""}|${extra.location ?? ""}`;
		if (this.seen.has(key)) return;
		this.seen.add(key);
		this.warnings.push({ kind, message, ...extra });
	}
}
