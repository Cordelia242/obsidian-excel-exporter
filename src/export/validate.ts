import type { VaultAdapter } from "../adapter/vault-adapter";
import type { CompiledDefinition } from "../config/schema";
import type { Scope } from "../graph/context";
import { resolveHead, resolvePath } from "../graph/path";
import { MISSING } from "../values/normalize";
import { KNOWN_PIPES } from "../values/pipes";
import { iterateCollection } from "../excel/placeholders";
import { scanWorksheet } from "../excel/scanner";
import { rowRepeatsFor } from "../excel/mapping";
import { collectionScope, loadPreparedTemplate, loadTemplate, prepareRun, rootScope, type RunOptions } from "./runner";

export type CheckStatus = "ok" | "unknown" | "unresolved" | "no-sample" | "error";

export interface PlaceholderCheck {
	sheet: string;
	address: string;
	placeholder: string;
	status: CheckStatus;
	detail?: string;
}

export interface EachCheck {
	sheet: string;
	address: string;
	collection: string;
	status: CheckStatus;
	/** Items in the sample root. */
	count?: number;
}

export interface TemplateValidation {
	roots: number;
	sampleRoot?: string;
	placeholders: PlaceholderCheck[];
	eachRows: EachCheck[];
	issues: string[];
}

/** Scans the template and checks every placeholder against the first root (§9, command 3). */
export async function validateTemplate(
	cd: CompiledDefinition,
	adapter: VaultAdapter,
	options: RunOptions = {},
): Promise<TemplateValidation> {
	const data = await loadTemplate(adapter, cd.def.template);
	const run = prepareRun(cd, adapter, options);
	const wb = await loadPreparedTemplate(data, cd, run);
	const firstSheet = wb.worksheets[0]?.name ?? "";
	const result: TemplateValidation = { roots: run.roots.length, placeholders: [], eachRows: [], issues: [] };
	if (cd.def.mode !== "single" && run.roots.length) result.sampleRoot = run.roots[0].path;
	if (!run.roots.length) result.issues.push("Ninguna nota cumple el filtro de la raíz: no hay datos de muestra");

	wb.worksheets.forEach((ws, sheetIdx) => {
		const perRoot = cd.def.mode === "file-per-root" || (cd.def.mode === "sheet-per-root" && sheetIdx === 0);
		const base: Scope | null = perRoot ? (run.roots.length ? rootScope(run, run.roots[0], 0) : null) : collectionScope(run);
		const scan = scanWorksheet(ws, rowRepeatsFor(cd.mapping, ws, firstSheet));
		for (const i of scan.issues) result.issues.push(`${ws.name}!${i.address}: ${i.message}`);

		const rowScope = new Map<number, Scope | null>();
		for (const e of scan.eachRows) {
			const each = e.tpl.each;
			if (!each) continue;
			const check: EachCheck = { sheet: ws.name, address: e.address, collection: each.src, status: "ok" };
			if (each.error) {
				check.status = "error";
				rowScope.set(e.row, null);
			} else if (!base) {
				check.status = "no-sample";
				rowScope.set(e.row, null);
			} else {
				const items = iterateCollection(each.path, base, run.rt);
				if (items === null) {
					check.status = "unknown";
					rowScope.set(e.row, null);
				} else {
					check.count = items.length;
					if (!items.length) check.status = "no-sample";
					rowScope.set(e.row, items[0] ?? null);
				}
			}
			result.eachRows.push(check);
		}

		for (const c of scan.cells) {
			const scope = rowScope.has(c.row) ? rowScope.get(c.row) ?? null : base;
			for (const part of c.tpl.parts) {
				if (part.kind !== "ph") continue;
				const check: PlaceholderCheck = { sheet: ws.name, address: c.address, placeholder: `{{${part.src}}}`, status: "ok" };
				const unknownPipe = part.pipes.find((p) => !KNOWN_PIPES.includes(p.name));
				if (part.error) {
					check.status = "error";
					check.detail = part.error;
				} else if (unknownPipe) {
					check.status = "error";
					check.detail = `Pipe desconocido "${unknownPipe.name}"`;
				} else if (!scope) {
					check.status = "no-sample";
				} else if (resolveHead(part.path[0], scope) === MISSING) {
					check.status = "unknown";
					check.detail = `"${part.path[0]}" no es un alias, relación ni variable especial`;
				} else if (resolvePath(part.path, scope, run.rt) === MISSING) {
					check.status = "unresolved";
					check.detail = "No resuelve en la nota de muestra";
				}
				result.placeholders.push(check);
			}
		}
	});
	return result;
}
