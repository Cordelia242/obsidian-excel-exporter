import { conditionsToExpr } from "../../setup/conditions";
import { propertiesOf, relationSuggestions, slugify } from "../../setup/discovery";
import { renameVariable, uniqueName } from "../../store/templates";
import { button, callout, combobox, iconButton, section } from "../components";
import { conditionBuilder } from "../condition-builder";
import type { TemplateSession } from "./session";

/** Step 2: notes related to each root (backlinks) and subsets of them. */
export function renderStepRelations(el: HTMLElement, s: TemplateSession, rerender: () => void): void {
	const cfg = s.cfg;
	const label = cfg.root.label || "Nota";
	if (!cfg.root.where.trim()) {
		callout(el, "Primero elige en el paso 1 qué notas exporta el template.", "warn");
		return;
	}
	const candidates = s.candidates();
	const backlinks = Object.entries(cfg.relations).filter(([, r]) => !r.source);
	const derived = Object.entries(cfg.relations).filter(([, r]) => r.source);

	el.createDiv({
		cls: "xte-help xte-lead",
		text: `Trae información de otras notas que apuntan a cada ${label}. Por ejemplo, las Interviews cuyo campo «process» enlaza al ${label}. Así puedes llenar una tabla con ellas o contarlas.`,
	});

	// Suggestions
	const existing = new Set(backlinks.map(([, r]) => `${(r.on ?? "").split("->")[0].trim()}|${r.from ?? ""}`));
	const sugs = relationSuggestions(s.notes(), candidates, s.rt).filter(
		(x) => !existing.has(`${x.field}|${x.from ? conditionsToExpr([x.from]) : ""}`),
	);
	if (sugs.length) {
		const body = section(el, "Sugerencias encontradas en tu vault");
		for (const sug of sugs.slice(0, 6)) {
			const card = body.createDiv({ cls: "xte-sugg-card" });
			const text = card.createDiv({ cls: "xte-sugg-text" });
			text.createDiv({ cls: "xte-sugg-title", text: sug.label });
			text.createDiv({
				cls: "xte-help",
				text: `${sug.count} notas apuntan a un ${label} mediante «${sug.field}». Ej.: ${sug.examples.join(", ")}`,
			});
			button(card, "Agregar", () => {
				const name = uniqueName(cfg, sug.label, slugify);
				cfg.relations[name] = {
					label: sug.label,
					from: sug.from ? conditionsToExpr([sug.from]) : undefined,
					on: `${sug.field} -> ${cfg.root.alias}`,
				};
				s.changed({ structure: true });
				rerender();
			}, { icon: "plus", cta: true });
		}
	}

	// Configured relations
	const listBody = section(el, backlinks.length ? "Datos relacionados del template" : "Datos relacionados", backlinks.length ? undefined : "Todavía no agregaste ninguno. Es opcional: si solo necesitas datos de la propia nota, sigue al paso 3.");
	const run = s.run();
	const sample = run?.roots[0];
	for (const [name, rel] of backlinks) {
		const card = listBody.createDiv({ cls: "xte-rel-card" });
		const head = card.createDiv({ cls: "xte-rel-head" });
		const nameInput = head.createEl("input", { type: "text", cls: "xte-rel-name", value: rel.label ?? name });
		nameInput.title = "Nombre que verás en el editor";
		nameInput.onchange = () => {
			rel.label = nameInput.value.trim() || name;
			const next = uniqueName({ ...cfg, relations: Object.fromEntries(Object.entries(cfg.relations).filter(([k]) => k !== name)) }, rel.label, slugify);
			renameVariable(cfg, name, next);
			s.changed({ structure: true });
			rerender();
		};
		iconButton(head, "trash-2", "Eliminar", () => {
			delete cfg.relations[name];
			for (const [k, r] of Object.entries(cfg.relations)) if (r.source === name) delete cfg.relations[k];
			s.changed({ structure: true });
			rerender();
		});

		const grid = card.createDiv({ cls: "xte-form-grid" });
		grid.createDiv({ cls: "xte-form-label", text: "Notas a buscar" });
		const fromBox = grid.createDiv();
		conditionBuilder(fromBox, {
			expr: rel.from ?? "",
			notes: () => s.notes(),
			emptyText: "Todas las notas del vault (agrega una condición para acotar).",
			onChange: (expr) => {
				rel.from = expr || undefined;
				s.changed({ structure: true });
			},
		});

		grid.createDiv({ cls: "xte-form-label", text: `Campo que apunta al ${label}` });
		const field = (rel.on ?? "").split("->")[0].trim();
		combobox(grid.createDiv(), {
			value: field,
			placeholder: "Elige el campo…",
			options: () =>
				propertiesOf(s.matching(rel.from))
					.filter((p) => p.kind === "link" || p.kind === "list")
					.map((p) => ({ value: p.key, label: p.key, detail: p.sample })),
			onChange: (v) => {
				rel.on = `${v} -> ${cfg.root.alias}`;
				s.changed({ structure: true });
				rerender();
			},
		});

		grid.createDiv({ cls: "xte-form-label", text: "Ordenar por" });
		const sortRow = grid.createDiv({ cls: "xte-inline" });
		const [sf, sd] = (rel.sort ?? "").split(/\s+(?=asc$|desc$)/i);
		combobox(sortRow, {
			value: sf ?? "",
			placeholder: "Sin orden",
			options: () => [{ value: "", label: "Sin orden" }, ...propertiesOf(s.matching(rel.from)).map((p) => ({ value: p.key, label: p.key, detail: p.sample }))],
			onChange: (v) => {
				rel.sort = v ? `${/^[\w.-]+$/.test(v) ? v : `["${v}"]`} ${dir.value}` : undefined;
				s.changed({ structure: true });
			},
		});
		const dir = sortRow.createEl("select", { cls: "dropdown" });
		dir.createEl("option", { value: "asc", text: "Ascendente" });
		dir.createEl("option", { value: "desc", text: "Descendente" });
		dir.value = (sd ?? "asc").toLowerCase();
		dir.onchange = () => {
			if (rel.sort) rel.sort = rel.sort.replace(/\s+(asc|desc)$/i, ` ${dir.value}`);
			s.changed({ structure: true });
		};

		if (sample && run) {
			const items = run.rt.relationsOf(sample)?.get(name) ?? [];
			callout(
				card,
				`Ejemplo: «${sample.basename}» tiene ${items.length}${items.length ? `: ${items.slice(0, 4).map((n) => n.basename).join(", ")}${items.length > 4 ? "…" : ""}` : ""}.`,
				items.length ? "ok" : "info",
			);
		}

		// Subsets of this relation
		const subs = derived.filter(([, r]) => r.source === name);
		const subBox = card.createDiv({ cls: "xte-subsets" });
		for (const [subName, sub] of subs) {
			const subCard = subBox.createDiv({ cls: "xte-subset" });
			const sh = subCard.createDiv({ cls: "xte-rel-head" });
			sh.createSpan({ cls: "xte-help", text: "Subconjunto:" });
			const sn = sh.createEl("input", { type: "text", cls: "xte-rel-name", value: sub.label ?? subName });
			sn.onchange = () => {
				sub.label = sn.value.trim() || subName;
				const next = uniqueName({ ...cfg, relations: Object.fromEntries(Object.entries(cfg.relations).filter(([k]) => k !== subName)) }, sub.label, slugify);
				renameVariable(cfg, subName, next);
				s.changed({ structure: true });
				rerender();
			};
			iconButton(sh, "trash-2", "Eliminar subconjunto", () => {
				delete cfg.relations[subName];
				s.changed({ structure: true });
				rerender();
			});
			subCard.createDiv({ cls: "xte-help", text: `Solo las ${rel.label ?? name} que cumplen:` });
			conditionBuilder(subCard, {
				expr: sub.filter ?? "",
				notes: () => s.matching(rel.from),
				onChange: (expr) => {
					sub.filter = expr || undefined;
					s.changed({ structure: true });
				},
			});
			if (sample && run) {
				const n = run.rt.relationsOf(sample)?.get(subName)?.length ?? 0;
				subCard.createDiv({ cls: "xte-help", text: `En «${sample.basename}»: ${n}.` });
			}
		}
		button(subBox, `Agregar subconjunto (ej.: solo las ${rel.label ?? name} contratadas)`, () => {
			const subName = uniqueName(cfg, `${rel.label ?? name} filtradas`, slugify);
			cfg.relations[subName] = { label: `${rel.label ?? name} filtradas`, source: name };
			s.changed({ structure: true });
			rerender();
		}, { icon: "filter", cls: "xte-link-btn" });
	}

	button(listBody, "Agregar datos relacionados manualmente", () => {
		const name = uniqueName(cfg, "Relacionadas", slugify);
		cfg.relations[name] = { label: "Relacionadas", on: `-> ${cfg.root.alias}` };
		s.changed({ structure: true });
		rerender();
	}, { icon: "plus", cls: "xte-link-btn" });
}
