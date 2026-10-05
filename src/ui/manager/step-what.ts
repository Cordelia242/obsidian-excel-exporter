import { setIcon } from "obsidian";
import { conditionsToExpr } from "../../setup/conditions";
import { noteTypeSuggestions, propertiesOf, slugify } from "../../setup/discovery";
import { renameVariable } from "../../store/templates";
import { callout, combobox, section } from "../components";
import { conditionBuilder } from "../condition-builder";
import type { TemplateSession } from "./session";

function commonFirstWord(names: string[]): string | null {
	if (names.length < 2) return null;
	const first = names[0].split(/\s+/)[0];
	return first.length > 2 && names.every((n) => n.split(/\s+/)[0] === first) ? first : null;
}

/** Step 1: which notes the template exports (root type), what to call them, default filter and order. */
export function renderStepWhat(el: HTMLElement, s: TemplateSession, rerender: () => void): void {
	const cfg = s.cfg;
	const label = cfg.root.label || "Nota";

	// 1. Note type
	const typeBody = section(
		el,
		"¿De qué tipo de nota sale cada exportación?",
		"Elige una de las sugerencias o arma la condición. Por ejemplo: «Procesos de selección» o «Personas».",
		"1",
	);
	const suggestions = noteTypeSuggestions(s.notes());
	if (suggestions.length) {
		const chips = typeBody.createDiv({ cls: "xte-suggest-chips" });
		chips.createSpan({ cls: "xte-help", text: "Sugerencias de tu vault:" });
		for (const sug of suggestions) {
			const expr = conditionsToExpr([sug.condition]);
			const chip = chips.createEl("button", { cls: `xte-suggest${cfg.root.where === expr ? " is-selected" : ""}` });
			chip.createSpan({ text: sug.label });
			chip.createSpan({ cls: "xte-suggest-count", text: String(sug.count) });
			chip.title = `Ejemplos: ${sug.examples.join(", ")}`;
			chip.onclick = () => {
				cfg.root.where = expr;
				// Name the notes after what they have in common ("Proceso Backend", "Proceso Data" → "Proceso").
				const word = commonFirstWord(s.candidates().map((n) => n.basename)) ?? sug.label.replace(/^Carpeta /, "");
				if (!cfg.root.label || cfg.root.label === "Nota" || cfg.root.label === label) setRootLabel(s, word);
				s.changed({ structure: true });
				rerender();
			};
		}
	}
	const condBox = typeBody.createDiv();
	conditionBuilder(condBox, {
		expr: cfg.root.where,
		notes: () => s.notes(),
		addLabel: "Agregar condición",
		emptyText: "Todavía no elegiste el tipo de nota.",
		onChange: (expr) => {
			cfg.root.where = expr;
			s.changed({ structure: true });
			updateCount();
		},
	});
	const countBox = typeBody.createDiv();
	const updateCount = () => {
		countBox.empty();
		if (!cfg.root.where.trim()) return;
		const c = s.candidates();
		if (!c.length) callout(countBox, "Ninguna nota cumple estas condiciones.", "warn");
		else {
			const names = c.slice(0, 5).map((n) => n.basename).join(", ");
			callout(countBox, `${c.length} ${c.length === 1 ? "nota coincide" : "notas coinciden"}: ${names}${c.length > 5 ? "…" : ""}`, "ok");
		}
	};
	updateCount();

	// 2. Name
	const nameBody = section(el, "¿Cómo llamas a cada una de estas notas?", "Se usa en todo el editor y al exportar: «¿Qué Procesos quieres exportar?».", "2");
	const nameInput = nameBody.createEl("input", { type: "text", cls: "xte-input-md", value: cfg.root.label ?? "" });
	nameInput.placeholder = "Ej.: Proceso, Persona, Candidato";
	nameInput.onchange = () => {
		setRootLabel(s, nameInput.value.trim() || "Nota");
		s.changed({ structure: true });
		rerender();
	};

	// 3. Default filter
	const filterBody = section(
		el,
		`¿Cuáles ${label}s se exportan por defecto?`,
		`Opcional. Al exportar siempre puedes elegir ${label.toLowerCase()}s concretos; este filtro define la opción «Todos».`,
		"3",
	);
	conditionBuilder(filterBody, {
		expr: cfg.root.filter ?? "",
		notes: () => s.candidates(),
		addLabel: "Agregar filtro",
		emptyText: `Sin filtro: «Todos» exporta todos los ${label}s.`,
		onChange: (expr) => {
			cfg.root.filter = expr || undefined;
			s.changed({ structure: true });
			updateFilterCount();
		},
	});
	const filterCount = filterBody.createDiv();
	const updateFilterCount = () => {
		filterCount.empty();
		if (!cfg.root.filter || !cfg.root.where.trim()) return;
		const f = s.filtered().length;
		const t = s.candidates().length;
		callout(filterCount, `${f} de ${t} pasan el filtro.`, f ? "ok" : "warn");
	};
	updateFilterCount();

	// 4. Order
	const orderBody = section(el, "Orden", `En qué orden se exportan los ${label}s (y las hojas o filas, según el modo).`, "4");
	const row = orderBody.createDiv({ cls: "xte-inline" });
	const [field = "", dir = "asc"] = parseSortSimple(cfg.root.sort);
	combobox(row, {
		value: field,
		placeholder: "Sin orden",
		options: () => [
			{ value: "", label: "Sin orden" },
			{ value: "file.name", label: "Nombre de la nota" },
			...propertiesOf(s.candidates()).map((p) => ({ value: p.key, label: p.key, detail: p.sample })),
		],
		onChange: (v) => {
			cfg.root.sort = v ? `${sortField(v)} ${dirSel.value}` : undefined;
			s.changed({ structure: true });
		},
	});
	const dirSel = row.createEl("select", { cls: "dropdown" });
	dirSel.createEl("option", { value: "asc", text: "Ascendente (A→Z, antiguo→nuevo)" });
	dirSel.createEl("option", { value: "desc", text: "Descendente (Z→A, nuevo→antiguo)" });
	dirSel.value = dir;
	dirSel.onchange = () => {
		const [f] = parseSortSimple(cfg.root.sort);
		if (f) cfg.root.sort = `${sortField(f)} ${dirSel.value}`;
		s.changed({ structure: true });
	};
	const tip = orderBody.createDiv({ cls: "xte-help" });
	setIcon(tip.createSpan(), "lightbulb");
	tip.appendText(" Siguiente paso: traer los datos relacionados (por ejemplo, las Interviews de cada Proceso).");
}

function sortField(f: string): string {
	return /^[\w.@-]+$/.test(f) ? f : `["${f}"]`;
}

function parseSortSimple(sort: string | undefined): [string, string] | [] {
	if (!sort) return [];
	const m = /^\s*(?:\["(.+)"\]|(.+?))\s+(asc|desc)\s*$/i.exec(sort.split(",")[0]);
	if (!m) return [sort.trim(), "asc"];
	return [m[1] ?? m[2], m[3].toLowerCase()];
}

/** Changes the friendly name and, safely, the internal alias used in expressions. */
export function setRootLabel(s: TemplateSession, label: string): void {
	const cfg = s.cfg;
	cfg.root.label = label;
	const alias = slugify(label);
	if (alias && alias !== cfg.root.alias && !cfg.relations[alias]) renameVariable(cfg, cfg.root.alias, alias);
}
