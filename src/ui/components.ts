import { setIcon } from "obsidian";

export interface ComboOption {
	value: string;
	label: string;
	/** Secondary text (example value, count, folder…). */
	detail?: string;
	icon?: string;
}

export interface ComboboxOptions {
	options: ComboOption[] | (() => ComboOption[]);
	value?: string;
	placeholder?: string;
	/** Accept a typed value that is not in the list. */
	allowCustom?: boolean;
	/** Clear the input after picking (multi-select pickers). */
	clearOnPick?: boolean;
	onChange: (value: string, option?: ComboOption) => void;
	cls?: string;
}

export interface Combobox {
	el: HTMLElement;
	input: HTMLInputElement;
	setValue: (value: string) => void;
	focus: () => void;
}

/** Searchable dropdown: type to filter, click or Enter to pick. */
export function combobox(parent: HTMLElement, o: ComboboxOptions): Combobox {
	const wrap = parent.createDiv({ cls: `xte-combo ${o.cls ?? ""}` });
	const input = wrap.createEl("input", { type: "text", placeholder: o.placeholder ?? "Buscar…" });
	const caret = wrap.createSpan({ cls: "xte-combo-caret" });
	setIcon(caret, "chevron-down");
	const list = wrap.createDiv({ cls: "xte-combo-list" });
	list.hide();
	let current = o.value ?? "";
	let active = 0;
	let shown: ComboOption[] = [];
	const all = () => (typeof o.options === "function" ? o.options() : o.options);
	const labelOf = (v: string) => all().find((x) => x.value === v)?.label ?? v;
	input.value = current ? labelOf(current) : "";

	const render = () => {
		const q = input.value.trim().toLowerCase();
		const showAll = !q || (current !== "" && input.value === labelOf(current));
		shown = all()
			.filter((x) => showAll || x.label.toLowerCase().includes(q) || (x.detail ?? "").toLowerCase().includes(q))
			.slice(0, 80);
		list.empty();
		if (o.allowCustom && q && !shown.some((x) => x.label.toLowerCase() === q)) {
			shown.unshift({ value: input.value.trim(), label: input.value.trim(), detail: "usar este valor" });
		}
		if (!shown.length) list.createDiv({ cls: "xte-combo-empty", text: "Sin resultados" });
		shown.forEach((opt, i) => {
			const row = list.createDiv({ cls: `xte-combo-item${i === active ? " is-active" : ""}${opt.value === current ? " is-current" : ""}` });
			if (opt.icon) setIcon(row.createSpan({ cls: "xte-combo-icon" }), opt.icon);
			row.createSpan({ cls: "xte-combo-label", text: opt.label });
			if (opt.detail) row.createSpan({ cls: "xte-combo-detail", text: opt.detail });
			row.onmousedown = (ev) => {
				ev.preventDefault();
				pick(opt);
			};
		});
	};
	const open = () => {
		active = 0;
		render();
		list.show();
	};
	const close = () => list.hide();
	const pick = (opt: ComboOption) => {
		current = o.clearOnPick ? "" : opt.value;
		input.value = o.clearOnPick ? "" : opt.label;
		close();
		o.onChange(opt.value, opt);
	};
	input.onfocus = () => {
		input.select();
		open();
	};
	input.oninput = () => {
		active = 0;
		render();
		list.show();
	};
	input.onblur = () => {
		close();
		if (!o.clearOnPick) input.value = current ? labelOf(current) : "";
	};
	input.onkeydown = (ev) => {
		if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
			ev.preventDefault();
			if (list.isShown()) active = Math.max(0, Math.min(shown.length - 1, active + (ev.key === "ArrowDown" ? 1 : -1)));
			open();
		} else if (ev.key === "Enter") {
			ev.preventDefault();
			if (shown[active]) pick(shown[active]);
		} else if (ev.key === "Escape") {
			close();
			input.blur();
		}
	};
	caret.onmousedown = (ev) => {
		ev.preventDefault();
		input.focus();
	};
	return {
		el: wrap,
		input,
		setValue: (v) => {
			current = v;
			input.value = v ? labelOf(v) : "";
		},
		focus: () => input.focus(),
	};
}

export interface MultiPickerOptions {
	options: ComboOption[];
	selected: string[];
	placeholder?: string;
	onChange: (selected: string[]) => void;
}

/** Multi-select picker: chips for the chosen items + a searchable combobox. */
export function multiPicker(parent: HTMLElement, o: MultiPickerOptions): void {
	const wrap = parent.createDiv({ cls: "xte-multi" });
	const chips = wrap.createDiv({ cls: "xte-chips" });
	let selected = [...o.selected];
	const byValue = new Map(o.options.map((x) => [x.value, x]));
	const draw = () => {
		chips.empty();
		for (const v of selected) {
			const chip = chips.createDiv({ cls: "xte-chip" });
			chip.createSpan({ text: byValue.get(v)?.label ?? v });
			const x = chip.createSpan({ cls: "xte-chip-x", attr: { "aria-label": "Quitar" } });
			setIcon(x, "x");
			x.onclick = () => {
				selected = selected.filter((s) => s !== v);
				draw();
				o.onChange(selected);
			};
		}
	};
	combobox(wrap, {
		options: () => o.options.filter((x) => !selected.includes(x.value)),
		placeholder: o.placeholder,
		clearOnPick: true,
		onChange: (v) => {
			if (!selected.includes(v)) selected.push(v);
			draw();
			o.onChange(selected);
		},
	});
	draw();
}

export interface ChoiceCard {
	value: string;
	title: string;
	description?: string;
	icon?: string;
}

/** Radio group shown as cards (for choices that need an explanation). */
export function choiceCards(parent: HTMLElement, cards: ChoiceCard[], value: string, onChange: (v: string) => void): void {
	const wrap = parent.createDiv({ cls: "xte-choices" });
	for (const c of cards) {
		const el = wrap.createDiv({ cls: `xte-choice${c.value === value ? " is-selected" : ""}`, attr: { role: "radio", tabindex: "0" } });
		if (c.icon) setIcon(el.createDiv({ cls: "xte-choice-icon" }), c.icon);
		const body = el.createDiv();
		body.createDiv({ cls: "xte-choice-title", text: c.title });
		if (c.description) body.createDiv({ cls: "xte-choice-desc", text: c.description });
		const choose = () => {
			wrap.querySelectorAll(".xte-choice").forEach((x) => x.removeClass("is-selected"));
			el.addClass("is-selected");
			onChange(c.value);
		};
		el.onclick = choose;
		el.onkeydown = (ev) => {
			if (ev.key === "Enter" || ev.key === " ") {
				ev.preventDefault();
				choose();
			}
		};
	}
}

export function iconButton(parent: HTMLElement, icon: string, label: string, onClick: () => void, cls = ""): HTMLButtonElement {
	const b = parent.createEl("button", { cls: `xte-icon-btn ${cls}`, attr: { "aria-label": label, title: label } });
	setIcon(b, icon);
	b.onclick = (ev) => {
		ev.stopPropagation();
		onClick();
	};
	return b;
}

export function button(parent: HTMLElement, text: string, onClick: () => void, opts: { cta?: boolean; icon?: string; cls?: string } = {}): HTMLButtonElement {
	const b = parent.createEl("button", { cls: `${opts.cta ? "mod-cta " : ""}${opts.cls ?? ""}` });
	if (opts.icon) setIcon(b.createSpan({ cls: "xte-btn-icon" }), opts.icon);
	b.createSpan({ text });
	b.onclick = () => onClick();
	return b;
}

/** Section with a numbered/titled header and help text. */
export function section(parent: HTMLElement, title: string, help?: string, badge?: string): HTMLElement {
	const s = parent.createDiv({ cls: "xte-section" });
	const h = s.createDiv({ cls: "xte-section-head" });
	if (badge) h.createSpan({ cls: "xte-step-badge", text: badge });
	h.createSpan({ cls: "xte-section-title", text: title });
	if (help) s.createDiv({ cls: "xte-help", text: help });
	return s.createDiv({ cls: "xte-section-body" });
}

export function callout(parent: HTMLElement, text: string, kind: "info" | "ok" | "warn" = "info", icon?: string): HTMLElement {
	const el = parent.createDiv({ cls: `xte-callout xte-callout-${kind}` });
	setIcon(el.createSpan({ cls: "xte-callout-icon" }), icon ?? (kind === "ok" ? "check-circle" : kind === "warn" ? "alert-triangle" : "info"));
	el.createSpan({ text });
	return el;
}

/** Compact single-choice pills (wrap). */
export function pillChoices(parent: HTMLElement, options: Array<{ value: string; label: string; icon?: string; title?: string }>, value: string, onChange: (v: string) => void): HTMLElement {
	const wrap = parent.createDiv({ cls: "xte-pills" });
	for (const o of options) {
		const b = wrap.createEl("button", { cls: `xte-pill-btn${o.value === value ? " is-selected" : ""}` });
		if (o.icon) setIcon(b.createSpan({ cls: "xte-btn-icon" }), o.icon);
		b.createSpan({ text: o.label });
		if (o.title) b.title = o.title;
		b.onclick = () => onChange(o.value);
	}
	return wrap;
}

export interface ChecklistOptions {
	options: ComboOption[];
	selected: string[];
	placeholder?: string;
	onChange: (selected: string[]) => void;
}

/** Searchable list with checkboxes, always visible (no hidden dropdown). */
export function checklist(parent: HTMLElement, o: ChecklistOptions): void {
	const wrap = parent.createDiv({ cls: "xte-checklist" });
	const top = wrap.createDiv({ cls: "xte-checklist-top" });
	const search = top.createEl("input", { type: "search", placeholder: o.placeholder ?? "Buscar…" });
	const count = top.createSpan({ cls: "xte-help" });
	const list = wrap.createDiv({ cls: "xte-checklist-list" });
	const actions = wrap.createDiv({ cls: "xte-checklist-actions" });
	const selected = new Set(o.selected);
	let visible: ComboOption[] = [];
	const emit = () => {
		count.setText(`${selected.size} ${selected.size === 1 ? "elegido" : "elegidos"}`);
		o.onChange(o.options.filter((x) => selected.has(x.value)).map((x) => x.value));
	};
	const draw = () => {
		const q = search.value.trim().toLowerCase();
		visible = o.options.filter((x) => !q || x.label.toLowerCase().includes(q) || (x.detail ?? "").toLowerCase().includes(q));
		list.empty();
		if (!visible.length) list.createDiv({ cls: "xte-combo-empty", text: "Sin resultados" });
		for (const x of visible) {
			const row = list.createEl("label", { cls: `xte-check-row${selected.has(x.value) ? " is-checked" : ""}` });
			const cb = row.createEl("input", { type: "checkbox" });
			cb.checked = selected.has(x.value);
			const txt = row.createDiv();
			txt.createDiv({ cls: "xte-check-label", text: x.label });
			if (x.detail) txt.createDiv({ cls: "xte-check-detail", text: x.detail });
			cb.onchange = () => {
				if (cb.checked) selected.add(x.value);
				else selected.delete(x.value);
				row.toggleClass("is-checked", cb.checked);
				emit();
			};
		}
	};
	button(actions, "Marcar los visibles", () => {
		for (const x of visible) selected.add(x.value);
		draw();
		emit();
	}, { cls: "xte-link-btn" });
	button(actions, "Desmarcar todos", () => {
		selected.clear();
		draw();
		emit();
	}, { cls: "xte-link-btn xte-muted-btn" });
	search.oninput = draw;
	draw();
	count.setText(`${selected.size} ${selected.size === 1 ? "elegido" : "elegidos"}`);
	window.setTimeout(() => search.focus(), 50);
}
