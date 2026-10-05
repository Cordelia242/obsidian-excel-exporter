import type { NoteRecord } from "../model/note-record";
import { conditionsToExpr, exprToConditions, OP_LABELS, VALUELESS_OPS, type CondOp, type Condition } from "../setup/conditions";
import { propertiesOf, valuesOfProperty, type PropKind } from "../setup/discovery";
import { button, combobox, iconButton } from "./components";

export interface ConditionBuilderOptions {
	/** Current filter text. */
	expr: string;
	/** Notes whose properties/values are offered. */
	notes: () => NoteRecord[];
	onChange: (expr: string) => void;
	addLabel?: string;
	emptyText?: string;
}

const OPS_BY_KIND: Record<string, CondOp[]> = {
	list: ["contains", "!contains", "exists", "!exists"],
	link: ["==", "!=", "exists", "!exists"],
	number: ["==", "!=", ">", ">=", "<", "<=", "exists", "!exists"],
	date: [">=", "<=", "==", "!=", "exists", "!exists"],
	stars: ["==", ">=", "<=", "exists", "!exists"],
	default: ["==", "!=", "contains", "!contains", "exists", "!exists"],
};

const FILE_FIELDS = [
	{ value: "file.name", label: "Nombre de la nota", detail: "archivo" },
	{ value: "file.folder", label: "Carpeta", detail: "archivo" },
];

/**
 * Visual "AND" filter builder: [field ▾] [operator ▾] [value ▾].
 * Fields and values come from the notes. Expressions it cannot represent
 * are shown as text with an option to edit them directly.
 */
export function conditionBuilder(parent: HTMLElement, o: ConditionBuilderOptions): void {
	const wrap = parent.createDiv({ cls: "xte-conds" });
	const parsed = exprToConditions(o.expr);
	if (parsed === null) {
		renderAdvanced(wrap, o);
		return;
	}
	let conds: Condition[] = parsed;
	const kinds = () => new Map(propertiesOf(o.notes()).map((p) => [p.key, p.kind]));
	const emit = () => o.onChange(conditionsToExpr(conds));

	const draw = () => {
		wrap.empty();
		if (!conds.length && o.emptyText) wrap.createDiv({ cls: "xte-help", text: o.emptyText });
		const kindMap = kinds();
		conds.forEach((c, i) => {
			const row = wrap.createDiv({ cls: "xte-cond" });
			if (i > 0) row.createSpan({ cls: "xte-cond-and", text: "y" });
			const kind: PropKind | "default" = kindMap.get(c.field) ?? "default";
			combobox(row, {
				cls: "xte-cond-field",
				value: c.field,
				placeholder: "Campo…",
				options: () => [
					...propertiesOf(o.notes()).map((p) => ({ value: p.key, label: p.key, detail: p.sample })),
					...FILE_FIELDS,
				],
				onChange: (v) => {
					c.field = v;
					const ops = OPS_BY_KIND[kindMap.get(v) ?? "default"] ?? OPS_BY_KIND.default;
					if (!ops.includes(c.op)) c.op = ops[0];
					c.value = "";
					emit();
					draw();
				},
			});
			const ops = OPS_BY_KIND[kind] ?? OPS_BY_KIND.default;
			const opSel = row.createEl("select", { cls: "dropdown xte-cond-op" });
			for (const op of ops.includes(c.op) ? ops : [c.op, ...ops]) opSel.createEl("option", { value: op, text: OP_LABELS[op] });
			opSel.value = c.op;
			opSel.onchange = () => {
				c.op = opSel.value as CondOp;
				emit();
				draw();
			};
			if (!VALUELESS_OPS.includes(c.op)) {
				combobox(row, {
					cls: "xte-cond-value",
					value: c.value,
					placeholder: "Valor…",
					allowCustom: true,
					options: () =>
						c.field
							? valuesOfProperty(o.notes(), c.field).map((v) => ({ value: v.value, label: v.value, detail: `${v.count}` }))
							: [],
					onChange: (v) => {
						c.value = v;
						emit();
					},
				});
			}
			iconButton(row, "x", "Quitar condición", () => {
				conds.splice(i, 1);
				emit();
				draw();
			}, "xte-cond-remove");
		});
		const actions = wrap.createDiv({ cls: "xte-cond-actions" });
		button(actions, o.addLabel ?? "Agregar condición", () => {
			conds.push({ field: "", op: "==", value: "" });
			draw();
		}, { icon: "plus", cls: "xte-link-btn" });
		button(actions, "Escribir como texto", () => {
			wrap.empty();
			renderAdvanced(wrap, { ...o, expr: conditionsToExpr(conds) });
		}, { icon: "code", cls: "xte-link-btn xte-muted-btn" });
	};
	draw();
}

function renderAdvanced(wrap: HTMLElement, o: ConditionBuilderOptions): void {
	wrap.createDiv({
		cls: "xte-help",
		text: 'Filtro escrito a mano. Ejemplo: status != "Closed" and priority >= 3',
	});
	const input = wrap.createEl("textarea", { cls: "xte-expr-text" });
	input.value = o.expr;
	input.rows = 2;
	input.oninput = () => o.onChange(input.value);
	button(wrap, "Volver al editor visual", () => {
		const conds = exprToConditions(input.value);
		if (conds === null) {
			wrap.querySelector(".xte-adv-error")?.remove();
			wrap.createDiv({ cls: "xte-adv-error xte-error-text", text: "Este filtro usa 'or', 'not' o fechas: solo se puede editar como texto." });
			return;
		}
		wrap.empty();
		conditionBuilder(wrap.parentElement ?? wrap, { ...o, expr: input.value });
		wrap.remove();
	}, { icon: "list", cls: "xte-link-btn xte-muted-btn" });
}
