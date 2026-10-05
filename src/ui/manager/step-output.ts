import { renderString } from "../../excel/placeholders";
import { rootScope, sanitizeFileName } from "../../export/runner";
import { callout, choiceCards, combobox, section } from "../components";
import type { TemplateSession } from "./session";

interface Token {
	label: string;
	text: string;
}

/** Text input with clickable tokens ("Nombre del Proceso", "Fecha de hoy") and a live example. */
function tokenInput(parent: HTMLElement, value: string, tokens: Token[], example: (v: string) => string, onChange: (v: string) => void): void {
	const input = parent.createEl("input", { type: "text", cls: "xte-input-lg", value });
	const chips = parent.createDiv({ cls: "xte-token-chips" });
	chips.createSpan({ cls: "xte-help", text: "Insertar:" });
	const ex = parent.createDiv({ cls: "xte-example" });
	const update = () => {
		ex.setText(`Ejemplo: ${example(input.value)}`);
		onChange(input.value);
	};
	for (const t of tokens) {
		const b = chips.createEl("button", { cls: "xte-token", text: t.label });
		b.onclick = () => {
			const pos = input.selectionStart ?? input.value.length;
			input.value = input.value.slice(0, pos) + t.text + input.value.slice(input.selectionEnd ?? pos);
			update();
			input.focus();
		};
	}
	input.oninput = update;
	ex.setText(`Ejemplo: ${example(value)}`);
}

/** Step 4: how files are produced, names, folder, overwrite. */
export async function renderStepOutput(el: HTMLElement, s: TemplateSession, rerender: () => void): Promise<void> {
	const cfg = s.cfg;
	const label = cfg.root.label || "Nota";
	const alias = cfg.root.alias;
	const wb = await s.workbook();
	const run = s.run();

	const modeBody = section(el, "¿Cómo se generan los archivos?", undefined, "1");
	choiceCards(
		modeBody,
		[
			{ value: "file-per-root", icon: "files", title: `Un archivo por cada ${label}`, description: `Ej.: exportas 3 ${label}s → 3 archivos Excel.` },
			{ value: "sheet-per-root", icon: "copy", title: `Un archivo con una hoja por ${label}`, description: "Una hoja del Excel se copia por cada uno; las demás hojas quedan como resumen." },
			{ value: "single", icon: "table", title: `Un solo archivo con todos los ${label}s`, description: "Para listados o consolidados: usa filas que se repiten." },
		],
		cfg.mode,
		(v) => {
			cfg.mode = v as typeof cfg.mode;
			if (v === "file-per-root" && !cfg.output.filename?.includes(`{{${alias}`)) cfg.output.filename = `{{${alias}.file.name}}.xlsx`;
			if (v !== "file-per-root" && cfg.output.filename?.includes(`{{${alias}.`)) cfg.output.filename = `${cfg.name}.xlsx`;
			s.changed({ structure: true });
			rerender();
		},
	);

	const sample = run?.roots[0];
	const scope = run && sample ? rootScope(run, sample, 0) : null;
	const render = (tpl: string) => {
		if (!run || !scope) return tpl;
		try {
			return renderString(tpl, scope, { rt: run.rt, opts: run.normalize, location: "output", warnUnresolved: false });
		} catch {
			return tpl;
		}
	};

	if (cfg.mode === "sheet-per-root" && wb) {
		const sheetBody = section(el, `¿Qué hoja se copia por cada ${label}?`, undefined, "1b");
		const sel = sheetBody.createEl("select", { cls: "dropdown" });
		for (const w of wb.worksheets) sel.createEl("option", { value: w.name, text: w.name });
		sel.value = cfg.output.repeatSheet ?? wb.worksheets[0]?.name ?? "";
		sel.onchange = () => {
			cfg.output.repeatSheet = sel.value;
			s.changed({ structure: true });
		};
		sheetBody.createDiv({ cls: "xte-help", text: "Nombre de cada hoja:" });
		tokenInput(
			sheetBody,
			cfg.output.sheetName ?? `{{${alias}.file.name}}`,
			[{ label: `Nombre del ${label}`, text: `{{${alias}.file.name}}` }],
			(v) => render(v).slice(0, 31),
			(v) => {
				cfg.output.sheetName = v;
				s.changed();
			},
		);
	}

	const nameBody = section(el, "Nombre del archivo", undefined, "2");
	const tokens: Token[] = [
		...(cfg.mode === "file-per-root" ? [{ label: `Nombre del ${label}`, text: `{{${alias}.file.name}}` }] : []),
		{ label: "Fecha de hoy", text: '{{@today | date:"yyyy-MM-dd"}}' },
		{ label: "Nombre del template", text: "{{@exportName}}" },
	];
	tokenInput(nameBody, cfg.output.filename ?? "", tokens, (v) => sanitizeFileName(render(v || "{{@exportName}}")), (v) => {
		cfg.output.filename = v || undefined;
		s.changed();
	});

	const folderBody = section(el, "Carpeta donde se guardan", undefined, "3");
	const folders = s.app.vault.getAllFolders(false).map((f) => f.path).sort();
	combobox(folderBody, {
		value: cfg.output.folder ?? s.plugin.settings.outputFolder,
		allowCustom: true,
		options: () => folders.map((p) => ({ value: p, label: p })),
		onChange: (v) => {
			cfg.output.folder = v;
			s.changed();
		},
	});
	folderBody.createDiv({ cls: "xte-help", text: "Si la carpeta no existe, se crea al exportar." });

	const owBody = section(el, "Si el archivo ya existe", undefined, "4");
	choiceCards(
		owBody,
		[
			{ value: "ask", title: "Preguntarme" },
			{ value: "suffix", title: "Crear una copia nueva (… (2).xlsx)" },
			{ value: "overwrite", title: "Reemplazarlo" },
		],
		cfg.output.overwrite ?? "ask",
		(v) => {
			cfg.output.overwrite = v as "ask" | "suffix" | "overwrite";
			s.changed();
		},
	);

	if (!run) callout(el, `Para ver ejemplos de nombres, elige en el paso 1 qué ${label}s exporta el template.`, "info");
}
