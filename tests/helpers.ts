import ExcelJS from "exceljs";
import { parse as parseYaml } from "yaml";
import notes from "./fixtures/vault.json";
import { MemoryVaultAdapter, type MockNote } from "../src/adapter/memory-adapter";
import { parseDefinition } from "../src/config/parse";
import { Runtime } from "../src/graph/context";
import { Report } from "../src/model/report";
import { loadWorkbook, saveWorkbook } from "../src/excel/workbook-io";

export const NOW = new Date(2026, 9, 5, 10, 30);

export function vault(extra: MockNote[] = []): MemoryVaultAdapter {
	return new MemoryVaultAdapter([...(notes as MockNote[]), ...extra], "Contratación");
}

export function runtime(adapter = vault()): Runtime {
	return new Runtime(adapter, new Report(), NOW, "Test");
}

export function note(adapter: MemoryVaultAdapter, basename: string) {
	const n = adapter.listNotes().find((x) => x.basename === basename);
	if (!n) throw new Error(`fixture note not found: ${basename}`);
	return n;
}

export function definition(yaml: string) {
	return parseDefinition(yaml, (t) => parseYaml(t));
}

export async function toBuffer(wb: ExcelJS.Workbook): Promise<ArrayBuffer> {
	return saveWorkbook(wb);
}

export async function readXlsx(adapter: MemoryVaultAdapter, path: string): Promise<ExcelJS.Workbook> {
	const data = adapter.files.get(path);
	if (!data) throw new Error(`missing output ${path}; have: ${[...adapter.files.keys()].join(", ")}`);
	return loadWorkbook(data);
}

const HEADER_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
const ROW_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };

/** Template of §14.1 (Proceso.xlsx). */
export async function procesoTemplate(): Promise<ArrayBuffer> {
	const wb = new ExcelJS.Workbook();
	const ws = wb.addWorksheet("Proceso");
	ws.columns = [
		{ width: 28 }, { width: 16 }, { width: 22 }, { width: 34 },
		{ width: 12 }, { width: 10 }, { width: 10 }, { width: 14 },
	];
	ws.getCell("A1").value = "Proceso";
	ws.getCell("B1").value = "{{proceso.file.name}}";
	ws.mergeCells("B1:D1");
	ws.getCell("A2").value = "Equipo";
	ws.getCell("B2").value = "{{proceso.team}}";
	ws.getCell("D2").value = "Rol";
	ws.getCell("E2").value = "{{proceso.role}}";
	ws.getCell("A3").value = "Seniority";
	ws.getCell("B3").value = "{{proceso.seniority}}";
	ws.getCell("D3").value = "Vacantes";
	ws.getCell("E3").value = "{{proceso.quantity}}";
	ws.getCell("A4").value = "Estado";
	ws.getCell("B4").value = "{{proceso.status}}";
	ws.getCell("D4").value = "Contratados";
	ws.getCell("E4").value = "{{hires | count}}";
	const headers = ["Candidato", "Seniority", "Stack", "LinkedIn", "Fecha", "Speech", "Tech", "Decisión"];
	headers.forEach((h, i) => {
		const c = ws.getCell(6, i + 1);
		c.value = h;
		c.fill = HEADER_FILL;
		c.font = { bold: true, color: { argb: "FFFFFFFF" } };
	});
	const row7 = [
		"{{#each interviews}}{{interviews.interviewed}}",
		"{{interviews.interviewed.seniority}}",
		"{{interviews.interviewed.stack}}",
		"{{interviews.interviewed.linkedin}}",
		"{{interviews.date}}",
		"{{interviews.speech rating | stars}}",
		"{{interviews.tech rating | stars}}",
		"{{interviews.decision}}",
	];
	row7.forEach((v, i) => {
		const c = ws.getCell(7, i + 1);
		c.value = v;
		c.fill = ROW_FILL;
		c.border = { bottom: { style: "thin" } };
	});
	ws.getCell("E7").numFmt = "dd/mm/yyyy";
	ws.getRow(7).height = 22;
	ws.getCell("A8").value = "Promedio tech";
	ws.getCell("G8").value = { formula: "AVERAGE(G7:G7)" } as ExcelJS.CellFormulaValue;
	ws.mergeCells("A8:B8");
	ws.getCell("A9").value = "Total filas";
	ws.getCell("G9").value = { formula: "COUNTA(H7:H7)+G8*0" } as ExcelJS.CellFormulaValue;
	ws.getCell("A10").value = "Generado {{@today | date:\"dd/MM/yyyy\"}}";
	return saveWorkbook(wb);
}

/** Template of §14.2 (Consolidado.xlsx). */
export async function consolidadoTemplate(): Promise<ArrayBuffer> {
	const wb = new ExcelJS.Workbook();
	const ws = wb.addWorksheet("Consolidado");
	["Proceso", "Rol", "Prioridad", "Candidato", "Tech", "Decisión", "#"].forEach((h, i) => (ws.getCell(1, i + 1).value = h));
	[
		"{{#each proceso.interviews}}{{proceso.file.name}}",
		"{{proceso.role}}",
		"{{proceso.priority}}",
		"{{interviews.interviewed}}",
		"{{interviews.tech rating | stars}}",
		"{{interviews.decision}}",
		"{{@index}}/{{@count}}",
	].forEach((v, i) => (ws.getCell(2, i + 1).value = v));
	ws.getCell("A3").value = "Total";
	ws.getCell("E3").value = { formula: "SUM(E2:E2)" } as ExcelJS.CellFormulaValue;
	return saveWorkbook(wb);
}

export const PROCESO_DEF = `
name: Procesos activos
template: Templates/Excel/Proceso.xlsx
root:
  alias: proceso
  where: 'categories contains "Talent Acquisition Process"'
  filter: 'status != "Closed"'
  sort: 'created desc'
mode: file-per-root
relations:
  interviews:
    from: 'categories contains "Interview Evaluations"'
    on: 'process -> proceso'
    sort: 'date asc'
  hires:
    source: interviews
    filter: 'decision == "Hire"'
output:
  folder: Exports/out
  filename: '{{proceso.file.name}} - {{@today | date:"yyyy-MM-dd"}}.xlsx'
  sheetName: '{{proceso.file.name}}'
  overwrite: suffix
`;
