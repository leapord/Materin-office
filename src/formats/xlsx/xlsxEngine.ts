/**
 * xlsx engine on ExcelJS. The workbook is loaded fully; edits are value-level
 * through the GridEngine contract. Serialization re-emits the whole workbook,
 * so a pre-scan (scanXlsxRiskParts) warns about parts ExcelJS may drop.
 */
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { toArrayBuffer } from "../../core/bufferUtils";
import type {
	GridCellDisplay,
	GridCellKind,
	GridEngine,
	SheetInfo,
} from "../engine";
import { formatGridValue } from "./display";
import { parseUserInput } from "./model";

/** Parts ExcelJS drops on write-back — surface them as a persistent warning. */
export const XLSX_RISK_PREFIXES = [
	"xl/charts/",
	"xl/media/",
	"xl/pivotCache/",
];

/** JSZip pre-scan for parts ExcelJS is known to drop on re-serialization. */
export async function scanXlsxRiskParts(bytes: Uint8Array): Promise<string[]> {
	const zip = await JSZip.loadAsync(toArrayBuffer(bytes));
	const hits = new Set<string>();
	for (const path of Object.keys(zip.files)) {
		for (const prefix of XLSX_RISK_PREFIXES) {
			if (path.startsWith(prefix)) {
				hits.add(prefix.replace(/\/$/, ""));
			}
		}
	}
	return [...hits].sort();
}

/** Fallback column width (px) when the sheet stores none. */
export const DEFAULT_COL_WIDTH = 64;
const CHAR_PX = 7;
const CHAR_PAD_PX = 5;
const MIN_GRID_ROWS = 100;
const MIN_GRID_COLS = 26;
const PAD_ROWS = 100;
const PAD_COLS = 10;


function classifyKind(value: ExcelJS.CellValue): GridCellKind {
	if (value === null || value === undefined) {
		return "empty";
	}
	if (typeof value === "object" && value !== null && "formula" in value) {
		return "formula";
	}
	if (value instanceof Date) {
		return "date";
	}
	if (typeof value === "number") {
		return "number";
	}
	if (typeof value === "boolean") {
		return "bool";
	}
	return "text";
}

export class XlsxEngine implements GridEngine {
	readonly kind = "xlsx" as const;

	private constructor(
		private readonly wb: ExcelJS.Workbook,
		private readonly risks: string[],
	) {}

	private _active = 0;

	get activeSheet(): number {
		return this._active;
	}

	selectSheet(index: number): void {
		this._active = index;
	}


	static async load(bytes: Uint8Array): Promise<XlsxEngine> {
		const risks = await scanXlsxRiskParts(bytes);
		const wb = new ExcelJS.Workbook();
		await wb.xlsx.load(toArrayBuffer(bytes));
		return new XlsxEngine(wb, risks);
	}

	get sheets(): readonly SheetInfo[] {
		return this.wb.worksheets.map((ws) => {
			const rows = Math.max(
				ws.actualRowCount + PAD_ROWS,
				MIN_GRID_ROWS,
			);
			const cols = Math.max(
				ws.actualColumnCount + PAD_COLS,
				MIN_GRID_COLS,
			);
			return { name: ws.name, rowCount: rows, colCount: cols };
		});
	}

	private get sheet(): ExcelJS.Worksheet {
		return this.wb.worksheets[this.activeSheet];
	}

	private cellAt(row: number, col: number): ExcelJS.Cell | undefined {
		const found = this.sheet.findRow(row);
		if (!found) {
			return undefined;
		}
		return found.findCell(col);
	}

	getCellDisplay(sheet: number, row: number, col: number): GridCellDisplay {
		const ws = this.wb.worksheets[sheet];
		const found = ws.findRow(row)?.findCell(col);
		const value = found?.value ?? null;
		const kind = classifyKind(value);
		let text = "";
		if (found && kind !== "empty") {
			const numFmt = found.numFmt;
			if (kind === "formula") {
				const fv = value as { formula?: string; result?: ExcelJS.CellValue };
				text = fv.result !== null && fv.result !== undefined
					? formatGridValue(fv.result, numFmt)
					: `=${fv.formula ?? ""}`;
			} else {
				text = formatGridValue(value, numFmt);
			}
		}
		return {
			text,
			raw: this.editorRaw(value, kind),
			kind,
			align: kind === "number" || kind === "date" || kind === "bool" || (kind === "formula" && typeof (value as { result?: unknown }).result === "number") ? "right" : "left",
		};
	}

	private editorRaw(value: ExcelJS.CellValue, kind: GridCellKind): string {
		switch (kind) {
			case "empty":
				return "";
			case "formula":
				return `=${(value as { formula?: string }).formula ?? ""}`;
			case "bool":
				return (value as boolean) ? "true" : "false";
			case "date":
				return toIsoDate(value as Date);
			case "number":
				return typeof value === "number" ? String(value) : "";
			default:
				return typeof value === "object" && value !== null
					? extractRichText(value)
					: String(value ?? "");
		}
	}

	setCellValue(sheet: number, row: number, col: number, raw: string): boolean {
		const ws = this.wb.worksheets[sheet];
		const cell = ws.getCell(row, col);
		const parsed = parseUserInput(raw);
		const before = this.getCellDisplay(sheet, row, col).raw;
		switch (parsed.t) {
			case "empty":
				cell.value = null;
				break;
			case "number":
				cell.value = parsed.v;
				break;
			case "bool":
				cell.value = parsed.v;
				break;
			case "date":
				cell.value = parsed.v;
				break;
			case "formula":
				cell.value = { formula: parsed.f };
				break;
			case "text":
				cell.value = parsed.v;
				break;
		}
		return this.getCellDisplay(sheet, row, col).raw !== before;
	}

	getColumnWidth(_sheet: number, col: number): number {
		const ws = this.wb.worksheets[_sheet];
		const stored = ws.getColumn(col).width;
		if (stored && stored > 0) {
			return Math.min(Math.max(Math.round(stored * CHAR_PX + CHAR_PAD_PX), 40), 320);
		}
		return DEFAULT_COL_WIDTH;
	}

	/** Parts ExcelJS may drop — the view shows a warning banner when non-empty. */
	get riskParts(): string[] {
		return [...this.risks];
	}

	async serialize(): Promise<Uint8Array> {
		const out = await this.wb.xlsx.writeBuffer();
		return new Uint8Array(out);
	}
}

function toIsoDate(date: Date): string {
	const pad = (n: number) => String(n).padStart(2, "0");
	return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function extractRichText(value: object): string {
	const rich = (value as { richText?: Array<{ text: string }> }).richText;
	if (Array.isArray(rich)) {
		return rich.map((run) => run.text).join("");
	}
	const inner = (value as { text?: unknown }).text;
	return typeof inner === "string" ? inner : "";
}
