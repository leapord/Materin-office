/**
 * xls (BIFF8) engine on SheetJS CE. Value-level editing only: styles are kept
 * as far as SheetJS round-trips them, formulas typed in the grid are stored
 * as text (BIFF write support is limited) - the view shows a one-time banner.
 */
import * as XLSX from "xlsx";
import type {
	GridCellDisplay,
	GridCellKind,
	GridEngine,
	SheetInfo,
} from "../engine";
import { colName, formatGridValue } from "../xlsx/display";
import { parseUserInput } from "../xlsx/model";

const MIN_GRID_ROWS = 100;
const MIN_GRID_COLS = 26;
const PAD_ROWS = 100;
const PAD_COLS = 10;
const DEFAULT_COL_PX = 64;

interface SheetMeta {
	name: string;
	rows: number;
	cols: number;
}

/** SheetJS cell type narrowed to what we render. */
type SCell = {
	t?: string;
	v?: unknown;
	w?: string;
	z?: string | number;
	f?: string;
};

function cellKind(cell: SCell | undefined): GridCellKind {
	if (!cell || (cell.v === undefined && cell.t === undefined)) {
		return "empty";
	}
	if (cell.f) {
		return "formula";
	}
	switch (cell.t) {
		case "n":
			return "number";
		case "b":
			return "bool";
		case "d":
			return "date";
		default:
			return "text";
	}
}

export class XlsEngine implements GridEngine {
	readonly kind = "xls" as const;

	private wb: XLSX.WorkBook;
	private metas: SheetMeta[] = [];
	private widths: Map<string, number> = new Map();
	private _active = 0;

	private constructor(wb: XLSX.WorkBook) {
		this.wb = wb;
		this.metas = wb.SheetNames.map((name, i) => {
			const ws = wb.Sheets[name];
			const ref = ws?.["!ref"] ?? "A1:A1";
			const range = XLSX.utils.decode_range(ref);
			return {
				name,
				rows: Math.max(range.e.r + 1 + PAD_ROWS, MIN_GRID_ROWS),
				cols: Math.max(range.e.c + 1 + PAD_COLS, MIN_GRID_COLS),
			};
		});
	}

	static load(bytes: Uint8Array): XlsEngine {
		const wb = XLSX.read(new Uint8Array(bytes), {
			type: "array",
			cellStyles: true,
			cellDates: true,
		});
		return new XlsEngine(wb);
	}

	get sheets(): readonly SheetInfo[] {
		return this.metas.map((m) => ({
			name: m.name,
			rowCount: m.rows,
			colCount: m.cols,
		}));
	}

	get activeSheet(): number {
		return this._active;
	}

	selectSheet(index: number): void {
		this._active = index;
	}

	getColumnWidth(_sheet: number, col: number): number {
		const ws = this.wb.Sheets[this.wb.SheetNames[_sheet]];
		const info = (ws?.["!cols"] as Array<{ wpx?: number; wch?: number }>)?.[col - 1];
		if (info?.wpx && info.wpx > 0) {
			return Math.min(Math.round(info.wpx), 320);
		}
		if (info?.wch && info.wch > 0) {
			return Math.min(Math.round(info.wch * 7 + 5), 320);
		}
		return DEFAULT_COL_PX;
	}

	getCellDisplay(sheet: number, row: number, col: number): GridCellDisplay {
		const ws = this.wb.Sheets[this.wb.SheetNames[sheet]];
		const addr = colName(col) + String(row);
		const cell = ws?.[addr] as SCell | undefined;
		const kind = cellKind(cell);
		let text = "";
		if (cell) {
			if (kind === "formula") {
				text = cell.w ?? `=${cell.f ?? ""}`;
			} else if (cell.w) {
				text = cell.w;
			} else {
				text = formatGridValue(cell.v, cell.z);
			}
		}
		const raw = kind === "empty"
			? ""
			: kind === "formula"
				? `=${cell?.f ?? ""}`
				: safePrimitive(cell?.v);
		return {
			text,
			raw,
			kind,
			align: kind === "number" || kind === "date" || kind === "bool" ? "right" : "left",
		};
	}

	setCellValue(sheet: number, row: number, col: number, raw: string): boolean {
		const ws = this.wb.Sheets[this.wb.SheetNames[sheet]];
		if (!ws) {
			return false;
		}
		const addr = colName(col) + String(row);
		const before = (ws[addr] as SCell | undefined)?.v;
		const parsed = parseUserInput(raw);
		switch (parsed.t) {
			case "empty":
				delete ws[addr];
				break;
			case "number":
				ws[addr] = { t: "n", v: parsed.v };
				break;
			case "bool":
				ws[addr] = { t: "b", v: parsed.v };
				break;
			case "date":
				ws[addr] = { t: "d", v: parsed.v };
				break;
			case "formula":
				// BIFF8 formula write is unreliable in SheetJS CE - store as text.
				ws[addr] = { t: "s", v: `=${parsed.f}` };
				break;
			case "text":
				ws[addr] = { t: "s", v: parsed.v };
				break;
		}
		if (parsed.t !== "empty" && !ws["!ref"]) {
			ws["!ref"] = addr;
		} else if (parsed.t !== "empty") {
			this.growRef(ws, row, col);
		}
		return ws[addr] !== undefined
			? (ws[addr] as SCell).v !== before
			: before !== undefined;
	}

	private growRef(ws: XLSX.WorkSheet, row: number, col: number): void {
		const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1");
		range.e.r = Math.max(range.e.r, row - 1);
		range.e.c = Math.max(range.e.c, col - 1);
		range.s.r = Math.min(range.s.r, row - 1);
		range.s.c = Math.min(range.s.c, col - 1);
		ws["!ref"] = XLSX.utils.encode_range(range);
	}

	async serialize(): Promise<Uint8Array> {
		const out = XLSX.write(this.wb, { bookType: "xls", type: "array" }) as ArrayBuffer;
		return new Uint8Array(out);
	}
}

function safePrimitive(v: unknown): string {
	if (typeof v === "string") {
		return v;
	}
	if (typeof v === "number" || typeof v === "boolean") {
		return String(v);
	}
	return "";
}
