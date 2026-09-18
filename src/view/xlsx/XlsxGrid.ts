/**
 * Virtual-scrolling spreadsheet grid: sticky row/column headers, windowed
 * cell rendering, keyboard navigation and IME-safe cell editing.
 */
import type { GridEngine } from "../../formats/engine";
import { toA1 } from "../../formats/xlsx/display";
import { XlsxCellEditor, type CommitMove } from "./XlsxCellEditor";

export const ROW_HEIGHT = 24;
const OVERSCAN_ROWS = 6;
const OVERSCAN_COLS = 2;
export interface GridCallbacks {
	onSelect(addr: string, raw: string): void;
	onMutated(): void;
	status(text: string): void;
}

interface Selection {
	row: number;
	col: number;
}

export class XlsxGrid {
	private readonly scroller: HTMLElement;
	private readonly sizer: HTMLElement;
	private readonly colHeader: HTMLElement;
	private readonly rowHeader: HTMLElement;
	private readonly editor = new XlsxCellEditor();

	private colOffsets: number[] = [0];
	private totalWidth = 0;
	private totalHeight = 0;
	private rows = 0;
	private cols = 0;
	private sel: Selection = { row: 1, col: 1 };
	private renderPending = false;
	private destroyed = false;

	constructor(
		host: HTMLElement,
		private readonly engine: GridEngine,
		private readonly callbacks: GridCallbacks,
	) {
		const grid = host.createDiv("materin-office-grid");
		this.colHeader = grid.createDiv("materin-office-grid-colheader");
		this.rowHeader = grid.createDiv("materin-office-grid-rowheader");
		this.scroller = grid.createDiv("materin-office-grid-scroller");
		this.scroller.setAttribute("tabindex", "0");
		this.sizer = this.scroller.createDiv("materin-office-grid-sizer");

		this.scroller.addEventListener("scroll", () => this.scheduleRender());
		this.scroller.addEventListener("mousedown", (e) => this.onMouseDown(e));
		this.scroller.addEventListener("dblclick", () => this.startEdit());
		this.scroller.addEventListener("keydown", (e) => this.onKeyDown(e));
	}

	/** Re-initializes the grid for one sheet. */
	loadSheet(index: number): void {
		this.engine.selectSheet(index);
		const info = this.engine.sheets[index];
		this.rows = info.rowCount;
		this.cols = info.colCount;
		this.colOffsets = [0];
		let acc = 0;
		for (let c = 1; c <= this.cols; c++) {
			acc += this.engine.getColumnWidth(index, c);
			this.colOffsets.push(acc);
		}
		this.totalWidth = acc;
		this.totalHeight = this.rows * ROW_HEIGHT;
		this.sel = { row: 1, col: 1 };
		this.scroller.scrollTop = 0;
		this.scroller.scrollLeft = 0;
		this.sizer.setCssStyles({
			width: `${this.totalWidth}px`,
			height: `${this.totalHeight}px`,
		});
		this.render();
		this.callbacks.onSelect(this.addr(), this.rawAt(this.sel));
		this.scroller.focus();
	}

	refresh(): void {
		this.render();
	}

	/** Applies formula-bar input to the current selection. */
	applyRaw(raw: string): void {
		this.applyCell(raw);
		this.refocus();
	}

	/** Syncs the formula bar and returns focus to the grid. */
	refocus(): void {
		this.callbacks.onSelect(this.addr(), this.rawAt(this.sel));
		this.scroller.focus();
	}

	private addr(): string {
		return toA1(this.sel.row, this.sel.col);
	}

	private rawAt(sel: Selection): string {
		return this.engine.getCellDisplay(this.engine.activeSheet, sel.row, sel.col).raw;
	}

	private scheduleRender(): void {
		if (this.renderPending || this.destroyed) {
			return;
		}
		this.renderPending = true;
		window.requestAnimationFrame(() => {
			this.renderPending = false;
			this.render();
		});
	}

	private visibleWindow(): { r0: number; r1: number; c0: number; c1: number } {
		const top = this.scroller.scrollTop;
		const left = this.scroller.scrollLeft;
		const h = this.scroller.clientHeight;
		const w = this.scroller.clientWidth;
		const r0 = Math.max(1, Math.floor(top / ROW_HEIGHT) - OVERSCAN_ROWS + 1);
		const r1 = Math.min(this.rows, Math.ceil((top + h) / ROW_HEIGHT) + OVERSCAN_ROWS);
		let c0 = 1;
		while (c0 <= this.cols && this.colOffsets[c0] < left) {
			c0++;
		}
		c0 = Math.max(1, c0 - OVERSCAN_COLS);
		let c1 = c0;
		while (c1 <= this.cols && this.colOffsets[c1 - 1] < left + w) {
			c1++;
		}
		c1 = Math.min(this.cols, c1 + OVERSCAN_COLS - 1);
		return { r0, r1, c0, c1 };
	}

	private render(): void {
		if (this.destroyed) {
			return;
		}
		const sheet = this.engine.activeSheet;
		const { r0, r1, c0, c1 } = this.visibleWindow();

		this.sizer.empty();
		for (let r = r0; r <= r1; r++) {
			for (let c = c0; c <= c1; c++) {
				const cell = this.engine.getCellDisplay(sheet, r, c);
				const el = this.sizer.createDiv("materin-office-grid-cell");
				el.setText(cell.text);
				el.setCssStyles({
					left: `${this.colOffsets[c - 1]}px`,
					top: `${(r - 1) * ROW_HEIGHT}px`,
					width: `${this.colOffsets[c] - this.colOffsets[c - 1]}px`,
					height: `${ROW_HEIGHT}px`,
					textAlign: cell.align,
				});
				if (c === this.sel.col && r === this.sel.row) {
					el.addClass("is-selected");
				}
			}
		}
		this.renderHeaders(r0, r1, c0, c1);
	}

	private renderHeaders(r0: number, r1: number, c0: number, c1: number): void {
		this.colHeader.empty();
		this.rowHeader.empty();
		for (let c = c0; c <= c1; c++) {
			const el = this.colHeader.createDiv("materin-office-grid-hcell");
			el.setText(toA1(0, c).replace(/\d+$/, ""));
			el.setCssStyles({
				left: `${this.colOffsets[c - 1]}px`,
				width: `${this.colOffsets[c] - this.colOffsets[c - 1]}px`,
			});
		}
		for (let r = r0; r <= r1; r++) {
			const el = this.rowHeader.createDiv("materin-office-grid-hcell");
			el.setText(String(r));
			el.setCssStyles({ top: `${(r - 1) * ROW_HEIGHT}px` });
		}
		this.colHeader.setCssStyles({ transform: `translateX(${-this.scroller.scrollLeft}px)` });
		this.rowHeader.setCssStyles({ transform: `translateY(${-this.scroller.scrollTop}px)` });
	}

	private cellRect(row: number, col: number): { x: number; y: number; w: number; h: number } {
		return {
			x: this.colOffsets[col - 1],
			y: (row - 1) * ROW_HEIGHT,
			w: this.colOffsets[col] - this.colOffsets[col - 1],
			h: ROW_HEIGHT,
		};
	}

	private onMouseDown(event: MouseEvent): void {
		const target = (event.target as HTMLElement).closest?.(".materin-office-grid-cell");
		if (!(target instanceof HTMLElement)) {
			return;
		}
		const x = parseFloat(target.style.left) || 0;
		const y = parseFloat(target.style.top) || 0;
		const sel = this.selFromXY(x, y);
		if (sel && (sel.row !== this.sel.row || sel.col !== this.sel.col)) {
			this.moveSelection(sel);
		}
		this.scroller.focus();
	}

	private selFromXY(x: number, y: number): Selection | null {
		let col = 1;
		while (col <= this.cols && this.colOffsets[col] <= x) {
			col++;
		}
		const row = Math.floor(y / ROW_HEIGHT) + 1;
		if (row < 1 || row > this.rows || col < 1 || col > this.cols) {
			return null;
		}
		return { row, col };
	}

	private moveSelection(sel: Selection): void {
		this.sel = {
			row: Math.min(Math.max(sel.row, 1), this.rows),
			col: Math.min(Math.max(sel.col, 1), this.cols),
		};
		this.rerenderSelection();
		this.callbacks.onSelect(this.addr(), this.rawAt(this.sel));
	}

	private rerenderSelection(): void {
		const prev = this.sizer.querySelector(".materin-office-grid-cell.is-selected");
		if (prev instanceof HTMLElement) {
			prev.removeClass("is-selected");
		}
		const x = this.colOffsets[this.sel.col - 1];
		const y = (this.sel.row - 1) * ROW_HEIGHT;
		const hit = this.sizer.querySelector(
			`.materin-office-grid-cell[style*="left: ${x}px"][style*="top: ${y}px"]`,
		);
		if (hit instanceof HTMLElement) {
			hit.addClass("is-selected");
		} else {
			this.render();
		}
	}

	private onKeyDown(event: KeyboardEvent): void {
		if (this.editor.isOpen) {
			return;
		}
		const step = (dr: number, dc: number) => {
			event.preventDefault();
			this.moveSelection({ row: this.sel.row + dr, col: this.sel.col + dc });
		};
		switch (event.key) {
			case "ArrowUp":
				return step(-1, 0);
			case "ArrowDown":
				return step(1, 0);
			case "ArrowLeft":
				return step(0, -1);
			case "ArrowRight":
				return step(0, 1);
			case "PageDown":
				return step(20, 0);
			case "PageUp":
				return step(-20, 0);
			case "Home":
				return step(0, -this.sel.col + 1);
			case "End":
				return step(0, this.cols - this.sel.col);
			case "Enter":
			case "F2":
				event.preventDefault();
				return this.startEdit();
			case "Delete":
			case "Backspace":
				event.preventDefault();
				this.applyCell("");
				return;
			default:
				break;
		}
		if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
			event.preventDefault();
			this.startEdit(event.key);
		}
	}

	private startEdit(seed?: string): void {
		const rect = this.cellRect(this.sel.row, this.sel.col);
		const initial = seed ?? this.rawAt(this.sel);
		this.editor.open(this.sizer, rect, initial, {
			commit: (raw, move) => this.commitEdit(raw, move),
			cancel: () => this.scroller.focus(),
		});
	}

	private commitEdit(raw: string, move: CommitMove): void {
		if (this.applyCell(raw)) {
			this.callbacks.onMutated();
		}
		if (move === "down") {
			this.moveSelection({ row: this.sel.row + 1, col: this.sel.col });
		} else if (move === "right") {
			this.moveSelection({ row: this.sel.row, col: this.sel.col + 1 });
		} else {
			this.scroller.focus();
		}
	}

	/** Applies raw input to the selected cell; returns true when it changed. */
	private applyCell(raw: string): boolean {
		const sheet = this.engine.activeSheet;
		const { row, col } = this.sel;
		const changed = this.engine.setCellValue(sheet, row, col, raw);
		if (changed) {
			this.render();
			this.callbacks.onMutated();
		}
		return changed;
	}

	destroy(): void {
		this.destroyed = true;
		this.editor.close();
	}
}
