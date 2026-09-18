/**
 * Spreadsheet composition: formula bar + virtual grid + sheet tabs.
 * Implements the KindUi contract for xlsx and xls.
 */
import type { OfficeKind } from "../../core/detect";
import type { GridEngine } from "../../formats/engine";
import type { KindUi, KindUiContext } from "../kindUi";
import { XlsxGrid } from "./XlsxGrid";
import { XlsxSheetTabs } from "./XlsxSheetTabs";

export class XlsxViewUi implements KindUi {
	readonly kind: OfficeKind;

	private grid: XlsxGrid | null = null;
	private formulaInput: HTMLInputElement | null = null;
	private addrBox: HTMLElement | null = null;
	private tabs: XlsxSheetTabs | null = null;
	private updatingFormula = false;

	constructor(
		kind: OfficeKind,
		private readonly engine: GridEngine,
		private readonly context: KindUiContext,
	) {
		this.kind = kind;
	}

	async render(host: HTMLElement): Promise<void> {
		const root = host.createDiv("materin-office-xlsx");
		if (this.kind === "xls") {
			root.createDiv({
				text: "xls 为值级编辑：样式与公式写回有限，建议转存为 xlsx 获得完整体验。",
				cls: "materin-office-banner materin-office-banner-warning",
			});
		}
		if (this.kind === "xlsx" && this.riskNote) {
			root.createDiv({
				text: this.riskNote,
				cls: "materin-office-banner materin-office-banner-warning",
			});
		}
		this.renderFormulaBar(root);
		const gridHost = root.createDiv("materin-office-grid-host");
		this.grid = new XlsxGrid(gridHost, this.engine, {
			onSelect: (addr, raw) => this.updateFormulaBar(addr, raw),
			onMutated: () => this.context.setDirty(),
			status: (text) => this.context.status(text),
		});
		this.tabs = new XlsxSheetTabs(root, this.engine, (index) => {
			this.grid?.loadSheet(index);
			this.tabs?.render(index);
		});
		this.tabs.render(this.engine.activeSheet);
		this.grid.loadSheet(this.engine.activeSheet);
	}

	/** ExcelJS write-back risk summary (charts/media/pivot caches). */
	private get riskNote(): string | null {
		if ("riskParts" in this.engine) {
			const parts = (this.engine as { riskParts: string[] }).riskParts;
			if (parts.length > 0) {
				return `检测到 ${parts.join("、")}：写回可能丢失对应内容（ExcelJS 已知限制），首次保存会自动备份原文件。`;
			}
		}
		return null;
	}

	private renderFormulaBar(root: HTMLElement): void {
		const bar = root.createDiv("materin-office-formula-bar");
		this.addrBox = bar.createDiv("materin-office-formula-addr");
		this.addrBox.setText("A1");
		this.formulaInput = bar.createEl("input", "materin-office-formula-input");
		this.formulaInput.type = "text";
		this.formulaInput.addEventListener("keydown", (event) => {
			if (event.isComposing || !this.formulaInput) {
				return;
			}
			if (event.key === "Enter" || event.key === "Tab") {
				event.preventDefault();
				this.commitFormula(this.formulaInput.value);
			} else if (event.key === "Escape") {
				event.preventDefault();
				this.refreshFormulaBox();
			}
		});
	}

	private commitFormula(raw: string): void {
		this.grid?.applyRaw(raw);
	}

	private updateFormulaBar(addr: string, raw: string): void {
		if (this.addrBox) {
			this.addrBox.setText(addr);
		}
		if (this.formulaInput) {
			this.updatingFormula = true;
			this.formulaInput.value = raw;
			this.updatingFormula = false;
		}
	}

	/** Re-reads the selected cell into the formula bar (after commit/Esc). */
	refreshFormulaBox(): void {
		this.grid?.refocus();
	}

	applySettings(): void {
		// Settings relevant to the grid land in a later phase.
	}

	destroy(): void {
		this.grid?.destroy();
	}
}
