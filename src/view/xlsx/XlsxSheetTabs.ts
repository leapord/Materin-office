/**
 * Bottom sheet-tab strip.
 */
import type { GridEngine } from "../../formats/engine";

export class XlsxSheetTabs {
	private readonly root: HTMLElement;
	private readonly tabs: HTMLElement[] = [];

	constructor(
		host: HTMLElement,
		private readonly engine: GridEngine,
		private onSelect: (index: number) => void,
	) {
		this.root = host.createDiv("materin-office-sheet-tabs");
	}

	render(active: number): void {
		this.root.empty();
		this.tabs.length = 0;
		this.engine.sheets.forEach((sheet, index) => {
			const tab = this.root.createEl("button", {
				text: sheet.name,
				cls: index === active
					? "materin-office-sheet-tab is-active"
					: "materin-office-sheet-tab",
			});
			tab.addEventListener("click", () => this.onSelect(index));
			this.tabs.push(tab);
		});
	}
}
