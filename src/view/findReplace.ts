/**
 * Find & replace panel shared by the text-document kinds. IME-safe: Enter in
 * the inputs never commits while composing.
 */
import { FindReplaceOptions } from "../formats/engine";

export interface FindReplaceHost {
	/** Count of current matches (already options-aware). */
	countMatches(query: string, options: FindReplaceOptions): number;
	/** Applies replace-all; returns the count. */
	replaceAll(query: string, replacement: string, options: FindReplaceOptions): number;
	/** Flip the dirty flag after a successful replace. */
	onMutated(): void;
	refresh(): void;
}

export class FindReplacePanel {
	private readonly root: HTMLElement;
	private queryInput: HTMLInputElement;
	private replaceInput: HTMLInputElement;
	private caseToggle: HTMLInputElement;
	private countEl: HTMLElement;
	private expanded = false;

	constructor(
		host: HTMLElement,
		private readonly engine: FindReplaceHost,
	) {
		this.root = host.createDiv("materin-office-fr");
		this.queryInput = this.root.createEl("input", "materin-office-fr-query");
		this.queryInput.type = "text";
		this.queryInput.placeholder = "查找…";
		this.replaceInput = this.root.createEl("input", "materin-office-fr-replace");
		this.replaceInput.type = "text";
		this.replaceInput.placeholder = "替换为…";
		this.caseToggle = this.root.createEl("input", "materin-office-fr-case");
		this.caseToggle.type = "checkbox";
		this.caseToggle.checked = false;
		const caseLabel = this.root.createEl("label", "materin-office-fr-case-label");
		caseLabel.setText("区分大小写");
		caseLabel.prepend(this.caseToggle);
		this.countEl = this.root.createDiv("materin-office-fr-count");
		const actions = this.root.createDiv("materin-office-fr-actions");
		actions.createEl("button", { text: "全部替换", cls: "mod-cta" })
			.addEventListener("click", () => this.applyReplaceAll());
		const hide = actions.createEl("button", { text: "收起", cls: "mod-muted" });
		hide.addEventListener("click", () => this.toggle(false));
		for (const input of [this.queryInput, this.replaceInput]) {
			input.addEventListener("input", () => this.updateCount());
		}
		this.caseToggle.addEventListener("change", () => this.updateCount());
		this.root.hide();
	}

	toggle(show?: boolean): void {
		this.expanded = show ?? !this.expanded;
		if (this.expanded) {
			this.root.show();
			this.queryInput.focus();
		} else {
			this.root.hide();
		}
		this.updateCount();
	}

	private options(): FindReplaceOptions {
		return { caseSensitive: this.caseToggle.checked };
	}

	private updateCount(): void {
		if (!this.expanded) {
			this.countEl.setText("");
			return;
		}
		const query = this.queryInput.value;
		if (query === "") {
			this.countEl.setText("");
			return;
		}
		const count = this.engine.countMatches(query, this.options());
		this.countEl.setText(`${count} 处匹配`);
	}

	private applyReplaceAll(): void {
		const query = this.queryInput.value;
		if (query === "") {
			return;
		}
		const count = this.engine.replaceAll(
			query,
			this.replaceInput.value,
			this.options(),
		);
		if (count > 0) {
			this.engine.onMutated();
			this.engine.refresh();
		}
		this.countEl.setText(`已替换 ${count} 处`);
	}
}
