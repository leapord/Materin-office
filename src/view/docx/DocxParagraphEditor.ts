/**
 * Paragraph list editor: one row per paragraph; clicking swaps in an
 * auto-growing textarea (a plain input — contenteditable is deliberately
 * avoided for Chinese IME reliability). Commit on blur or Ctrl+Enter.
 */
import type { TextDocEngine } from "../../formats/engine";

export interface ParagraphEditorCallbacks {
	onMutated(): void;
}

export class DocxParagraphEditor {
	private root: HTMLElement | null = null;
	private editingIndex = -1;
	private textarea: HTMLTextAreaElement | null = null;
	private destroyed = false;

	constructor(
		private readonly engine: TextDocEngine,
		private readonly callbacks: ParagraphEditorCallbacks,
	) {}

	render(host: HTMLElement): void {
		host.empty();
		this.root = host.createDiv("materin-office-paras");
		this.rerenderAll();
	}

	rerenderAll(): void {
		if (!this.root || this.destroyed) {
			return;
		}
		this.editingIndex = -1;
		this.textarea = null;
		this.root.empty();
		this.engine.paragraphs.forEach((para, index) => {
			this.renderParagraph(para.text, para.inTable, index);
		});
	}

	private renderParagraph(text: string, inTable: boolean, index: number): void {
		if (!this.root) {
			return;
		}
		const row = this.root.createDiv("materin-office-para");
		if (inTable) {
			row.addClass("is-in-table");
		}
		if (text === "") {
			row.addClass("is-empty");
			row.createDiv({ text: " ", cls: "materin-office-para-text" });
		} else {
			row.createDiv({ text, cls: "materin-office-para-text" });
		}
		row.addEventListener("click", () => this.startEdit(index));
	}

	private startEdit(index: number): void {
		if (this.editingIndex === index || !this.root) {
			return;
		}
		this.commitCurrent();
		this.editingIndex = index;
		const rows = this.root.querySelectorAll(".materin-office-para");
		const row = rows[index];
		if (!row.instanceOf(HTMLElement)) {
			return;
		}
		const textEl = row.querySelector(".materin-office-para-text");
		if (textEl) {
			textEl.remove();
		}
		const original = this.engine.getParagraphText(index);
		const area = row.createEl("textarea", "materin-office-para-editor");
		area.value = original;
		this.textarea = area;
		area.addEventListener(
			"keydown",
			(event) => {
				if (event.key === "Enter" && (event.ctrlKey || event.metaKey) && !event.isComposing) {
					event.preventDefault();
					this.commitCurrent();
				} else if (event.key === "Escape") {
					event.preventDefault();
					this.cancelEdit();
				}
			},
		);
		area.addEventListener("blur", () => this.commitCurrent());
		area.addEventListener("input", () => this.autosize(area));
		this.autosize(area);
		area.focus();
	}

	private autosize(textarea: HTMLTextAreaElement): void {
		textarea.setCssProps({ height: "auto" });
		textarea.setCssProps({ height: `${textarea.scrollHeight}px` });
	}

	private commitCurrent(): void {
		if (this.editingIndex === -1 || !this.textarea) {
			return;
		}
		const index = this.editingIndex;
		const area = this.textarea;
		this.editingIndex = -1;
		this.textarea = null;
		const next = area.value;
		const prev = this.engine.getParagraphText(index);
		if (next === prev) {
			this.rerenderAll();
			return;
		}
		const changed = this.engine.replaceRange(index, 0, prev.length, next);
		if (changed) {
			this.callbacks.onMutated();
		}
		this.rerenderAll();
	}

	private cancelEdit(): void {
		if (this.editingIndex === -1) {
			return;
		}
		this.rerenderAll();
	}

	destroy(): void {
		this.destroyed = true;
		this.root?.empty();
		this.root = null;
	}
}
