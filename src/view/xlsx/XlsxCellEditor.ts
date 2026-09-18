/**
 * Single-cell input overlay. Commit semantics are IME-safe: Enter/Tab only
 * commit when event.isComposing is false; blur commits unless cancelled.
 */

export interface CellRect {
	x: number;
	y: number;
	w: number;
	h: number;
}

export type CommitMove = "down" | "right" | "none";

export interface CellEditorCallbacks {
	commit(raw: string, move: CommitMove): void;
	cancel(): void;
}

export class XlsxCellEditor {
	private input: HTMLInputElement | null = null;
	private cancelled = false;

	get isOpen(): boolean {
		return this.input !== null;
	}

	open(sizer: HTMLElement, rect: CellRect, seed: string, callbacks: CellEditorCallbacks): void {
		this.close();
		this.cancelled = false;
		const input = sizer.createEl("input", "materin-office-cell-input");
		input.type = "text";
		input.value = seed;
		input.style.left = `${rect.x}px`;
		input.style.top = `${rect.y}px`;
		input.style.width = `${Math.max(rect.w, 80)}px`;
		input.style.height = `${rect.h}px`;
		this.input = input;

		input.addEventListener("keydown", (event) => {
			if (event.key === "Escape") {
				event.preventDefault();
				this.cancelled = true;
				callbacks.cancel();
				this.close();
				return;
			}
			if (event.isComposing) {
				return;
			}
			if (event.key === "Enter") {
				event.preventDefault();
				callbacks.commit(input.value, "down");
				this.close();
			} else if (event.key === "Tab") {
				event.preventDefault();
				callbacks.commit(input.value, "right");
				this.close();
			}
		});
		input.addEventListener("blur", () => {
			if (!this.cancelled) {
				callbacks.commit(input.value, "none");
			}
			this.close();
		});
		input.focus();
		input.select();
	}

	close(): void {
		this.input?.remove();
		this.input = null;
	}
}
