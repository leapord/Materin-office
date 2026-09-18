/**
 * High-fidelity read-only preview on docx-preview.
 */
import { renderAsync } from "docx-preview";

export class DocxPreview {
	private container: HTMLElement | null = null;

	async render(host: HTMLElement, bytes: Uint8Array): Promise<void> {
		host.empty();
		this.container = host.createDiv("materin-office-docx-preview");
		const blob = new Blob([bytes.slice()], {
			type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
		});
		await renderAsync(blob, this.container, undefined, {
			className: "docx",
			ignoreLastRenderedPageBreak: false,
			breakPages: true,
			inWrapper: true,
		});
	}

	destroy(): void {
		this.container?.remove();
		this.container = null;
	}
}
