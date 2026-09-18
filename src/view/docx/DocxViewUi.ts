/**
 * docx composition: mode toolbar (预览/编辑) + preview or paragraph editor +
 * find & replace panel.
 */
import type { KindUi, KindUiContext } from "../kindUi";
import { FindReplacePanel } from "../findReplace";
import type { TextDocEngine } from "../../formats/engine";
import { DocxPreview } from "./DocxPreview";
import { DocxParagraphEditor } from "./DocxParagraphEditor";

export class DocxViewUi implements KindUi {
	readonly kind = "docx" as const;

	private mode: "preview" | "edit" = "preview";
	private root: HTMLElement | null = null;
	private contentHost: HTMLElement | null = null;
	private preview = new DocxPreview();
	private editor: DocxParagraphEditor | null = null;
	private frPanel: FindReplacePanel | null = null;
	private initialBytes: Uint8Array;

	constructor(
		private readonly engine: TextDocEngine,
		private readonly context: KindUiContext,
		bytes: Uint8Array,
	) {
		this.initialBytes = bytes;
		this.mode = context.settings().docxEditMode;
	}

	async render(host: HTMLElement): Promise<void> {
		this.root = host.createDiv("materin-office-docx");
		const toolbar = this.root.createDiv("materin-office-docx-toolbar");
		toolbar.createEl("button", { text: "预览", cls: this.mode === "preview" ? "is-active" : "" })
			.addEventListener("click", () => { void this.switchMode("preview"); });
		toolbar.createEl("button", { text: "编辑", cls: this.mode === "edit" ? "is-active" : "" })
			.addEventListener("click", () => { void this.switchMode("edit"); });
		toolbar.createEl("button", { text: "查找替换", cls: "mod-muted" })
			.addEventListener("click", () => this.frPanel?.toggle());
		this.frPanel = new FindReplacePanel(this.root, {
			countMatches: (q, o) => this.engine.countMatches(q, o),
			replaceAll: (q, r, o) => this.engine.findReplaceAll(q, r, o),
			onMutated: () => this.context.setDirty(),
			refresh: () => this.refreshContent(),
		});
		this.contentHost = this.root.createDiv("materin-office-docx-content");
		await this.renderMode();
	}

	private async renderMode(): Promise<void> {
		if (!this.contentHost) {
			return;
		}
		this.editor?.destroy();
		this.editor = null;
		this.preview.destroy();
		this.contentHost.empty();
		if (this.mode === "preview") {
			await this.preview.render(this.contentHost, this.initialBytes);
		} else {
			this.editor = new DocxParagraphEditor(this.engine, {
				onMutated: () => this.context.setDirty(),
			});
			this.editor.render(this.contentHost);
		}
	}

	private async switchMode(mode: "preview" | "edit"): Promise<void> {
		if (this.mode === mode) {
			return;
		}
		this.mode = mode;
		this.syncToolbar();
		await this.renderMode();
	}

	private syncToolbar(): void {
		const buttons = this.root?.querySelectorAll(".materin-office-docx-toolbar button");
		if (!buttons || buttons.length < 2) {
			return;
		}
		buttons[0].className = this.mode === "preview" ? "is-active" : "";
		buttons[1].className = this.mode === "edit" ? "is-active" : "";
	}

	/** Re-render after find & replace mutated the model. */
	refreshContent(): void {
		if (this.mode === "edit") {
			this.editor?.rerenderAll();
			return;
		}
		if (this.contentHost) {
			void this.engine.serialize().then((bytes) => {
				void this.preview.render(this.contentHost!, bytes);
			});
		}
	}

	applySettings(): void {
		// docxEditMode applies to newly opened views.
	}

	destroy(): void {
		this.editor?.destroy();
		this.preview.destroy();
		this.root?.empty();
	}
}
