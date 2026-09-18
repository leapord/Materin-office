/**
 * pptx composition: slide navigation toolbar + canvas + find & replace panel
 * (reusing the engine-agnostic FindReplacePanel).
 */
import type { KindUi, KindUiContext } from "../kindUi";
import { FindReplacePanel } from "../findReplace";
import type { SlideCanvasEngine } from "../../formats/engine";
import { PptxCanvas } from "./PptxCanvas";

export class PptxViewUi implements KindUi {
	readonly kind = "pptx" as const;

	private root: HTMLElement | null = null;
	private contentHost: HTMLElement | null = null;
	private canvas: PptxCanvas | null = null;
	private frPanel: FindReplacePanel | null = null;
	private prevBtn: HTMLButtonElement | null = null;
	private nextBtn: HTMLButtonElement | null = null;
	private pageLabel: HTMLElement | null = null;

	constructor(
		private readonly engine: SlideCanvasEngine,
		private readonly context: KindUiContext,
	) {}

	async render(host: HTMLElement): Promise<void> {
		this.root = host.createDiv("materin-office-pptx-ui");
		const toolbar = this.root.createDiv("materin-office-pptx-toolbar");
		this.prevBtn = toolbar.createEl("button", { text: "◀" });
		this.prevBtn.addEventListener("click", () => this.turnPage(-1));
		this.pageLabel = toolbar.createDiv("materin-office-pptx-page");
		this.nextBtn = toolbar.createEl("button", "mod-muted");
		this.nextBtn.textContent = "▶";
		this.nextBtn.addEventListener("click", () => this.turnPage(1));
		const frButton = toolbar.createEl("button", {
			text: "查找替换",
			cls: "mod-muted",
		});
		frButton.addEventListener("click", () => this.frPanel?.toggle());
		this.frPanel = new FindReplacePanel(this.root, {
			countMatches: (q, o) => this.engine.countMatches(q, o),
			replaceAll: (q, r, o) => this.engine.findReplaceAll(q, r, o),
			onMutated: () => this.context.setDirty(),
			refresh: () => this.canvas?.rerenderAll(),
		});
		this.contentHost = this.root.createDiv("materin-office-pptx-content");
		await this.renderContent();
	}

	private async renderContent(): Promise<void> {
		if (!this.contentHost) {
			return;
		}
		this.canvas?.destroy();
		this.canvas = new PptxCanvas(this.engine, {
			onMutated: () => this.context.setDirty(),
			status: (text) => this.context.status(text),
		});
		this.canvas.render(this.contentHost);
		this.syncPager();
	}

	private turnPage(delta: number): void {
		if (!this.canvas) {
			return;
		}
		this.canvas.setSlide(this.canvas.page + delta);
		this.syncPager();
	}

	private syncPager(): void {
		const total = this.engine.slides.length;
		const page = (this.canvas?.page ?? 0) + 1;
		if (this.pageLabel) {
			this.pageLabel.setText(`${page} / ${total}`);
		}
		if (this.prevBtn) {
			this.prevBtn.disabled = page <= 1;
		}
		if (this.nextBtn) {
			this.nextBtn.disabled = page >= total;
		}
	}

	applySettings(): void {
		// no pptx-specific settings yet
	}

	destroy(): void {
		this.canvas?.destroy();
		this.canvas = null;
		this.root?.empty();
	}
}
