/**
 * pptx canvas: renders one slide's text shapes positioned per their xfrm
 * (EMU → percent of slide size), with click-to-edit textareas using the same
 * IME-safe pattern as the docx paragraph editor.
 */
import type { SlideCanvasEngine, SlideShapeInfo } from "../../formats/engine";

export interface PptxCanvasCallbacks {
	onMutated(): void;
	status(text: string): void;
}

export class PptxCanvas {
	private root: HTMLElement | null = null;
	private slideIndex = 0;
	private editing: { slide: number; shape: number } | null = null;
	private textarea: HTMLTextAreaElement | null = null;
	private destroyed = false;

	constructor(
		private readonly engine: SlideCanvasEngine,
		private readonly callbacks: PptxCanvasCallbacks,
	) {}

	render(host: HTMLElement): void {
		host.empty();
		this.root = host.createDiv("materin-office-pptx");
		this.rerenderAll();
	}

	get page(): number {
		return this.slideIndex;
	}

	setSlide(index: number): void {
		this.slideIndex = Math.max(
			0,
			Math.min(this.engine.slides.length - 1, index),
		);
		this.rerenderAll();
	}

	rerenderAll(): void {
		if (!this.root || this.destroyed) {
			return;
		}
		this.editing = null;
		this.textarea = null;
		this.root.empty();
		const widthEmu = Math.max(this.engine.slideSize.widthEmu, 1);
		const heightEmu = Math.max(this.engine.slideSize.heightEmu, 1);
		const slide = this.root.createDiv("materin-office-pptx-slide");
		slide.style.aspectRatio = `${widthEmu} / ${heightEmu}`;
		const shapes = this.engine.slides[this.slideIndex] ?? [];
		let flowCount = 0;
		shapes.forEach((shape, index) => {
			if (shape.x === null) {
				flowCount += 1;
			} else {
				this.renderCanvasShape(slide, shape, index);
			}
		});
		if (flowCount > 0) {
			this.renderFlowSection(shapes);
		}
		this.callbacks.status(
			`幻灯片 ${this.slideIndex + 1} / ${this.engine.slides.length}`,
		);
	}

	private renderCanvasShape(
		slideEl: HTMLElement,
		shape: SlideShapeInfo,
		index: number,
	): void {
		const box = slideEl.createDiv("materin-office-pptx-shape");
		if (shape.placeholder === "title" || shape.placeholder === "ctrTitle") {
			box.addClass("is-title");
		}
		const w = this.engine.slideSize.widthEmu;
		const h = this.engine.slideSize.heightEmu;
		box.style.left = `${((shape.x ?? 0) / w) * 100}%`;
		box.style.top = `${((shape.y ?? 0) / h) * 100}%`;
		box.style.width = `${((shape.cx ?? 0) / w) * 100}%`;
		box.style.height = `${((shape.cy ?? 0) / h) * 100}%`;
		this.renderShapeContent(box, shape, index);
	}

	private renderFlowSection(shapes: readonly SlideShapeInfo[]): void {
		if (!this.root) {
			return;
		}
		const section = this.root.createDiv("materin-office-pptx-flow");
		section.createDiv({
			text: "未定位文本框",
			cls: "materin-office-pptx-flow-label",
		});
		shapes.forEach((shape, index) => {
			if (shape.x === null) {
				const box = section.createDiv("materin-office-pptx-flow-shape");
				this.renderShapeContent(box, shape, index);
			}
		});
	}

	private renderShapeContent(
		box: HTMLElement,
		shape: SlideShapeInfo,
		index: number,
	): void {
		box.title = shape.name;
		box.dataset.shapeIndex = String(index);
		if (shape.text === "") {
			box.addClass("is-empty");
			box.createDiv({ text: " ", cls: "materin-office-pptx-shape-text" });
		} else {
			box.createDiv({ text: shape.text, cls: "materin-office-pptx-shape-text" });
		}
		box.addEventListener("click", () => this.startEdit(index));
	}

	private startEdit(index: number): void {
		if (this.editing || !this.root) {
			return;
		}
		this.editing = { slide: this.slideIndex, shape: index };
		const box = this.root.querySelector(
			`[data-shape-index="${index}"]`,
		);
		if (!(box instanceof HTMLElement)) {
			this.editing = null;
			return;
		}
		const textEl = box.querySelector(".materin-office-pptx-shape-text");
		if (textEl) {
			textEl.remove();
		}
		const area = box.createEl("textarea", "materin-office-pptx-shape-editor");
		area.value = this.engine.getShapeText(this.slideIndex, index);
		this.textarea = area;
		area.addEventListener("keydown", (event) => {
			if (event.key === "Escape") {
				event.preventDefault();
				this.cancelEdit();
				return;
			}
			if (
				event.key === "Enter" &&
				(event.ctrlKey || event.metaKey) &&
				!event.isComposing
			) {
				event.preventDefault();
				this.commitCurrent();
			}
		});
		area.addEventListener("blur", () => this.commitCurrent());
		area.focus();
	}

	private commitCurrent(): void {
		if (!this.editing || !this.textarea) {
			return;
		}
		const { slide, shape } = this.editing;
		const next = this.textarea.value;
		this.editing = null;
		this.textarea = null;
		const prev = this.engine.getShapeText(slide, shape);
		if (next === prev) {
			this.rerenderAll();
			return;
		}
		if (this.engine.replaceShapeText(slide, shape, next)) {
			this.callbacks.onMutated();
		}
		this.rerenderAll();
	}

	private cancelEdit(): void {
		if (!this.editing) {
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
