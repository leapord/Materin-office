/**
 * doc composition: read-only paragraph list (legacy binary Word has no
 * editing surface) plus an external-open affordance.
 */
import type { KindUi, KindUiContext } from "../kindUi";
import type { TextExtractEngine } from "../../formats/engine";

export class DocViewUi implements KindUi {
	readonly kind = "doc" as const;

	private root: HTMLElement | null = null;

	constructor(
		private readonly engine: TextExtractEngine,
		private readonly context: KindUiContext,
	) {}

	async render(host: HTMLElement): Promise<void> {
		this.root = host.createDiv("materin-office-doc-ui");
		const toolbar = this.root.createDiv("materin-office-doc-toolbar");
		toolbar.createSpan({
			text: "只读预览",
			cls: "materin-office-doc-badge",
		});
		toolbar.createEl("button", {
			text: "用系统程序打开",
			cls: "mod-muted",
		}).addEventListener("click", () => this.context.openExternal());
		const body = this.root.createDiv("materin-office-doc-body");
		const list = body.createDiv("materin-office-paras");
		for (const paragraph of this.engine.paragraphs) {
			const row = list.createDiv("materin-office-para");
			if (paragraph === "") {
				row.addClass("is-empty");
				row.createDiv({ text: " ", cls: "materin-office-para-text" });
			} else {
				row.createDiv({ text: paragraph, cls: "materin-office-para-text" });
			}
		}
	}

	applySettings(): void {
		// read-only, no settings
	}

	destroy(): void {
		this.root?.empty();
		this.root = null;
	}
}
