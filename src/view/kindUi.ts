/**
 * KindUi: the per-format content renderer hosted by OfficeView.
 */
import type { OfficeKind } from "../core/detect";
import type { OfficeSettings } from "../settings";
import type { OfficeEngine } from "../formats/engine";

export interface KindUiContext {
	/** Flip the view dirty flag (triggers the save affordances). */
	setDirty(): void;
	/** Current plugin settings. */
	settings(): OfficeSettings;
	/** Update the status line at the bottom of the view. */
	status(text: string): void;
	/** Opens the current file with the system default application. */
	openExternal(): void;
}

export interface KindUi {
	readonly kind: OfficeKind;
	render(host: HTMLElement): Promise<void> | void;
	applySettings(): void;
	destroy(): void;
}

export async function createKindUi(
	kind: OfficeKind,
	engine: OfficeEngine,
	bytes: Uint8Array,
	context: KindUiContext,
): Promise<KindUi> {
	if (kind === "xlsx" || kind === "xls") {
		const { XlsxViewUi } = await import("./xlsx/XlsxViewUi");
		if (engine.kind !== "xlsx" && engine.kind !== "xls") {
			throw new Error("grid engine required");
		}
		return new XlsxViewUi(kind, engine, context);
	}
	if (kind === "docx" && engine.kind === "docx") {
		const { DocxViewUi } = await import("./docx/DocxViewUi");
		return new DocxViewUi(engine, context, bytes);
	}
	if (kind === "pptx" && engine.kind === "pptx") {
		const { PptxViewUi } = await import("./pptx/PptxViewUi");
		return new PptxViewUi(engine, context);
	}
	if (kind === "doc" && engine.kind === "doc") {
		const { DocViewUi } = await import("./doc/DocViewUi");
		return new DocViewUi(engine, context);
	}
	throw new Error(`No editor UI for kind: ${kind}`);
}
