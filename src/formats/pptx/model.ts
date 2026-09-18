/**
 * pptx model: a paragraph is its a:p element plus the ordered a:t nodes that
 * carry its visible text. DrawingML a:t preserves whitespace natively — no
 * xml:space handling needed (unlike docx w:t).
 */
import { NS } from "../../core/ooxml/ns";
import type { SegmentNode } from "../../core/textRun";

export interface PptxParagraph {
	/** The a:p element. */
	readonly el: Element;
	/** Ordered text-bearing nodes of this paragraph. */
	readonly nodes: SegmentNode[];
	/** Concatenated visible text. */
	readonly text: string;
}

/** One p:sp text shape on a slide. Geometry fields may be null (no xfrm). */
export interface PptxShape {
	readonly el: Element;
	readonly txBody: Element | null;
	readonly name: string;
	readonly placeholder: string | null;
	x: number | null;
	y: number | null;
	cx: number | null;
	cy: number | null;
	/** Mutable: refreshed from the DOM after structural paragraph edits. */
	paras: PptxParagraph[];
}

/** Re-collects a shape's paragraphs after add/remove edits. */
export function refreshShapeParas(shape: PptxShape): void {
	shape.paras = collectShapeParas(shape.el);
}

/** Collects the a:p paragraphs of one shape's txBody. */
export function collectShapeParas(spEl: Element): PptxParagraph[] {
	const txBody = spEl.getElementsByTagNameNS(NS.p, "txBody")[0] as
		| Element
		| undefined;
	if (!txBody) {
		return [];
	}
	const paras: PptxParagraph[] = [];
	const pEls = txBody.getElementsByTagNameNS(NS.a, "p");
	for (let i = 0; i < pEls.length; i++) {
		const nodes = collectATextNodes(pEls[i]);
		paras.push({ el: pEls[i], nodes, text: paragraphText(nodes) });
	}
	return paras;
}

/** Wraps one a:t element as a textRun segment. */
export function makeATextNode(el: Element): SegmentNode {
	return {
		el,
		getText: () => el.textContent ?? "",
		setText: (next: string) => {
			el.textContent = next;
		},
	};
}

/** Recreates the concatenated text after edits. */
export function paragraphText(nodes: readonly SegmentNode[]): string {
	return nodes.map((node) => node.getText()).join("");
}

/** Collects all a:t descendants of an a:p element in document order. */
export function collectATextNodes(paragraphEl: Element): SegmentNode[] {
	const found = paragraphEl.getElementsByTagNameNS(NS.a, "t");
	const out: SegmentNode[] = [];
	for (let i = 0; i < found.length; i++) {
		out.push(makeATextNode(found[i]));
	}
	return out;
}
