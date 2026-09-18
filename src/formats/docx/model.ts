/**
 * docx model: a paragraph is its w:p element plus the ordered w:t nodes that
 * carry its visible text. Editing only ever touches w:t text content.
 */
import { NS } from "../../core/ooxml/ns";
import { needsSpacePreserve, type SegmentNode } from "../../core/textRun";

export interface DocxParagraph {
	/** The w:p element (used for nothing but debugging identity). */
	readonly el: Element;
	/** Ordered text-bearing nodes of this paragraph. */
	readonly nodes: readonly SegmentNode[];
	/** Concatenated visible text. */
	readonly text: string;
	/** True when the paragraph lives inside a table cell. */
	readonly inTable: boolean;
}

const XML_SPACE_QNAME = "xml:space";

/** Wraps one w:t element as a textRun segment with xml:space management. */
export function makeTextNode(el: Element): SegmentNode {
	return {
		el,
		getText: () => el.textContent ?? "",
		setText: (next: string) => {
			el.textContent = next;
			if (needsSpacePreserve(next)) {
				el.setAttributeNS(NS.xml, XML_SPACE_QNAME, "preserve");
			}
		},
	};
}

/** Recreates the concatenated text after edits. */
export function paragraphText(nodes: readonly SegmentNode[]): string {
	return nodes.map((node) => node.getText()).join("");
}

/** Collects all w:t descendants of a w:p element in document order. */
export function collectTextNodes(paragraphEl: Element): SegmentNode[] {
	const found = paragraphEl.getElementsByTagNameNS(NS.w, "t");
	const out: SegmentNode[] = [];
	for (let i = 0; i < found.length; i++) {
		out.push(makeTextNode(found[i]));
	}
	return out;
}
