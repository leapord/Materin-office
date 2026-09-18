/**
 * document.xml -> DocxParagraph[]. Every w:p descendant of w:body is
 * collected (table cells included); headers/footers live in separate parts
 * and are not editable in v1.
 */
import { NS } from "../../core/ooxml/ns";
import type { SegmentNode } from "../../core/textRun";
import { collectTextNodes, paragraphText, type DocxParagraph } from "./model";

const W_BODY = "body";
const W_P = "p";
const W_TC = "tc";

/** Collects paragraphs in document order, tagging table-cell membership. */
export function parseDocumentBody(doc: Document): DocxParagraph[] {
	const bodies = doc.getElementsByTagNameNS(NS.w, W_BODY);
	if (bodies.length === 0) {
		return [];
	}
	const out: DocxParagraph[] = [];
	const seen = new Set<Element>();
	collectFrom(bodies[0], false, out, seen);
	return out;
}

function collectFrom(
	root: Element,
	inTable: boolean,
	out: DocxParagraph[],
	seen: Set<Element>,
): void {
	const paragraphs = root.getElementsByTagNameNS(NS.w, W_P);
	const cells = root.getElementsByTagNameNS(NS.w, W_TC);
	const tableCells: Element[] = [];
	if (!inTable) {
		for (let i = 0; i < cells.length; i++) {
			tableCells.push(cells[i]);
		}
	}
	for (let i = 0; i < paragraphs.length; i++) {
		const el = paragraphs[i];
		if (seen.has(el)) {
			continue;
		}
		seen.add(el);
		const nested = tableCells.some((tc) => tc.contains(el));
		const nodes: SegmentNode[] = collectTextNodes(el);
		out.push({
			el,
			nodes,
			text: paragraphText(nodes),
			inTable: inTable || nested,
		});
	}
}
