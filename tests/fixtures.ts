/**
 * In-memory OOXML fixtures for vitest. Builds minimal packages from XML
 * template strings via JSZip — no binary blobs in the repo. Grows per phase
 * (docx here; xlsx/pptx builders added by their phases).
 */
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import type { Node as XmlDomNode } from "@xmldom/xmldom";
import type { XmlAdapter } from "../src/core/ooxml/xmlAdapter";

/** @xmldom/xmldom-backed adapter for tests. */
export const testAdapter: XmlAdapter = {
	parse(text: string): Document {
		const parser = new DOMParser({
			onError: (severity, msg) => {
				if (severity !== "warning") {
					throw new Error(`XML 解析失败: ${msg}`);
				}
		},
		});
		const doc = parser.parseFromString(text, "application/xml");
		if (doc.getElementsByTagName("parsererror").length > 0) {
			throw new Error("XML 解析失败");
		}
		return doc as unknown as Document;
	},
	serializeToString(doc: Document): string {
		return new XMLSerializer().serializeToString(doc as unknown as XmlDomNode);
	},
};

export async function buildZip(files: Record<string, string>): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [path, text] of Object.entries(files)) {
		zip.file(path, text);
	}
	return zip.generateAsync({ type: "uint8array" });
}

export const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`;

const ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

/** A minimal single-part OOXML package whose document.xml carries bodyXml. */
export async function buildDocx(bodyXml: string): Promise<Uint8Array> {
	return buildZip({
		"[Content_Types].xml": CONTENT_TYPES_XML,
		"_rels/.rels": ROOT_RELS_XML,
		"word/document.xml": docxDocumentXml(bodyXml),
	});
}

export function docxDocumentXml(bodyXml: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${NS_W}"><w:body>${bodyXml}</w:body></w:document>`;
}

/** One paragraph with one run holding text. */
export function docxParagraph(text: string): string {
	return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

// ── pptx builders ────────────────────────────────────────────────────────────

export const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const NS_R =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const REL_SLIDE = `${NS_R}/slide`;
export const REL_SLIDE_LAYOUT = `${NS_R}/slideLayout`;

/** presentation.xml with a sldIdLst referencing the given rIds in order. */
export function pptxPresentationXml(rIds: string[]): string {
	const ids = rIds
		.map((rId, i) => `<p:sldId id="${256 + i}" r:id="${rId}"/>`)
		.join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:sldIdLst>${ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`;
}

/** presentation rels mapping rIds to slide part targets (relative to ppt/). */
export function pptxPresentationRels(targets: Record<string, string>): string {
	const rels = Object.entries(targets)
		.map(
			([rId, target]) =>
				`<Relationship Id="${rId}" Type="${REL_SLIDE}" Target="${target}"/>`,
		)
		.join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

/** A slide part whose shape tree carries the given p:sp XML fragments. */
export function pptxSlideXml(spXml: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${spXml}</p:spTree></p:cSld></p:sld>`;
}

export interface PptxTextBoxOptions {
	name: string;
	/** a:off/a:ext in EMU; omit for no local xfrm. */
	x?: number;
	y?: number;
	cx?: number;
	cy?: number;
	/** p:ph type; omit for a plain text box. */
	phType?: string;
	phIdx?: string;
	/** a:p XML fragments; defaults to one empty paragraph. */
	paragraphs?: string;
}

/** One p:sp text shape. */
export function pptxTextBox(options: PptxTextBoxOptions): string {
	const { name, x, y, cx, cy, phType, phIdx, paragraphs } = options;
	const ph =
		phType === undefined && phIdx === undefined
			? ""
			: `<p:ph${phType ? ` type="${phType}"` : ""}${phIdx ? ` idx="${phIdx}"` : ""}/>`;
	const xfrm =
		x === undefined
			? ""
			: `<a:xfrm><a:off x="${x}" y="${y ?? 0}"/><a:ext cx="${cx ?? 0}" cy="${cy ?? 0}"/></a:xfrm>`;
	return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs ?? "<a:p/>"}</p:txBody></p:sp>`;
}

/** One paragraph with one run holding text. */
export function pptxParagraph(text: string): string {
	return `<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`;
}

/** A slide layout providing geometry for a title placeholder. */
export function pptxLayoutXml(titleXfrm: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr>${titleXfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`;
}

/** Slide rels pointing at a layout (target relative to ppt/slides/). */
export function pptxSlideRels(target: string): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="${target}"/></Relationships>`;
}

/** A minimal pptx package from part map; files are passed through verbatim. */
export async function buildPptx(files: Record<string, string>): Promise<Uint8Array> {
	return buildZip(files);
}
