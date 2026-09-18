/**
 * pptx parse: slide ordering from presentation.xml + rels, and shape parsing
 * from slide parts. Geometry inheritance resolves a slide placeholder's
 * position from the slide layout when the slide has no local xfrm.
 */
import { NS } from "../../core/ooxml/ns";
import type { OoxmlPackage } from "../../core/ooxml/package";
import type { XmlAdapter } from "../../core/ooxml/xmlAdapter";
import type { PptxShape } from "./model";
import { collectShapeParas } from "./model";

export const PRESENTATION_PART = "ppt/presentation.xml";
const SLIDES_DIR = "ppt/slides/";
const REL_TYPE_SLIDE = "officeDocument/2006/relationships/slide";
const REL_TYPE_LAYOUT = "officeDocument/2006/relationships/slideLayout";
const PKG_RELS_NS =
	"http://schemas.openxmlformats.org/package/2006/relationships";

/** Resolves a rels target ("slides/x.xml", "../y.xml") against a source part. */
export function resolvePartPath(sourcePart: string, target: string): string {
	const dir = sourcePart.includes("/")
		? sourcePart.slice(0, sourcePart.lastIndexOf("/"))
		: "";
	const segments = dir === "" ? [] : dir.split("/");
	for (const seg of target.split("/")) {
		if (seg === "..") {
			segments.pop();
		} else if (seg !== "." && seg !== "") {
			segments.push(seg);
		}
	}
	return segments.join("/");
}

/** Parses a package rels part into [id, type, target] triples. */
function readRelationships(
	pkg: OoxmlPackage,
	adapter: XmlAdapter,
	sourcePart: string,
): { id: string; type: string; target: string }[] {
	const slash = sourcePart.lastIndexOf("/");
	const dir = slash === -1 ? "" : sourcePart.slice(0, slash);
	const name = slash === -1 ? sourcePart : sourcePart.slice(slash + 1);
	const relsPath = `${dir}/_rels/${name}.rels`;
	if (!pkg.hasPart(relsPath)) {
		return [];
	}
	const bytes = pkg.getPartBytes(relsPath);
	const text = new TextDecoder().decode(bytes);
	const dom = adapter.parse(text);
	const rels = dom.getElementsByTagNameNS(PKG_RELS_NS, "Relationship");
	const out: { id: string; type: string; target: string }[] = [];
	for (let i = 0; i < rels.length; i++) {
		const el = rels[i];
		out.push({
			id: el.getAttribute("Id") ?? "",
			type: el.getAttribute("Type") ?? "",
			target: el.getAttribute("Target") ?? "",
		});
	}
	return out;
}

/** Slide part paths in presentation order (sldIdLst → rels), numeric-sort fallback. */
export function slidePartPaths(pkg: OoxmlPackage, adapter: XmlAdapter): string[] {
	const rels = readRelationships(pkg, adapter, PRESENTATION_PART);
	const targetById = new Map(
		rels.map((rel) => [rel.id, rel] as const),
	);
	const ordered: string[] = [];
	if (pkg.hasPart(PRESENTATION_PART)) {
		const text = new TextDecoder().decode(pkg.getPartBytes(PRESENTATION_PART));
		const dom = adapter.parse(text);
		const sldIds = dom.getElementsByTagNameNS(NS.p, "sldId");
		for (let i = 0; i < sldIds.length; i++) {
			const rId = sldIds[i].getAttributeNS(NS.r, "id") ?? "";
			const rel = targetById.get(rId);
			if (!rel || !rel.type.endsWith(REL_TYPE_SLIDE)) {
				continue;
			}
			const path = resolvePartPath(PRESENTATION_PART, rel.target);
			if (pkg.hasPart(path)) {
				ordered.push(path);
			}
		}
	}
	const referenced = new Set(ordered);
	const fallback = pkg.partPaths()
		.filter(
			(path) =>
				path.startsWith(SLIDES_DIR) &&
				path.endsWith(".xml") &&
				!referenced.has(path),
		)
		.sort(compareSlideNumber);
	return [...ordered, ...fallback];
}

/** Numeric filename order: slide2 before slide10. */
function compareSlideNumber(a: string, b: string): number {
	const num = (path: string) => Number(path.slice(SLIDES_DIR.length, -4));
	return num(a) - num(b);
}

export interface ShapeGeometry {
	x: number | null;
	y: number | null;
	cx: number | null;
	cy: number | null;
}

/** Placeholder key: type (default "") + idx (default "0"). */
export function placeholderKey(el: Element): string | null {
	const ph = el.getElementsByTagNameNS(NS.p, "ph")[0] as Element | undefined;
	if (!ph) {
		return null;
	}
	return `${ph.getAttribute("type") ?? ""}#${ph.getAttribute("idx") ?? "0"}`;
}

/** Builds a placeholder-key → geometry map from a slide layout document. */
export function layoutGeometryMap(
	layoutDoc: Document,
): Map<string, ShapeGeometry> {
	const map = new Map<string, ShapeGeometry>();
	const sps = layoutDoc.getElementsByTagNameNS(NS.p, "sp");
	for (let i = 0; i < sps.length; i++) {
		const key = placeholderKey(sps[i]);
		const geometry = shapeGeometry(sps[i]);
		if (key && geometry.x !== null) {
			map.set(key, geometry);
		}
	}
	return map;
}

/** Reads a:xfrm off/ext from a shape's spPr, or nulls when absent. */
export function shapeGeometry(el: Element): ShapeGeometry {
	const xfrm = el.getElementsByTagNameNS(NS.a, "xfrm")[0] as
		| Element
		| undefined;
	if (!xfrm) {
		return { x: null, y: null, cx: null, cy: null };
	}
	const off = xfrm.getElementsByTagNameNS(NS.a, "off")[0] as
		| Element
		| undefined;
	const ext = xfrm.getElementsByTagNameNS(NS.a, "ext")[0] as
		| Element
		| undefined;
	const read = (parent: Element | undefined, attr: string): number | null => {
		if (!parent) {
			return null;
		}
		const value = parent.getAttribute(attr);
		return value === null || value === "" ? null : Number(value);
	};
	return {
		x: read(off, "x"),
		y: read(off, "y"),
		cx: read(ext, "cx"),
		cy: read(ext, "cy"),
	};
}

/** Resolves a slide's layout geometry map via its rels (null when absent). */
export function slideLayoutGeometry(
	pkg: OoxmlPackage,
	adapter: XmlAdapter,
	slidePath: string,
): Map<string, ShapeGeometry> | null {
	const rels = readRelationships(pkg, adapter, slidePath);
	const layoutRel = rels.find((rel) => rel.type.endsWith(REL_TYPE_LAYOUT));
	if (!layoutRel) {
		return null;
	}
	const path = resolvePartPath(slidePath, layoutRel.target);
	if (!pkg.hasPart(path)) {
		return null;
	}
	const text = new TextDecoder().decode(pkg.getPartBytes(path));
	return layoutGeometryMap(adapter.parse(text));
}

/** Parses all p:sp text shapes of a slide in document order. */
export function parseSlideShapes(
	slideDoc: Document,
	layoutGeometry: Map<string, ShapeGeometry> | null,
): PptxShape[] {
	const out: PptxShape[] = [];
	const sps = slideDoc.getElementsByTagNameNS(NS.p, "sp");
	for (let i = 0; i < sps.length; i++) {
		const el = sps[i];
		const txBody = el.getElementsByTagNameNS(NS.p, "txBody")[0] as
			| Element
			| undefined;
		if (!txBody) {
			continue; // non-text shapes are not editable
		}
		const nameEl = el.getElementsByTagNameNS(NS.p, "cNvPr")[0] as
			| Element
			| undefined;
		const key = placeholderKey(el);
		const geometry = shapeGeometry(el);
		if (geometry.x === null && key && layoutGeometry) {
			const inherited = layoutGeometry.get(key);
			if (inherited) {
				geometry.x = inherited.x;
				geometry.y = inherited.y;
				geometry.cx = inherited.cx;
				geometry.cy = inherited.cy;
			}
		}
		const phEl = el.getElementsByTagNameNS(NS.p, "ph")[0] as
			| Element
			| undefined;
		out.push({
			el,
			txBody: txBody ?? null,
			name: nameEl?.getAttribute("name") ?? "",
			placeholder: phEl?.getAttribute("type") ?? null,
			x: geometry.x,
			y: geometry.y,
			cx: geometry.cx,
			cy: geometry.cy,
			paras: collectShapeParas(el),
		});
	}
	return out;
}
