/**
 * pptx engine: loads each slide part into the shape model; edits replace
 * shape text with per-run formatting preservation (shared textRun algorithm),
 * and serialization re-serializes ONLY edited slide parts — every other part
 * stays byte-identical, so layouts, media and theme survive untouched.
 */
import { toArrayBuffer } from "../../core/bufferUtils";
import { NS } from "../../core/ooxml/ns";
import { OoxmlPackage } from "../../core/ooxml/package";
import { XmlPart } from "../../core/ooxml/xmlPart";
import type { XmlAdapter } from "../../core/ooxml/xmlAdapter";
import {
	buildSegments,
	indexOfQuery,
	replaceInRange,
} from "../../core/textRun";
import type {
	FindReplaceOptions,
	SlideCanvasEngine,
	SlideShapeInfo,
	SlideSize,
} from "../engine";
import {
	parseSlideShapes,
	slideLayoutGeometry,
	slidePartPaths,
	PRESENTATION_PART,
} from "./parse";
import {
	paragraphText,
	refreshShapeParas,
	type PptxParagraph,
	type PptxShape,
} from "./model";

interface SlideEntry {
	readonly path: string;
	readonly part: XmlPart;
	readonly shapes: PptxShape[];
}

export class PptxEngine implements SlideCanvasEngine {
	readonly kind = "pptx" as const;
	readonly slideSize: SlideSize;

	private readonly pkg: OoxmlPackage;
	private readonly adapter: XmlAdapter;
	private readonly entries: SlideEntry[];
	private readonly dirtyParts = new Set<string>();

	private constructor(
		pkg: OoxmlPackage,
		adapter: XmlAdapter,
		entries: SlideEntry[],
		slideSize: SlideSize,
	) {
		this.pkg = pkg;
		this.adapter = adapter;
		this.entries = entries;
		this.slideSize = slideSize;
	}

	static async load(
		bytes: Uint8Array,
		adapter: XmlAdapter,
	): Promise<PptxEngine> {
		const pkg = await OoxmlPackage.load(toArrayBuffer(bytes));
		const entries: SlideEntry[] = [];
		for (const path of slidePartPaths(pkg, adapter)) {
			const part = await XmlPart.openIn(pkg, path, adapter);
			const layoutGeometry = slideLayoutGeometry(pkg, adapter, path);
			const shapes = parseSlideShapes(part.doc, layoutGeometry);
			entries.push({ path, part, shapes });
		}
		return new PptxEngine(pkg, adapter, entries, readSlideSize(pkg, adapter));
	}

	get slides(): readonly SlideShapeInfo[][] {
		return this.entries.map((entry) =>
			entry.shapes.map((shape) => ({
				name: shape.name,
				placeholder: shape.placeholder,
				x: shape.x,
				y: shape.y,
				cx: shape.cx,
				cy: shape.cy,
				text: shape.paras.map((para) => paragraphText(para.nodes)).join("\n"),
			})),
		);
	}

	getShapeText(slide: number, shape: number): string {
		const target = this.entries[slide]?.shapes[shape];
		return target
			? target.paras.map((para) => paragraphText(para.nodes)).join("\n")
			: "";
	}

	replaceShapeText(slide: number, shape: number, text: string): boolean {
		const entry = this.entries[slide];
		const target = entry?.shapes[shape];
		if (!target || !target.txBody) {
			return false;
		}
		const lines = text.split("\n");
		const paras = target.paras;
		let changed = false;
		const overlap = Math.min(lines.length, paras.length);
		for (let i = 0; i < overlap; i++) {
			if (replaceParagraphText(paras[i], lines[i])) {
				changed = true;
			}
		}
		for (let i = paras.length; i < lines.length; i++) {
			const sourceEl = paras.length > 0 ? paras[paras.length - 1].el : undefined;
			appendClonedParagraph(target.txBody, sourceEl, lines[i]);
			changed = true;
		}
		if (paras.length > lines.length) {
			for (let i = paras.length - 1; i >= lines.length; i--) {
				paras[i].el.parentNode?.removeChild(paras[i].el);
				changed = true;
			}
		}
		if (changed) {
			refreshShapeParas(target);
			this.dirtyParts.add(entry.path);
		}
		return changed;
	}

	findReplaceAll(
		query: string,
		replacement: string,
		options: FindReplaceOptions,
	): number {
		if (query === "") {
			return 0;
		}
		let count = 0;
		for (let s = 0; s < this.entries.length; s++) {
			let slideDirty = false;
			for (const shape of this.entries[s].shapes) {
				for (const para of shape.paras) {
					let text = paragraphText(para.nodes);
					let at = 0;
					while (at <= text.length) {
						const hit = indexOfQuery(text, query, at, options.caseSensitive);
						if (hit === -1) {
							break;
						}
						const next =
							text.slice(0, hit) + replacement + text.slice(hit + query.length);
						if (replaceParagraphText(para, next)) {
							count++;
							slideDirty = true;
						}
						text = paragraphText(para.nodes);
						at = hit + replacement.length;
					}
				}
			}
			if (slideDirty) {
				this.dirtyParts.add(this.entries[s].path);
			}
		}
		return count;
	}

	countMatches(query: string, options: FindReplaceOptions): number {
		if (query === "") {
			return 0;
		}
		let count = 0;
		for (const entry of this.entries) {
			for (const shape of entry.shapes) {
				for (const para of shape.paras) {
					const text = paragraphText(para.nodes);
					let at = 0;
					while (at <= text.length) {
						const hit = indexOfQuery(text, query, at, options.caseSensitive);
						if (hit === -1) {
							break;
						}
						count++;
						at = hit + query.length;
					}
				}
			}
		}
		return count;
	}

	async serialize(): Promise<Uint8Array> {
		for (const entry of this.entries) {
			if (this.dirtyParts.has(entry.path)) {
				entry.part.saveTo(this.pkg, entry.path);
			}
		}
		return this.pkg.serialize();
	}
}

const DEFAULT_SLIDE_SIZE: SlideSize = { widthEmu: 9144000, heightEmu: 6858000 };

function readSlideSize(pkg: OoxmlPackage, adapter: XmlAdapter): SlideSize {
	if (!pkg.hasPart(PRESENTATION_PART)) {
		return { ...DEFAULT_SLIDE_SIZE };
	}
	const text = new TextDecoder().decode(pkg.getPartBytes(PRESENTATION_PART));
	const dom = adapter.parse(text);
	const sldSz = dom.getElementsByTagNameNS(NS.p, "sldSz")[0] as
		| Element
		| undefined;
	if (!sldSz) {
		return { ...DEFAULT_SLIDE_SIZE };
	}
	return {
		widthEmu: Number(sldSz.getAttribute("cx")) || DEFAULT_SLIDE_SIZE.widthEmu,
		heightEmu:
			Number(sldSz.getAttribute("cy")) || DEFAULT_SLIDE_SIZE.heightEmu,
	};
}

/** Replaces a whole paragraph's text; creates a run when the paragraph has none. */
function replaceParagraphText(para: PptxParagraph, next: string): boolean {
	const before = para.nodes.map((node) => node.getText()).join("");
	if (before === next) {
		return false;
	}
	if (para.nodes.length === 0) {
		if (next !== "") {
			insertRun(para.el, next);
		}
	} else {
		replaceInRange(buildSegments(para.nodes), 0, before.length, next);
	}
	return true;
}

function insertRun(pEl: Element, text: string): void {
	const doc = pEl.ownerDocument;
	if (!doc) {
		return;
	}
	const run = doc.createElementNS(NS.a, "r");
	const t = doc.createElementNS(NS.a, "t");
	t.textContent = text;
	run.appendChild(t);
	const endPara = pEl.getElementsByTagNameNS(NS.a, "endParaRPr")[0];
	pEl.insertBefore(run, endPara ?? null);
}

/**
 * Appends a new a:p holding one run with `text`, cloned from the last
 * paragraph so its list-level and run formatting carry over.
 */
function appendClonedParagraph(
	txBody: Element,
	sourceEl: Element | undefined,
	text: string,
): void {
	const doc = txBody.ownerDocument;
	if (!doc) {
		return;
	}
	let pEl: Element;
	if (sourceEl) {
		pEl = sourceEl.cloneNode(true) as Element;
	} else {
		pEl = doc.createElementNS(NS.a, "p");
	}
	const runs = pEl.getElementsByTagNameNS(NS.a, "r");
	if (runs.length === 0) {
		insertRun(pEl, text);
	} else {
		const first = runs[0];
		let t = first.getElementsByTagNameNS(NS.a, "t")[0] as
			| Element
			| undefined;
		if (!t) {
			t = doc.createElementNS(NS.a, "t");
			first.appendChild(t);
		}
		t.textContent = text;
		for (let i = runs.length - 1; i >= 1; i--) {
			runs[i].parentNode?.removeChild(runs[i]);
		}
	}
	// a:fld / a:br leftovers would inject stale field text or line breaks.
	const extras = [
		...Array.from(pEl.getElementsByTagNameNS(NS.a, "fld")),
		...Array.from(pEl.getElementsByTagNameNS(NS.a, "br")),
	];
	for (const extra of extras) {
		extra.parentNode?.removeChild(extra);
	}
	txBody.appendChild(pEl);
}
