/**
 * docx engine: loads word/document.xml into the paragraph model, edits are
 * text-range replacements over w:t runs, and serialization re-zips the
 * package with ONLY document.xml re-serialized — every other part stays
 * byte-identical, so formatting/styles/media survive untouched.
 */
import { toArrayBuffer } from "../../core/bufferUtils";
import { OoxmlPackage } from "../../core/ooxml/package";
import { XmlPart } from "../../core/ooxml/xmlPart";
import type { XmlAdapter } from "../../core/ooxml/xmlAdapter";
import {
	buildSegments,
	indexOfQuery,
	replaceInRange,
} from "../../core/textRun";
import type {
	DocParagraphInfo,
	FindReplaceOptions,
	TextDocEngine,
} from "../engine";
import { paragraphText, type DocxParagraph } from "./model";
import { parseDocumentBody } from "./parse";

const DOCUMENT_PART = "word/document.xml";

export class DocxEngine implements TextDocEngine {
	readonly kind = "docx" as const;

	private constructor(
		private readonly pkg: OoxmlPackage,
		private readonly part: XmlPart,
		private readonly paras: DocxParagraph[],
	) {}

	static async load(bytes: Uint8Array, adapter: XmlAdapter): Promise<DocxEngine> {
		const pkg = await OoxmlPackage.load(toArrayBuffer(bytes));
		const part = await XmlPart.openIn(pkg, DOCUMENT_PART, adapter);
		return new DocxEngine(pkg, part, parseDocumentBody(part.doc));
	}

	get paragraphs(): readonly DocParagraphInfo[] {
		return this.paras.map((p) => ({
			text: paragraphText(p.nodes),
			inTable: p.inTable,
		}));
	}

	getParagraphText(index: number): string {
		const para = this.paras[index];
		return para ? paragraphText(para.nodes) : "";
	}

	replaceRange(index: number, start: number, end: number, insertText: string): boolean {
		const para = this.paras[index];
		if (!para) {
			return false;
		}
		const segments = buildSegments(para.nodes);
		const before = paragraphText(para.nodes);
		replaceInRange(segments, start, end, insertText);
		const after = paragraphText(para.nodes);
		return after !== before;
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
		for (let i = 0; i < this.paras.length; i++) {
			let text = this.getParagraphText(i);
			let at = 0;
			while (at <= text.length) {
				const hit = indexOfQuery(text, query, at, options.caseSensitive);
				if (hit === -1) {
					break;
				}
				this.replaceRange(i, hit, hit + query.length, replacement);
				count++;
				text = this.getParagraphText(i);
				at = hit + replacement.length;
			}
		}
		return count;
	}

	countMatches(query: string, options: FindReplaceOptions): number {
		if (query === "") {
			return 0;
		}
		let count = 0;
		for (const para of this.paras) {
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
		return count;
	}

	async serialize(): Promise<Uint8Array> {
		this.part.saveTo(this.pkg, DOCUMENT_PART);
		return this.pkg.serialize();
	}
}
