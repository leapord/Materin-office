import { decodeUtf8, encodeUtf8 } from "../bufferUtils";
import { MissingPartError } from "../errors";
import type { XmlAdapter } from "./xmlAdapter";
import type { OoxmlPackage } from "./package";

/** U+FEFF, spelled out so the source never contains an invisible literal. */
const BOM = String.fromCharCode(0xfeff);

/**
 * Matches an XML declaration (plus any leading BOM/whitespace), used both to
 * capture it on open and to strip one a serializer may have re-added.
 */
const XML_DECL_RE = new RegExp(`^${BOM}?\\s*(<\\?xml[\\s\\S]*?\\?>)`);

/**
 * A mutable XML part of an OOXML package.
 *
 * XMLSerializer omits the XML declaration (and a leading BOM never survives a
 * string round-trip), so we capture both on open and re-prepend them verbatim
 * on save — that is what keeps `standalone="yes"` and encoding intact.
 */
export class XmlPart {
	readonly doc: Document;
	private readonly prefix: string;
	private readonly adapter: XmlAdapter;

	private constructor(doc: Document, prefix: string, adapter: XmlAdapter) {
		this.doc = doc;
		this.prefix = prefix;
		this.adapter = adapter;
	}

	static open(partText: string, adapter: XmlAdapter): XmlPart {
		const bom = partText.startsWith(BOM) ? BOM : "";
		const withoutBom = bom ? partText.slice(1) : partText;
		const match = XML_DECL_RE.exec(withoutBom);
		const decl = match ? match[1] : "";
		const body = match ? withoutBom.slice(match[0].length) : withoutBom;
		const doc = adapter.parse(body);
		return new XmlPart(doc, bom + decl, adapter);
	}

	static async openIn(
		pkg: OoxmlPackage,
		path: string,
		adapter: XmlAdapter,
	): Promise<XmlPart> {
		if (!pkg.hasPart(path)) {
			throw new MissingPartError(path);
		}
		return XmlPart.open(decodeUtf8(pkg.getPartBytes(path)), adapter);
	}

	/** Serializes the DOM and writes the result back into the package. */
	saveTo(pkg: OoxmlPackage, path: string): void {
		let out = this.adapter.serializeToString(this.doc);
		// Defensive: some serializers re-add a declaration — ours must win.
		out = out.replace(XML_DECL_RE, "");
		pkg.setPartBytes(path, encodeUtf8(this.prefix + out));
	}
}
