/**
 * Injectable XML parser/serializer so the OOXML layer runs both in the
 * Electron renderer (native DOMParser/XMLSerializer) and under vitest
 * (@xmldom/xmldom).
 */
export interface XmlAdapter {
	parse(text: string): Document;
	serializeToString(doc: Document): string;
}

/** Adapter over the renderer's native DOMParser/XMLSerializer. */
export function createDomXmlAdapter(): XmlAdapter {
	return {
		parse(text: string): Document {
			const doc = new DOMParser().parseFromString(text, "application/xml");
			if (doc.getElementsByTagName("parsererror").length > 0) {
				throw new Error("XML 解析失败");
			}
			return doc;
		},
		serializeToString(doc: Document): string {
			return new XMLSerializer().serializeToString(doc);
		},
	};
}
