import { describe, expect, it } from "vitest";
import { decodeUtf8, toArrayBuffer } from "../src/core/bufferUtils";
import { MissingPartError } from "../src/core/errors";
import { OoxmlPackage } from "../src/core/ooxml/package";
import { XmlPart } from "../src/core/ooxml/xmlPart";
import { buildDocx, NS_W as NS_W_NS, testAdapter } from "./fixtures";

function ab(bytes: Uint8Array): ArrayBuffer {
	return toArrayBuffer(bytes);
}

const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

describe("XmlPart declaration preservation", () => {
	it("keeps declaration with standalone and no double declaration after save", async () => {
		const bytes = await buildDocx(`<w:p/>`);
		const pkg = await OoxmlPackage.load(ab(bytes));
		const part = await XmlPart.openIn(pkg, "word/document.xml", testAdapter);

		part.saveTo(pkg, "word/document.xml");
		const out = decodeUtf8(pkg.getPartBytes("word/document.xml"));

		expect(out.startsWith(DECL)).toBe(true);
		expect(out.indexOf("<?xml")).toBe(out.lastIndexOf("<?xml"));
	});

	it("keeps a leading UTF-8 BOM across open/save", async () => {
		const bytes = await buildDocx(`<w:p/>`);
		const pkg = await OoxmlPackage.load(ab(bytes));
		// Synthesize a part that originally had a BOM.
		const raw = decodeUtf8(pkg.getPartBytes("word/document.xml"));
		const withBom = `${String.fromCharCode(0xfeff)}${raw}`;
		pkg.setPartBytes(
			"word/document.xml",
			new TextEncoder().encode(withBom),
		);

		const part = await XmlPart.openIn(pkg, "word/document.xml", testAdapter);
		part.saveTo(pkg, "word/document.xml");
		const out = decodeUtf8(pkg.getPartBytes("word/document.xml"));

		expect(out.charCodeAt(0)).toBe(0xfeff);
		expect(out.slice(1).startsWith(DECL)).toBe(true);
	});

	it("preserves namespaces through parse-mutate-serialize", async () => {
		const body = `<w:p><w:r><w:t>before</w:t></w:r></w:p>`;
		const bytes = await buildDocx(body);
		const pkg = await OoxmlPackage.load(ab(bytes));
		const part = await XmlPart.openIn(pkg, "word/document.xml", testAdapter);

		const t = part.doc.getElementsByTagNameNS(NS_W_NS, "t")[0];
		expect(t).toBeDefined();
		t.textContent = "已改";
		part.saveTo(pkg, "word/document.xml");

		const pkg2 = await OoxmlPackage.load(ab(await pkg.serialize()));
		const part2 = await XmlPart.openIn(pkg2, "word/document.xml", testAdapter);
		const t2 = part2.doc.getElementsByTagNameNS(NS_W_NS, "t")[0];
		expect(t2.textContent).toBe("已改");
		// namespace still resolves — element found via NS lookup again
		expect(t2.namespaceURI).toBe(NS_W_NS);
	});
});

describe("XmlPart error paths", () => {
	it("throws on missing part", async () => {
		const bytes = await buildDocx(`<w:p/>`);
		const pkg = await OoxmlPackage.load(ab(bytes));
		await expect(
			XmlPart.openIn(pkg, "word/absent.xml", testAdapter),
		).rejects.toThrow(MissingPartError);
	});
});
