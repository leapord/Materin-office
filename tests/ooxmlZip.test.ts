import { describe, expect, it } from "vitest";
import { decodeUtf8, encodeUtf8, toArrayBuffer } from "../src/core/bufferUtils";
import { CorruptFileError, MissingPartError } from "../src/core/errors";
import { OoxmlPackage } from "../src/core/ooxml/package";
import { buildZip } from "./fixtures";

function ab(bytes: Uint8Array): ArrayBuffer {
	return toArrayBuffer(bytes);
}

describe("OoxmlPackage", () => {
	it("preserves untouched parts byte-identically across a serialize round-trip", async () => {
		const original = await buildZip({
			"a.xml": `<?xml version="1.0"?><a/>`,
			"b/b.xml": `<?xml version="1.0" standalone="yes"?><b/>`,
		});
		const pkg = await OoxmlPackage.load(ab(original));

		expect(pkg.hasPart("a.xml")).toBe(true);
		expect(pkg.hasPart("missing.xml")).toBe(false);

		const before = decodeUtf8(pkg.getPartBytes("b/b.xml"));
		const rebuilt = await pkg.serialize();
		const reloaded = await OoxmlPackage.load(ab(rebuilt));
		expect(decodeUtf8(reloaded.getPartBytes("b/b.xml"))).toBe(before);
	});

	it("replaces edited parts while keeping others verbatim", async () => {
		const pkg = await OoxmlPackage.load(
			ab(
				await buildZip({
					"keep.xml": `<?xml version="1.0"?><keep/>`,
					"edit.xml": `<?xml version="1.0"?><edit>old</edit>`,
				}),
			),
		);

		pkg.setPartBytes(
			"edit.xml",
			encodeUtf8(`<?xml version="1.0"?><edit>new</edit>`),
		);

		const reloaded = await OoxmlPackage.load(ab(await pkg.serialize()));
		expect(decodeUtf8(reloaded.getPartBytes("edit.xml"))).toBe(
			`<?xml version="1.0"?><edit>new</edit>`,
		);
		expect(decodeUtf8(reloaded.getPartBytes("keep.xml"))).toBe(
			`<?xml version="1.0"?><keep/>`,
		);
	});

	it("throws MissingPartError for unknown parts", async () => {
		const pkg = await OoxmlPackage.load(
			ab(await buildZip({ "a.xml": `<?xml version="1.0"?><a/>` })),
		);
		expect(() => pkg.getPartBytes("nope.xml")).toThrow(MissingPartError);
		expect(() => pkg.setPartBytes("nope.xml", encodeUtf8("x"))).toThrow(
			MissingPartError,
		);
	});

	it("throws CorruptFileError for non-zip bytes", async () => {
		const garbage = new Uint8Array([0, 1, 2, 3, 4, 5]);
		await expect(
			OoxmlPackage.load(ab(garbage)),
		).rejects.toThrow(CorruptFileError);
	});

	it("tracks edited vs untouched parts", async () => {
		const pkg = await OoxmlPackage.load(
			ab(
				await buildZip({
					"a.xml": `<?xml version="1.0"?><a/>`,
					"b.xml": `<?xml version="1.0"?><b/>`,
				}),
			),
		);
		expect(pkg.editedPartPaths()).toEqual([]);
		pkg.setPartBytes("a.xml", encodeUtf8(`<?xml version="1.0"?><a edited="1"/>`));
		expect(pkg.editedPartPaths()).toEqual(["a.xml"]);
		expect(pkg.isEdited("a.xml")).toBe(true);
		expect(pkg.isEdited("b.xml")).toBe(false);
	});
});
