/**
 * PptxEngine tests: shape parsing, geometry inheritance, slide ordering,
 * text replacement with run preservation, and byte-preserving serialization.
 */
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { PptxEngine } from "../src/formats/pptx/pptxEngine";
import {
	buildPptx,
	pptxLayoutXml,
	pptxParagraph,
	pptxPresentationRels,
	pptxPresentationXml,
	pptxSlideRels,
	pptxSlideXml,
	pptxTextBox,
	testAdapter,
} from "./fixtures";

async function makeEngine(files: Record<string, string>): Promise<PptxEngine> {
	return PptxEngine.load(await buildPptx(files), testAdapter);
}

/** Minimal package: one slide holding one text shape with the given a:p XML. */
async function makeSingleShapeEngine(paragraphs: string): Promise<PptxEngine> {
	return makeEngine({
		"ppt/presentation.xml": pptxPresentationXml(["rId1"]),
		"ppt/_rels/presentation.xml.rels": pptxPresentationRels({
			rId1: "slides/slide1.xml",
		}),
		"ppt/slides/slide1.xml": pptxSlideXml(
			pptxTextBox({ name: "框", paragraphs }),
		),
	});
}

describe("PptxEngine.load", () => {
	it("loads shapes with geometry, placeholder and joined paragraph text", async () => {
		const engine = await makeEngine({
			"ppt/presentation.xml": pptxPresentationXml(["rId1", "rId2"]),
			"ppt/_rels/presentation.xml.rels": pptxPresentationRels({
				rId1: "slides/slide1.xml",
				rId2: "other/ignored.xml",
			}),
			"ppt/slides/slide1.xml": pptxSlideXml(
				pptxTextBox({
					name: "标题 1",
					phType: "title",
					paragraphs: pptxParagraph("第一页标题"),
				}) +
					pptxTextBox({
						name: "TextBox 3",
						x: 914400,
						y: 914400,
						cx: 6000000,
						cy: 1000000,
						paragraphs: pptxParagraph("行一") + pptxParagraph("行二"),
					}),
			),
			"ppt/slides/_rels/slide1.xml.rels": pptxSlideRels(
				"../slideLayouts/slideLayout1.xml",
			),
			"ppt/slideLayouts/slideLayout1.xml": pptxLayoutXml(
				`<a:xfrm><a:off x="838200" y="365125"/><a:ext cx="3000000" cy="1325563"/></a:xfrm>`,
			),
			"ppt/slides/slide2.xml": pptxSlideXml(
				`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="组合"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
					pptxTextBox({
						name: "组内框",
						x: 100,
						y: 100,
						cx: 200,
						cy: 50,
						paragraphs: pptxParagraph("组内文本"),
					}) +
					`</p:grpSp>`,
			),
			"ppt/theme/theme1.xml": `<theme/>`,
		});
		expect(engine.slides).toHaveLength(2);
		expect(engine.slides[0]).toHaveLength(2);
		const [title, box] = engine.slides[0];
		expect(title.name).toBe("标题 1");
		expect(title.placeholder).toBe("title");
		expect(title.x).toBe(838200);
		expect(title.cy).toBe(1325563);
		expect(box.x).toBe(914400);
		expect(box.cy).toBe(1000000);
		expect(box.text).toBe("行一\n行二");
		// slide2 is not referenced by rels — discovered via fallback listing;
		// p:sp inside p:grpSp is reached by the descendant walk
		expect(engine.slides[1]).toHaveLength(1);
		expect(engine.slides[1][0].text).toBe("组内文本");
	});

	it("orders slides by sldIdLst, not by filename", async () => {
		const engine = await makeEngine({
			"ppt/presentation.xml": pptxPresentationXml(["rId1", "rId2"]),
			"ppt/_rels/presentation.xml.rels": pptxPresentationRels({
				rId1: "slides/slide2.xml",
				rId2: "slides/slide1.xml",
			}),
			"ppt/slides/slide1.xml": pptxSlideXml(
				pptxTextBox({ name: "A", paragraphs: pptxParagraph("第一页") }),
			),
			"ppt/slides/slide2.xml": pptxSlideXml(
				pptxTextBox({ name: "B", paragraphs: pptxParagraph("第二页") }),
			),
		});
		expect(engine.slides[0][0].text).toBe("第二页");
		expect(engine.slides[1][0].text).toBe("第一页");
	});
});

describe("PptxEngine.replaceShapeText", () => {
	it("preserves run formatting of covered runs", async () => {
		const engine = await makeEngine({
			"ppt/presentation.xml": pptxPresentationXml(["rId1"]),
			"ppt/_rels/presentation.xml.rels": pptxPresentationRels({
				rId1: "slides/slide1.xml",
			}),
			"ppt/slides/slide1.xml": pptxSlideXml(
				pptxTextBox({
					name: "框",
					paragraphs:
						`<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Hello </a:t></a:r>` +
						`<a:r><a:t>world</a:t></a:r></a:p>` +
						pptxParagraph("第二段"),
				}),
			),
		});
		expect(engine.replaceShapeText(0, 0, "Hi there\n第二段")).toBe(true);
		expect(engine.getShapeText(0, 0)).toBe("Hi there\n第二段");
	});

	it("adds paragraphs for extra lines and removes dropped lines", async () => {
		const engine = await makeSingleShapeEngine(pptxParagraph("第一行"));
		expect(engine.replaceShapeText(0, 0, "一\n二\n三")).toBe(true);
		expect(engine.getShapeText(0, 0)).toBe("一\n二\n三");
		expect(engine.replaceShapeText(0, 0, "一")).toBe(true);
		expect(engine.getShapeText(0, 0)).toBe("一");
	});

	it("creates a run when the paragraph has none", async () => {
		const engine = await makeSingleShapeEngine(
			`<a:p><a:endParaRPr lang="zh-CN"/></a:p>`,
		);
		expect(engine.replaceShapeText(0, 0, "新文本")).toBe(true);
		expect(engine.getShapeText(0, 0)).toBe("新文本");
	});

	it("preserves leading and trailing whitespace in a:t", async () => {
		const engine = await makeSingleShapeEngine(pptxParagraph("x"));
		expect(engine.replaceShapeText(0, 0, "  空白保留  ")).toBe(true);
		expect(engine.getShapeText(0, 0)).toBe("  空白保留  ");
	});
});

describe("PptxEngine.findReplaceAll", () => {
	it("replaces across slides and returns the count", async () => {
		const engine = await makeEngine({
			"ppt/presentation.xml": pptxPresentationXml(["rId1", "rId2"]),
			"ppt/_rels/presentation.xml.rels": pptxPresentationRels({
				rId1: "slides/slide1.xml",
				rId2: "slides/slide2.xml",
			}),
			"ppt/slides/slide1.xml": pptxSlideXml(
				pptxTextBox({ name: "A", paragraphs: pptxParagraph("旧值A") }),
			),
			"ppt/slides/slide2.xml": pptxSlideXml(
				pptxTextBox({ name: "B", paragraphs: pptxParagraph("旧值B") }),
			),
		});
		const count = engine.findReplaceAll("旧", "新", { caseSensitive: false });
		expect(count).toBe(2);
		expect(engine.getShapeText(0, 0)).toBe("新值A");
		expect(engine.getShapeText(1, 0)).toBe("新值B");
	});
});

describe("PptxEngine.serialize", () => {
	it("keeps untouched parts byte-identical and writes the edited slide", async () => {
		const files: Record<string, string> = {
			"ppt/presentation.xml": pptxPresentationXml(["rId1"]),
			"ppt/_rels/presentation.xml.rels": pptxPresentationRels({
				rId1: "slides/slide1.xml",
			}),
			"ppt/slides/slide1.xml": pptxSlideXml(
				pptxTextBox({ name: "A", paragraphs: pptxParagraph("保存前") }),
			),
			"ppt/theme/theme1.xml": `<theme/>`,
		};
		const engine = await makeEngine(files);
		const before = await engine.serialize();
		const changed = engine.replaceShapeText(0, 0, "保存后");
		expect(changed).toBe(true);
		const out = await engine.serialize();
		const zip = await JSZip.loadAsync(out);
		const slideXml = await zip
			.file("ppt/slides/slide1.xml")!
			.async("string");
		expect(slideXml).toContain("保存后");
		expect(slideXml).toContain('standalone="yes"');
		const themeBefore = await JSZip.loadAsync(before);
		const themeAfter = await JSZip.loadAsync(out);
		expect(
			await themeAfter.file("ppt/theme/theme1.xml")!.async("uint8array"),
		).toEqual(
			await themeBefore.file("ppt/theme/theme1.xml")!.async("uint8array"),
		);
	});
});
