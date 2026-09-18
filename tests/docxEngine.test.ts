import { describe, expect, it } from "vitest";
import { toArrayBuffer } from "../src/core/bufferUtils";
import { DocxEngine } from "../src/formats/docx/docxEngine";
import { buildDocx, docxDocumentXml, docxParagraph, testAdapter } from "./fixtures";
import JSZip from "jszip";

/** Body: two plain paragraphs, one mixed-run paragraph, a table cell. */
function sampleBody(): string {
	return (
		"<w:p><w:r><w:t>第一段</w:t></w:r></w:p>" +
		"<w:p><w:r><w:t>Hel</w:t></w:r><w:r><w:t>lo wo</w:t></w:r><w:r><w:t>rld</w:t></w:r></w:p>" +
		"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格单元格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
		"<w:p><w:r><w:t>尾段</w:t></w:r></w:p>"
	);
}

describe("docx 引擎", () => {
	it("解析段落并标记表格单元格", async () => {
		const engine = await DocxEngine.load(await buildDocx(sampleBody()), testAdapter);
		expect(engine.paragraphs.map((p) => p.text)).toEqual([
			"第一段",
			"Hello world",
			"表格单元格",
			"尾段",
		]);
		expect(engine.paragraphs[2].inTable).toBe(true);
		expect(engine.paragraphs[0].inTable).toBe(false);
	});

	it("跨多个 run 替换并保留段落文本", async () => {
		const engine = await DocxEngine.load(await buildDocx(sampleBody()), testAdapter);
		expect(engine.getParagraphText(1)).toBe("Hello world");
		expect(engine.replaceRange(1, 3, 8, "X")).toBe(true);
		expect(engine.getParagraphText(1)).toBe("HelXrld");
	});

	it("替换结果与原文相同时返回 false", async () => {
		const engine = await DocxEngine.load(await buildDocx(sampleBody()), testAdapter);
		expect(engine.replaceRange(1, 3, 8, "lo wo")).toBe(false);
		expect(engine.replaceRange(99, 0, 1, "字")).toBe(false);
	});

	it("越界索引返回 false", async () => {
		const engine = await DocxEngine.load(await buildDocx(sampleBody()), testAdapter);
		expect(engine.getParagraphText(99)).toBe("");
	});

	it("边缘空白写入触发 xml:space 保留", async () => {
		const engine = await DocxEngine.load(await buildDocx(sampleBody()), testAdapter);
		engine.replaceRange(0, 0, 3, " 带空白 ");
		const bytes = await engine.serialize();
		const zip = await JSZip.loadAsync(toArrayBuffer(bytes));
		const xml = await zip.file("word/document.xml")?.async("string");
		expect(xml).toContain('xml:space="preserve"');
	});

	it("未编辑的包条目往返保持内容不变", async () => {
		const zip = new JSZip();
		zip.file("[Content_Types].xml", "占位内容保持原样");
		zip.file("word/document.xml", docxDocumentXml(sampleBody()));
		const bytes = await zip.generateAsync({ type: "uint8array" });
		const engine = await DocxEngine.load(bytes, testAdapter);
		engine.replaceRange(0, 0, 3, "已改");
		const out = await engine.serialize();
		const outZip = await JSZip.loadAsync(toArrayBuffer(out));
		expect(await outZip.file("[Content_Types].xml")?.async("string")).toBe("占位内容保持原样");
	});

	it("findReplaceAll 替换全部命中并经序列化往返保持", async () => {
		const body = docxParagraph("Hello 甲") + docxParagraph("Hello 乙");
		const engine = await DocxEngine.load(await buildDocx(body), testAdapter);
		expect(engine.countMatches("Hello", { caseSensitive: true })).toBe(2);
		expect(engine.findReplaceAll("Hello", "你好", { caseSensitive: true })).toBe(2);
		const reloaded = await DocxEngine.load(await engine.serialize(), testAdapter);
		expect(reloaded.getParagraphText(0)).toBe("你好 甲");
		expect(reloaded.getParagraphText(1)).toBe("你好 乙");
	});

	it("大小写开关决定是否命中", async () => {
		const engine = await DocxEngine.load(await buildDocx(docxParagraph("Hello world")), testAdapter);
		expect(engine.countMatches("HELLO", { caseSensitive: true })).toBe(0);
		expect(engine.countMatches("HELLO", { caseSensitive: false })).toBe(1);
		expect(engine.findReplaceAll("HELLO", "你好", { caseSensitive: false })).toBe(1);
		expect(engine.getParagraphText(0)).toBe("你好 world");
	});

	it("同一段内多次命中全部替换", async () => {
		const engine = await DocxEngine.load(await buildDocx(docxParagraph("哈 哈 哈")), testAdapter);
		expect(engine.findReplaceAll("哈", "嘿", { caseSensitive: true })).toBe(3);
		expect(engine.getParagraphText(0)).toBe("嘿 嘿 嘿");
	});

	it("空查询返回零", async () => {
		const engine = await DocxEngine.load(await buildDocx(sampleBody()), testAdapter);
		expect(engine.countMatches("", { caseSensitive: true })).toBe(0);
		expect(engine.findReplaceAll("", "你好", { caseSensitive: true })).toBe(0);
	});
});
