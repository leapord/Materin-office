/**
 * Generates manual-test fixtures into the debug vault root:
 *   test-sample.docx — mixed-run CN/EN paragraphs + table + edge whitespace
 *   test-sample.xlsx — numbers / percent / date / formula / bold styling
 *   test-sample.pptx — two slides: title placeholder + multi-paragraph body
 *   test-sample.doc  — real binary Word via macOS textutil (skipped elsewhere)
 * Run: node scripts/makeFixtures.mjs
 */
import { writeFileSync, mkdirSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import ExcelJS from "exceljs";

const vaultRoot = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "debug-vault");
mkdirSync(vaultRoot, { recursive: true });

/** stdout writer (obsidianmd no-console applies repo-wide, even to scripts). */
function write(...parts) {
	process.stdout.write(`${parts.map(String).join(" ")}\n`);
}

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="word/styles.xml"/>
</Relationships>`;

const DOC_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
<w:style w:type="paragraph" w:styleId="Heading1">
<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
<w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
</w:style>
</w:styles>`;

function run(text, bold) {
	const rpr = bold ? "<w:rPr><w:b/></w:rPr>" : "";
	return `<w:r>${rpr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run("测试文档：Materin Office")} </w:p>
<w:p>${run("这是第一段，混合中英文与 ")}${run("加粗片段", true)}${run(" 以及普通文本。")}</w:p>
<w:p>${run("Hello ")}${run("wor", true)}${run("ld — English mixed-run paragraph.")}</w:p>
<w:tbl><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/></w:tblBorders>
<w:tr><w:tc><w:p>${run("表格单元格 甲")}</w:p></w:tc><w:tc><w:p>${run("表格单元格 乙")}</w:p></w:tc></w:tr>
</w:tbl>
<w:p>${run("  这一段两端带有空白，用于验证 xml:space 保留。  ")}</w:p>
<w:p>${run("尾段：把这里的文字替换成别的内容试试。")}</w:p>
</w:body></w:document>`;

async function makeDocx() {
	const zip = new JSZip();
	zip.file("[Content_Types].xml", CONTENT_TYPES);
	zip.file("_rels/.rels", ROOT_RELS);
	zip.file("word/document.xml", DOCUMENT);
	zip.file("word/styles.xml", DOC_STYLES);
	const bytes = await zip.generateAsync({ type: "nodebuffer" });
	writeFileSync(join(vaultRoot, "test-sample.docx"), bytes);
	write("test-sample.docx:", bytes.length, "bytes");
}

async function makeXlsx() {
	const wb = new ExcelJS.Workbook();
	const ws = wb.addWorksheet("数据");
	ws.columns = [
		{ header: "项目", key: "name", width: 24 },
		{ header: "数量", key: "qty", width: 10 },
		{ header: "单价", key: "price", width: 12 },
		{ header: "金额", key: "amount", width: 14 },
		{ header: "日期", key: "date", width: 14 },
	];
	const rows = [
		{ name: "键盘", qty: 3, price: 199, date: new Date(2026, 0, 15) },
		{ name: "显示器", qty: 1, price: 1299.5, date: new Date(2026, 1, 3) },
		{ name: "鼠标垫", qty: 5, price: 19.9, date: new Date(2026, 2, 20) },
	];
	for (const r of rows) {
		const row = ws.addRow(r);
		row.getCell("amount").value = { formula: `B${row.number}*C${row.number}` };
		row.getCell("qty").numFmt = "0";
		row.getCell("price").numFmt = "#,##0.00";
		row.getCell("amount").numFmt = "#,##0.00";
		row.getCell("date").numFmt = "yyyy-mm-dd";
	}
	ws.getRow(1).font = { bold: true };
	const total = ws.addRow({ name: "合计" });
	total.getCell("amount").value = { formula: "SUM(D2:D4)" };
	total.font = { bold: true };
	const pct = ws.addRow({ name: "完成度" });
	pct.getCell("amount").value = 0.618;
	pct.getCell("amount").numFmt = "0.0%";
	ws.getColumn("name").font = { name: "Microsoft YaHei" };
	wb.addWorksheet("说明").addRow(["第二个工作表：切换标签页试试。"]).font = { italic: true };
	const out = await wb.xlsx.writeBuffer();
	writeFileSync(join(vaultRoot, "test-sample.xlsx"), out);
	write("test-sample.xlsx:", out.length, "bytes");
}

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const R_NS =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const PPTX_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
</Types>`;

const PPTX_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

function pptxSp(name, x, y, cx, cy, paragraphs, ph = "") {
	return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
		`<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
		`<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

function pptxPara(text) {
	return `<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`;
}

async function makePptx() {
	const zip = new JSZip();
	const sp1 = pptxSp("标题 1", 838200, 365125, 10515600, 1325563,
		pptxPara("Materin Office 演示文稿"), '<p:ph type="title"/>');
	const sp2 = pptxSp("TextBox 4", 914400, 2057400, 10363200, 3200400,
		pptxPara("第一页正文：点击文本框即可编辑。") +
		pptxPara("第二段：混合 English 与中文。") +
		pptxPara("尾段：试一次查找替换。"));
	const sp3 = pptxSp("TextBox 2", 914400, 914400, 10363200, 3200400,
		pptxPara("第二页：slide2.xml 字节与 slide1 独立。") +
		pptxPara("编辑本页并保存后，slide1 应保持字节不变。"));
	zip.file("ppt/slides/slide1.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${sp1}${sp2}</p:spTree></p:cSld></p:sld>`);
	zip.file("ppt/slides/slide2.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${sp3}</p:spTree></p:cSld></p:sld>`);
	zip.file("ppt/presentation.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
	zip.file("ppt/_rels/presentation.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="${R_NS}/slide" Target="slides/slide2.xml"/></Relationships>`);
	zip.file("ppt/slides/_rels/slide1.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
	zip.file("ppt/slides/_rels/slide2.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
	zip.file("ppt/slideLayouts/slideLayout1.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`);
	zip.file("[Content_Types].xml", PPTX_CONTENT_TYPES);
	zip.file("_rels/.rels", PPTX_ROOT_RELS);
	const bytes = await zip.generateAsync({ type: "nodebuffer" });
	writeFileSync(join(vaultRoot, "test-sample.pptx"), bytes);
	write("test-sample.pptx:", bytes.length, "bytes");
}

function makeDoc() {
	// macOS textutil converts to real binary Word .doc (OLE); skipped elsewhere.
	try {
		const txt = join(vaultRoot, "test-doc-src.txt");
		writeFileSync(
			txt,
			"第一段：Materin Office 的 doc 只读提取样本。\n第二段：混合 English words 与中文。\n第三段：验证分段是否正确。",
		);
		execFileSync("textutil", [
			"-convert",
			"doc",
			txt,
			"-output",
			join(vaultRoot, "test-sample.doc"),
		]);
	} catch {
		write("textutil unavailable — skip test-sample.doc (non-macOS)");
		return;
	}
	write("test-sample.doc: generated via textutil");
}

await makeDocx();
await makeXlsx();
await makePptx();
makeDoc();
write("Fixtures written to", vaultRoot);
