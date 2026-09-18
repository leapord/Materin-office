import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { toArrayBuffer } from "../src/core/bufferUtils";
import { XlsxEngine, scanXlsxRiskParts } from "../src/formats/xlsx/xlsxEngine";

/** Builds a small workbook with values, a formula, dates and styles. */
async function buildWorkbookBytes(): Promise<Uint8Array> {
	const wb = new ExcelJS.Workbook();
	const ws = wb.addWorksheet("数据");
	ws.getCell("A1").value = "姓名";
	ws.getCell("B1").value = "数量";
	ws.getCell("A2").value = "苹果";
	ws.getCell("B2").value = 42;
	ws.getCell("B2").numFmt = "#,##0";
	ws.getCell("C2").value = 1.25;
	ws.getCell("C2").numFmt = "0%";
	ws.getCell("B3").value = { formula: "B2*2" };
	ws.getCell("A4").value = true;
	ws.getCell("B4").value = new Date(2024, 0, 31);
	ws.getColumn(1).width = 18;
	ws.getCell("A2").font = { bold: true };
	return new Uint8Array(await wb.xlsx.writeBuffer());
}

describe("XlsxEngine round-trip", () => {
	it("loads cells with numFmt-aware display and editor raw values", async () => {
		const engine = await XlsxEngine.load(await buildWorkbookBytes());
		expect(engine.sheets).toHaveLength(1);
		expect(engine.sheets[0].name).toBe("数据");
		const a1 = engine.getCellDisplay(0, 1, 1);
		expect(a1.text).toBe("姓名");
		expect(a1.kind).toBe("text");
		const b2 = engine.getCellDisplay(0, 2, 2);
		expect(b2.text).toBe("42");
		expect(b2.kind).toBe("number");
		expect(b2.align).toBe("right");
		const c2 = engine.getCellDisplay(0, 2, 3);
		expect(c2.text).toBe("125%");
		const b3 = engine.getCellDisplay(0, 3, 2);
		expect(b3.kind).toBe("formula");
		expect(b3.raw).toBe("=B2*2");
		const a4 = engine.getCellDisplay(0, 4, 1);
		expect(a4.text).toBe("TRUE");
		const empty = engine.getCellDisplay(0, 50, 10);
		expect(empty.kind).toBe("empty");
		expect(empty.text).toBe("");
	});

	it("applies typed input by inferred type and reports changed cells", async () => {
		const engine = await XlsxEngine.load(await buildWorkbookBytes());
		expect(engine.setCellValue(0, 1, 1, "水果")).toBe(true);
		expect(engine.setCellValue(0, 2, 2, "100")).toBe(true);
		expect(engine.getCellDisplay(0, 2, 2).kind).toBe("number");
		expect(engine.setCellValue(0, 3, 3, "true")).toBe(true);
		expect(engine.getCellDisplay(0, 3, 3).kind).toBe("bool");
		expect(engine.setCellValue(0, 4, 4, "=SUM(B2:C2)")).toBe(true);
		expect(engine.getCellDisplay(0, 4, 4).raw).toBe("=SUM(B2:C2)");
		expect(engine.setCellValue(0, 1, 1, "水果")).toBe(false);
		expect(engine.setCellValue(0, 90, 90, "")).toBe(false);
		expect(engine.setCellValue(0, 5, 5, "2024-01-31")).toBe(true);
		expect(engine.getCellDisplay(0, 5, 5).kind).toBe("date");
	});

	it("keeps styles and column widths through an edit round-trip", async () => {
		const engine = await XlsxEngine.load(await buildWorkbookBytes());
		engine.setCellValue(0, 1, 2, "数量（修改）");
		const out = await engine.serialize();
		const reloaded = await XlsxEngine.load(out);
		expect(reloaded.getCellDisplay(0, 1, 2).text).toBe("数量（修改）");

		const verify = new ExcelJS.Workbook();
		await verify.xlsx.load(toArrayBuffer(out));
		const ws = verify.getWorksheet("数据");
		expect(ws?.getCell("A2").font?.bold).toBe(true);
		expect(ws?.getColumn(1).width).toBe(18);
		expect(ws?.getCell("B2").value).toBe(42);
		expect(ws?.getCell("B3").value).toEqual({ formula: "B2*2" });
	});

	it("flags risk parts when charts or media are present", async () => {
		const wb = new ExcelJS.Workbook();
		wb.addWorksheet("Sheet1");
		const clean = new Uint8Array(await wb.xlsx.writeBuffer());
		expect(await scanXlsxRiskParts(clean)).toEqual([]);
		const fake = await buildZipWithRisk();
		const hits = await scanXlsxRiskParts(fake);
		expect(hits).toContain("xl/media");
	});
});

import { buildZip } from "./fixtures";

async function buildZipWithRisk(): Promise<Uint8Array> {
	return buildZip({
		"[Content_Types].xml": "<Types/>",
		"xl/media/image1.png": "fake",
		"worksheets/sheet1.xml": "<sheet/>",
	});
}
