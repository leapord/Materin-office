import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { XlsEngine } from "../src/formats/xls/xlsEngine";

function buildXlsBytes(): Uint8Array {
	const wb = XLSX.utils.book_new();
	const ws = XLSX.utils.aoa_to_sheet([
		["姓名", "数量"],
		["苹果", 42],
		[true, 1.25],
	]);
	XLSX.utils.book_append_sheet(wb, ws, "数据");
	return new Uint8Array(XLSX.write(wb, { bookType: "xls", type: "array" }) as ArrayBuffer);
}

describe("XlsEngine", () => {
	it("loads cells with display text and kinds", () => {
		const engine = XlsEngine.load(buildXlsBytes());
		expect(engine.sheets).toHaveLength(1);
		expect(engine.sheets[0].name).toBe("数据");
		expect(engine.getCellDisplay(0, 1, 1).text).toBe("姓名");
		const num = engine.getCellDisplay(0, 2, 2);
		expect(num.kind).toBe("number");
		expect(num.text).toBe("42");
		const bool = engine.getCellDisplay(0, 3, 1);
		expect(bool.kind).toBe("bool");
		expect(engine.getCellDisplay(0, 99, 9).kind).toBe("empty");
	});

	it("applies typed input and round-trips through serialize", async () => {
		const engine = XlsEngine.load(buildXlsBytes());
		expect(engine.setCellValue(0, 1, 1, "水果")).toBe(true);
		expect(engine.setCellValue(0, 2, 2, "100")).toBe(true);
		expect(engine.setCellValue(0, 2, 2, "100")).toBe(false);
		expect(engine.setCellValue(0, 5, 3, "=B2*2")).toBe(true);
		const out = await engine.serialize();
		const reloaded = XlsEngine.load(out);
		expect(reloaded.getCellDisplay(0, 1, 1).text).toBe("水果");
		expect(reloaded.getCellDisplay(0, 2, 2).text).toBe("100");
		// formulas are stored as text in xls (documented limitation)
		expect(reloaded.getCellDisplay(0, 5, 3).text).toBe("=B2*2");
	});

	it("switches sheets via selectSheet", () => {
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["s1"]]), "一");
		XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["s2"]]), "二");
		const engine = XlsEngine.load(
			new Uint8Array(XLSX.write(wb, { bookType: "xls", type: "array" }) as ArrayBuffer),
		);
		expect(engine.sheets).toHaveLength(2);
		engine.selectSheet(1);
		expect(engine.activeSheet).toBe(1);
		expect(engine.getCellDisplay(1, 1, 1).text).toBe("s2");
	});
});
