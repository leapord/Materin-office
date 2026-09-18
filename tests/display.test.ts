import { describe, expect, it } from "vitest";
import {
	colName,
	colIndexFromName,
	toA1,
	parseA1,
	formatGeneralNumber,
	formatNumberWithFmt,
	formatDateTokens,
	isDateFmt,
	excelSerialToDate,
} from "../src/formats/xlsx/display";
import { parseUserInput } from "../src/formats/xlsx/model";

describe("A1 math", () => {
	it("maps column indices to names and back", () => {
		expect(colName(1)).toBe("A");
		expect(colName(26)).toBe("Z");
		expect(colName(27)).toBe("AA");
		expect(colName(52)).toBe("AZ");
		expect(colName(53)).toBe("BA");
		expect(colIndexFromName("A")).toBe(1);
		expect(colIndexFromName("AA")).toBe(27);
		expect(colIndexFromName("BA")).toBe(53);
	});
	it("round-trips addresses", () => {
		expect(toA1(3, 2)).toBe("B3");
		expect(parseA1("B3")).toEqual({ row: 3, col: 2 });
		expect(parseA1("bad")).toEqual({ row: 0, col: 0 });
	});
});

describe("number and date display", () => {
	it("renders general numbers with bounded precision", () => {
		expect(formatGeneralNumber(5)).toBe("5");
		expect(formatGeneralNumber(0.1).startsWith("0.1")).toBe(true);
	});
	it("applies thousands separators and decimals from the numFmt", () => {
		expect(formatNumberWithFmt(1234.5, "#,##0.00")).toBe("1,234.50");
		expect(formatNumberWithFmt(1234, "#,##0")).toBe("1,234");
		expect(formatNumberWithFmt(0.25, "0%")).toBe("25%");
	});
	it("formats date tokens including minutes-after-hours", () => {
		const d = new Date(2024, 0, 31, 14, 5, 9);
		expect(formatDateTokens(d, "yyyy-mm-dd")).toBe("2024-01-31");
		expect(formatDateTokens(d, "hh:mm")).toBe("14:05");
		expect(formatDateTokens(d, "h:m:s")).toBe("14:5:9");
	});
	it("detects date numFmts", () => {
		expect(isDateFmt("yyyy-mm-dd")).toBe(true);
		expect(isDateFmt("General")).toBe(false);
		expect(isDateFmt("#,##0.00")).toBe(false);
	});
	it("converts Excel serials to dates", () => {
		expect(excelSerialToDate(25569).toISOString().startsWith("1970-01-01")).toBe(true);
	});
});

describe("parseUserInput", () => {
	it("parses numbers, bools, formulas, ISO dates and plain text", () => {
		expect(parseUserInput("42")).toEqual({ t: "number", v: 42 });
		expect(parseUserInput("-3.5e2")).toEqual({ t: "number", v: -350 });
		expect(parseUserInput("true")).toEqual({ t: "bool", v: true });
		expect(parseUserInput("=SUM(A1:A2)")).toEqual({ t: "formula", f: "SUM(A1:A2)" });
		expect(parseUserInput("2024-01-31")).toEqual({
			t: "date",
			v: new Date(2024, 0, 31),
		});
		expect(parseUserInput("hello")).toEqual({ t: "text", v: "hello" });
		expect(parseUserInput("  spaced ")).toEqual({ t: "text", v: "  spaced " });
	});
	it("maps blank input to empty (cell clear)", () => {
		expect(parseUserInput("")).toEqual({ t: "empty" });
		expect(parseUserInput("   ")).toEqual({ t: "empty" });
	});
	it("keeps a lone = as text", () => {
		expect(parseUserInput("=")).toEqual({ t: "text", v: "=" });
	});
});
