import { describe, expect, it } from "vitest";
import {
	detectKind,
	isEditableKind,
	OFFICE_EXTENSIONS,
} from "../src/core/detect";

describe("detectKind", () => {
	it("maps every claimed extension", () => {
		expect(detectKind("报告.docx")).toBe("docx");
		expect(detectKind("报表.xlsx")).toBe("xlsx");
		expect(detectKind("汇报.pptx")).toBe("pptx");
		expect(detectKind("旧报表.xls")).toBe("xls");
		expect(detectKind("旧报告.doc")).toBe("doc");
		expect(detectKind("旧汇报.ppt")).toBe("ppt");
	});

	it("handles directories and case-insensitively matches extensions", () => {
		expect(detectKind("方案/附件/报告.DOCX")).toBe("docx");
		expect(detectKind("notes/file.md")).toBeNull();
		expect(detectKind("no-extension")).toBeNull();
	});

	it("rejects dotfiles", () => {
		expect(detectKind(".docx")).toBeNull();
		expect(detectKind("notes/.gitignore")).toBeNull();
	});
});

describe("isEditableKind", () => {
	it("legacy doc/ppt are read-only, OOXML kinds are editable", () => {
		expect(isEditableKind("docx")).toBe(true);
		expect(isEditableKind("xlsx")).toBe(true);
		expect(isEditableKind("xls")).toBe(true);
		expect(isEditableKind("pptx")).toBe(true);
		expect(isEditableKind("doc")).toBe(false);
		expect(isEditableKind("ppt")).toBe(false);
	});
});

describe("OFFICE_EXTENSIONS", () => {
	it("claims exactly the six office extensions", () => {
		expect(OFFICE_EXTENSIONS).toHaveLength(6);
	});
});
