/**
 * View-layer integration tests under jsdom: renders the pptx canvas and the
 * doc read-only view with real engine instances and real fixture bytes, and
 * exercises the declarative settings tab — no Obsidian app required.
 */
// @vitest-environment jsdom
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { PptxEngine } from "../src/formats/pptx/pptxEngine";
import { DocTextEngine } from "../src/formats/doc/docTextEngine";
import { PptxCanvas } from "../src/view/pptx/PptxCanvas";
import { DocViewUi } from "../src/view/doc/DocViewUi";
import { DEFAULT_SETTINGS, OfficeSettingTab } from "../src/settings";
import { testAdapter } from "./fixtures";
import { createHost, installObsidianDomPolyfill } from "./uiDom";

const PPTX = join(__dirname, "..", "..", "debug-vault", "test-sample.pptx");
const DOC = join(__dirname, "..", "..", "debug-vault", "test-sample.doc");
const itIf = (cond: boolean) => (cond ? it : it.skip);

describe("PptxCanvas (jsdom)", () => {
	itIf(existsSync(PPTX))("renders, edits and commits via blur", async () => {
		const engine = await PptxEngine.load(
			new Uint8Array(readFileSync(PPTX)),
			testAdapter,
		);
		const host = createHost();
		let mutated = 0;
		const canvas = new PptxCanvas(engine, {
			onMutated: () => {
				mutated += 1;
			},
			status: () => {},
		});
		canvas.render(host);
		const slide = host.querySelector(".materin-office-pptx-slide");
		expect(slide).not.toBeNull();
		const shapes = Array.from(
			host.querySelectorAll(".materin-office-pptx-shape"),
		);
		expect(shapes).toHaveLength(2);
		const [title, box] = shapes as HTMLElement[];
		expect(title.style.left).toMatch(/%$/);
		expect(title.style.top).toMatch(/%$/);
		expect(title.style.width).toMatch(/%$/
		);
		expect(title.getAttribute("data-shape-index")).toBe("0");
		expect(box.textContent).toContain("第一页正文");
		// click-to-edit: textarea appears prefilled, blur commits into the engine
		title.click();
		const area = title.querySelector("textarea") as HTMLTextAreaElement;
		expect(area).not.toBeNull();
		expect(area.value).toBe(engine.getShapeText(0, 0));
		area.value = "改过的标题";
		area.dispatchEvent(new Event("blur"));
		expect(engine.getShapeText(0, 0)).toBe("改过的标题");
		expect(mutated).toBe(1);
		// second shape untouched
		expect(engine.getShapeText(0, 1)).toContain("第一页正文");
	});
});

describe("DocViewUi (jsdom)", () => {
	itIf(existsSync(DOC))("renders extracted paragraphs read-only", async () => {
		const engine = await DocTextEngine.load(new Uint8Array(readFileSync(DOC)));
		const host = createHost();
		const openExternal = vi.fn();
		const ui = new DocViewUi(engine, {
			setDirty: () => {},
			settings: () => {
				throw new Error("unused");
			},
			status: () => {},
			openExternal,
		});
		await ui.render(host);
		const rows = host.querySelectorAll(".materin-office-para");
		expect(rows.length).toBeGreaterThanOrEqual(3);
		const all = host.textContent ?? "";
		expect(all).toContain("第一段");
		expect(all).toContain("第二段");
		expect(all).toContain("第三段");
		const button = host.querySelector("button.mod-muted");
		expect(button?.textContent).toBe("用系统程序打开");
		(button as HTMLElement).click();
		expect(openExternal).toHaveBeenCalledTimes(1);
	});
});

describe("OfficeSettingTab (declarative)", () => {
	it("exposes five definitions and persists via setControlValue", () => {
		installObsidianDomPolyfill(window);
		const plugin = {
			settings: { ...DEFAULT_SETTINGS },
			saveSettings: vi.fn(async () => {}),
			notifyViews: vi.fn(),
		};
		const tab = new OfficeSettingTab({} as never, plugin as never);
		const defs = tab.getSettingDefinitions();
		expect(defs).toHaveLength(5);
		const controls = defs.map(
			(def) => (def as { control?: { key: string } }).control?.key,
		);
		expect(controls).toEqual([
			"backupEnabled",
			"backupKeepPerFile",
			"docxEditMode",
			"showXlsxGridLines",
			"showXlsxFormulas",
		]);
		tab.setControlValue("showXlsxGridLines", false);
		expect(plugin.settings.showXlsxGridLines).toBe(false);
		expect(plugin.saveSettings).toHaveBeenCalledTimes(1);
		expect(plugin.notifyViews).toHaveBeenCalledTimes(1);
		expect(tab.getControlValue("showXlsxGridLines")).toBe(false);
	});
});
