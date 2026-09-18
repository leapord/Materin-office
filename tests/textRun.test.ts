import { describe, expect, it } from "vitest";
import {
	buildSegments,
	needsSpacePreserve,
	replaceInRange,
	segmentsTotalLength,
	type SegmentNode,
} from "../src/core/textRun";

/** Plain-object segment node; the algorithm only sees getText/setText. */
function node(initial: string): SegmentNode {
	let text = initial;
	return {
		el: {} as Element,
		getText: () => text,
		setText: (next: string) => {
			text = next;
		},
	};
}

describe("buildSegments / segmentsTotalLength", () => {
	it("computes contiguous offsets across nodes", () => {
		const segs = buildSegments([node("abc"), node("de")]);
		expect(segs.map((s) => [s.start, s.end])).toEqual([
			[0, 3],
			[3, 5],
		]);
		expect(segmentsTotalLength(segs)).toBe(5);
		expect(segmentsTotalLength([])).toBe(0);
	});
});

describe("replaceInRange", () => {
	it("replaces within a single segment", () => {
		const n = node("hello world");
		const segs = buildSegments([n]);
		replaceInRange(segs, 6, 11, "there");
		expect(n.getText()).toBe("hello there");
	});

	it("spans three segments: insert lands in first intersecting, suffix survives", () => {
		const a = node("Hel");
		const b = node("lo w");
		const c = node("orld");
		replaceInRange(buildSegments([a, b, c]), 3, 8, "X"); // replace "lo wo"
		expect(a.getText()).toBe("Hel"); // entirely before the range — untouched
		expect(b.getText()).toBe("X"); // first intersecting: "lo w" → "X"
		expect(c.getText()).toBe("rld"); // keeps its uncovered suffix
	});

	it("handles replacement longer and shorter than the range", () => {
		const n = node("abcdef");
		const segs = buildSegments([n]);
		replaceInRange(segs, 1, 3, "XY");
		expect(n.getText()).toBe("aXYdef"); // shorter
		replaceInRange(buildSegments([n]), 0, 1, "1234");
		expect(n.getText()).toBe("1234XYdef"); // longer
	});

	it("handles ranges touching boundaries 0 and length", () => {
		const n = node("abc");
		replaceInRange(buildSegments([n]), 0, 0, ">"); // insert at 0
		expect(n.getText()).toBe(">abc");
		replaceInRange(buildSegments([n]), 4, 4, "<"); // insert at end
		expect(n.getText()).toBe(">abc<");
		replaceInRange(buildSegments([n]), 0, 5, ""); // clear all
		expect(n.getText()).toBe("");
	});

	it("inserts at a boundary between segments into the following one", () => {
		const a = node("ab");
		const b = node("cd");
		replaceInRange(buildSegments([a, b]), 2, 2, "-");
		expect(a.getText()).toBe("ab");
		expect(b.getText()).toBe("-cd");
	});

	it("appends at the very end to the last segment", () => {
		const a = node("ab");
		const b = node("cd");
		replaceInRange(buildSegments([a, b]), 4, 4, "!");
		expect(a.getText()).toBe("ab");
		expect(b.getText()).toBe("cd!");
	});

	it("is a no-op when segments are empty or replacement is empty on insert", () => {
		expect(() => replaceInRange([], 0, 0, "x")).not.toThrow(); // caller must create runs
		const n = node("abc");
		replaceInRange(buildSegments([n]), 1, 1, "");
		expect(n.getText()).toBe("abc");
	});

	it("throws RangeError on out-of-bounds or reversed ranges", () => {
		const n = node("abc");
		expect(() => replaceInRange(buildSegments([n]), -1, 1, "x")).toThrow(
			RangeError,
		);
		expect(() => replaceInRange(buildSegments([n]), 2, 1, "x")).toThrow(
			RangeError,
		);
		expect(() => replaceInRange(buildSegments([n]), 0, 4, "x")).toThrow(
			RangeError,
		);
	});

	it("does not rewrite nodes whose text is unchanged", () => {
		const n = node("abcdef");
		let writes = 0;
		const wrapped: SegmentNode = {
			el: {} as Element,
			getText: () => n.getText(),
			setText: (t: string) => {
				writes++;
				n.setText(t);
			},
		};
		replaceInRange(buildSegments([wrapped]), 1, 2, "b"); // same text back
		expect(writes).toBe(0);
	});
});

describe("needsSpacePreserve", () => {
	it("flags leading/trailing whitespace only", () => {
		expect(needsSpacePreserve("plain")).toBe(false);
		expect(needsSpacePreserve(" lead")).toBe(true);
		expect(needsSpacePreserve("trail ")).toBe(true);
		expect(needsSpacePreserve("")).toBe(false);
		expect(needsSpacePreserve("a b")).toBe(false);
	});
});
