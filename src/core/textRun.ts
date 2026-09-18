/**
 * Multi-run text replacement, shared by docx (w:t nodes) and pptx (a:t nodes).
 *
 * A paragraph's visible text is the concatenation of its text-bearing
 * elements. Editing an arbitrary [start, end) range must preserve every run's
 * formatting: the replacement lands in the FIRST intersecting element, fully
 * covered middle elements are emptied, and the LAST element keeps its
 * uncovered suffix. Elements never visited keep their XML untouched.
 */

/** A text-bearing element participating in a paragraph. */
export interface SegmentNode {
	readonly el: Element;
	getText(): string;
	setText(text: string): void;
}

export interface Segment extends SegmentNode {
	readonly text: string;
	readonly start: number;
	readonly end: number;
}

export function buildSegments(nodes: readonly SegmentNode[]): Segment[] {
	let offset = 0;
	return nodes.map((node) => {
		const text = node.getText();
		const segment: Segment = {
			el: node.el,
			getText: () => node.getText(),
			setText: (next: string) => node.setText(next),
			text,
			start: offset,
			end: offset + text.length,
		};
		offset += text.length;
		return segment;
	});
}

export function segmentsTotalLength(segments: readonly Segment[]): number {
	const last = segments[segments.length - 1];
	return last ? last.end : 0;
}

/** True when text starts or ends with whitespace — callers must set xml:space="preserve". */
export function needsSpacePreserve(text: string): boolean {
	return text.length > 0 && (/^\s/.test(text) || /\s$/.test(text));
}

/** Case-aware indexOf shared by the docx and pptx engines. */
export function indexOfQuery(
	haystack: string,
	needle: string,
	from: number,
	caseSensitive: boolean,
): number {
	return caseSensitive
		? haystack.indexOf(needle, from)
		: haystack.toLowerCase().indexOf(needle.toLowerCase(), from);
}

/**
 * Replaces [start, end) with insertText across the segment list.
 *
 * Insertion at a segment boundary lands at the start of the following
 * segment (or is appended to the last segment when the position is the very
 * end). Nodes are only written when their text actually changes.
 */
export function replaceInRange(
	segments: readonly Segment[],
	start: number,
	end: number,
	insertText: string,
): void {
	const total = segmentsTotalLength(segments);
	if (
		!Number.isInteger(start) ||
		!Number.isInteger(end) ||
		start < 0 ||
		end < start ||
		end > total
	) {
		throw new RangeError(
			`replaceInRange: invalid range [${start}, ${end}) over ${total} chars`,
		);
	}
	if (start === end && insertText.length === 0) {
		return;
	}

	let first = -1;
	let last = -1;
	for (let i = 0; i < segments.length; i++) {
		const seg = segments[i];
		if (seg.end > start && seg.start < end) {
			if (first === -1) {
				first = i;
			}
			last = i;
		}
	}

	if (first === -1) {
		// Pure insertion at a segment boundary (or into an empty document).
		if (insertText.length === 0) {
			return;
		}
		const target =
			segments.find((seg) => seg.start === start) ??
			segments[segments.length - 1];
		if (!target) {
			return; // No runs at all — the caller must create one first.
		}
		const at = start - target.start;
		target.setText(target.text.slice(0, at) + insertText + target.text.slice(at));
		return;
	}

	for (let i = first; i <= last; i++) {
		const seg = segments[i];
		const localStart = Math.max(start - seg.start, 0);
		const localEnd = Math.min(end - seg.start, seg.text.length);
		let text: string;
		if (i === first) {
			text =
				seg.text.slice(0, localStart) +
				insertText +
				(i === last ? seg.text.slice(localEnd) : "");
		} else if (i === last) {
			text = seg.text.slice(localEnd);
		} else {
			text = "";
		}
		if (text !== seg.text) {
			seg.setText(text);
		}
	}
}
