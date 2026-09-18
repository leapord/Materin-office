/**
 * Shared grid value types and the user-input parsing policy for the grid
 * engines (xlsx via ExcelJS, xls via SheetJS). obsidian-free.
 */

/** Parsed cell input, engine-agnostic. */
export type ParsedInput =
	| { t: "empty" }
	| { t: "number"; v: number }
	| { t: "bool"; v: boolean }
	| { t: "formula"; f: string }
	| { t: "date"; v: Date }
	| { t: "text"; v: string };

/** Typing "5" into a text cell keeps the sheet honest: numbers, bools, ISO
 * dates, and =formulas are recognized; anything else stays text. */
export function parseUserInput(raw: string): ParsedInput {
	const trimmed = raw.trim();
	if (trimmed === "") {
		return { t: "empty" };
	}
	if (trimmed.startsWith("=") && trimmed.length > 1) {
		return { t: "formula", f: trimmed.slice(1) };
	}
	const lower = trimmed.toLowerCase();
	if (lower === "true") {
		return { t: "bool", v: true };
	}
	if (lower === "false") {
		return { t: "bool", v: false };
	}
	if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(trimmed)) {
		return { t: "number", v: Number(trimmed) };
	}
	const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
	if (iso) {
		return { t: "date", v: new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])) };
	}
	return { t: "text", v: raw };
}
