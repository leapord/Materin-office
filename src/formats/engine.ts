/**
 * Engine contracts shared by the format engines and the kind UIs.
 * obsidian-free: engines are constructed from raw bytes and serialize back.
 */

/** One editable worksheet as the grid UI sees it. */
export interface SheetInfo {
	readonly name: string;
	/** Scrollable row count (used range + padding). */
	readonly rowCount: number;
	/** Scrollable column count (used range + padding). */
	readonly colCount: number;
}

export type GridCellKind = "empty" | "number" | "text" | "bool" | "date" | "formula";

/** What the grid renders and the cell editor pre-fills for one cell. */
export interface GridCellDisplay {
	/** Rendered text (numFmt-aware); "" when the cell is empty. */
	readonly text: string;
	/** Editor pre-fill: formula with "=", else the literal value. */
	readonly raw: string;
	readonly kind: GridCellKind;
	/** Numbers, bools and dates render right-aligned like Excel. */
	readonly align: "left" | "right";
}

/**
 * Value-level grid editing contract implemented by the xlsx (ExcelJS) and
 * xls (SheetJS) engines. Rich formatting is never edited here — untouched
 * style properties survive a round-trip only as far as the underlying
 * library keeps them.
 */
export interface GridEngine {
	readonly kind: "xlsx" | "xls";
	readonly sheets: readonly SheetInfo[];
	/** Index into `sheets` of the sheet the grid is showing. */
	readonly activeSheet: number;
	/** Switches the active sheet (index into sheets). */
	selectSheet(index: number): void;
	/** Width of one column in px (stored width or the default heuristic). */
	getColumnWidth(sheet: number, col: number): number;
	getCellDisplay(sheet: number, row: number, col: number): GridCellDisplay;
	/**
	 * Applies raw editor input. Returns true when the value actually changed
	 * (callers flip the dirty flag on true only).
	 */
	setCellValue(sheet: number, row: number, col: number, raw: string): boolean;
	serialize(): Promise<Uint8Array>;
}

/** Options for find & replace over a text document. */
export interface FindReplaceOptions {
	caseSensitive: boolean;
}

/** Text-level editing contract implemented by the docx engine. */
export interface TextDocEngine {
	readonly kind: "docx";
	/** Paragraph count (body order, table cells included). */
	readonly paragraphs: readonly DocParagraphInfo[];
	/** Concatenated visible text of one paragraph. */
	getParagraphText(index: number): string;
	/**
	 * Replaces [start, end) with insertText inside one paragraph, preserving
	 * per-run formatting. Returns true when the text actually changed.
	 */
	replaceRange(index: number, start: number, end: number, insertText: string): boolean;
	/** Replaces every within-paragraph occurrence; returns the replacement count. */
	findReplaceAll(query: string, replacement: string, options: FindReplaceOptions): number;
	/** Counts occurrences the way findReplaceAll would replace them. */
	countMatches(query: string, options: FindReplaceOptions): number;
	serialize(): Promise<Uint8Array>;
}

/** What the paragraph editor needs per paragraph. */
export interface DocParagraphInfo {
	readonly text: string;
	readonly inTable: boolean;
}

/**
 * One editable text shape on a slide, as the canvas UI sees it. Geometry is
 * in EMUs; a shape without any (and without an inherited one) carries nulls
 * and is listed below the canvas instead of positioned on it.
 */
export interface SlideShapeInfo {
	readonly name: string;
	readonly placeholder: string | null;
	readonly x: number | null;
	readonly y: number | null;
	readonly cx: number | null;
	readonly cy: number | null;
	/** Paragraphs joined with "\n". */
	readonly text: string;
}

export interface SlideSize {
	readonly widthEmu: number;
	readonly heightEmu: number;
}

/** Text-level editing contract implemented by the pptx engine. */
export interface SlideCanvasEngine {
	readonly kind: "pptx";
	readonly slideSize: SlideSize;
	/** Text shapes per slide, in presentation order. */
	readonly slides: readonly SlideShapeInfo[][];
	/** Paragraphs joined with "\n". */
	getShapeText(slide: number, shape: number): string;
	/**
	 * Replaces the whole shape text; "\n" separates paragraphs. Extra lines
	 * clone the last paragraph, dropped lines remove paragraphs. Returns true
	 * when anything changed.
	 */
	replaceShapeText(slide: number, shape: number, text: string): boolean;
	/** Replaces every within-paragraph occurrence; returns the replacement count. */
	findReplaceAll(query: string, replacement: string, options: FindReplaceOptions): number;
	/** Counts occurrences the way findReplaceAll would replace them. */
	countMatches(query: string, options: FindReplaceOptions): number;
	serialize(): Promise<Uint8Array>;
}

/**
 * Read-only text extraction contract for legacy binary formats (.doc).
 * No editing surface — serialize returns the original bytes and is never
 * called in practice (no UI path can flip the dirty flag).
 */
export interface TextExtractEngine {
	readonly kind: "doc";
	/** Non-empty body paragraphs, in document order. */
	readonly paragraphs: readonly string[];
	serialize(): Promise<Uint8Array>;
}

/** Any engine the view shell can host. */
export type OfficeEngine =
	| GridEngine
	| TextDocEngine
	| SlideCanvasEngine
	| TextExtractEngine;
