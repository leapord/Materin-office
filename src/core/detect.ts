/** Which Office document a vault file holds, decided purely from its extension. */
export type OfficeKind = "docx" | "doc" | "xlsx" | "xls" | "pptx" | "ppt";

const KIND_BY_EXTENSION: Record<string, OfficeKind> = {
	docx: "docx",
	doc: "doc",
	xlsx: "xlsx",
	xls: "xls",
	pptx: "pptx",
	ppt: "ppt",
};

/** Extensions this plugin claims via registerExtensions. */
export const OFFICE_EXTENSIONS: readonly string[] = Object.keys(
	KIND_BY_EXTENSION,
);

/** Legacy binary formats have no JS writer — they open read-only or not at all. */
const EDITABLE_KINDS: ReadonlySet<OfficeKind> = new Set([
	"docx",
	"xlsx",
	"xls",
	"pptx",
]);

/** Maps a vault path to its Office kind; null for non-Office files and dotfiles. */
export function detectKind(path: string): OfficeKind | null {
	const name = path.replace(/.*[/\\]/, "");
	const dot = name.lastIndexOf(".");
	if (dot <= 0) {
		// no extension, or the name is itself a dotfile like ".gitignore"
		return null;
	}
	const ext = name.slice(dot + 1).toLowerCase();
	return KIND_BY_EXTENSION[ext] ?? null;
}

export function isEditableKind(kind: OfficeKind): boolean {
	return EDITABLE_KINDS.has(kind);
}

export const KIND_LABELS: Record<OfficeKind, string> = {
	docx: "Word 文档",
	doc: "Word 97-2003",
	xlsx: "Excel 工作簿",
	xls: "Excel 97-2003",
	pptx: "PowerPoint 演示文稿",
	ppt: "PowerPoint 97-2003",
};
