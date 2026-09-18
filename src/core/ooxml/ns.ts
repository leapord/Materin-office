/** OOXML namespace URIs and qualified-name helper. */

export const NS = {
	w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	a: "http://schemas.openxmlformats.org/drawingml/2006/main",
	p: "http://schemas.openxmlformats.org/presentationml/2006/main",
	r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	xml: "http://www.w3.org/XML/1998/namespace",
} as const;

export type NsPrefix = keyof typeof NS;

/** Expanded qualified name, e.g. qn("w", "t") → "{…main}t" for getElementsByTagNameNS. */
export function qn(prefix: NsPrefix, local: string): string {
	return `{${NS[prefix]}}${local}`;
}
