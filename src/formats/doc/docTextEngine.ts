/**
 * doc engine: read-only text extraction from legacy binary Word files via
 * word-extractor. There is no editing surface — the view renders paragraphs
 * read-only, so serialize is never on a reachable path.
 */
// Buffer is the Electron-provided global (Obsidian desktop has nodeIntegration)
import type { TextExtractEngine } from "../engine";

/** Splits a Word body string into display paragraphs. */
export function splitBodyParagraphs(body: string): string[] {
	return body
		.split(/\r\n|\r|\n/)
		.filter((line, index, all) =>
			line.trim() !== "" || (index > 0 && index < all.length - 1),
		);
}

export class DocTextEngine implements TextExtractEngine {
	readonly kind = "doc" as const;

	private constructor(
		private readonly original: Uint8Array,
		readonly paragraphs: readonly string[],
	) {}

	static async load(bytes: Uint8Array): Promise<DocTextEngine> {
		const mod = await import("word-extractor");
		const WordExtractor = mod.default;
		const extractor = new WordExtractor();
		const document = await extractor.extract(Buffer.from(bytes));
		const body = document.getBody() ?? "";
		return new DocTextEngine(bytes, splitBodyParagraphs(body));
	}

	/** Never reachable (read-only view); returns the original bytes. */
	async serialize(): Promise<Uint8Array> {
		return this.original;
	}
}
