/** Minimal ambient types for the untyped word-extractor package. */
declare module "word-extractor" {
	export default class WordExtractor {
		constructor();
		extract(source: Buffer): Promise<{
			getBody(options?: { filterUnicode?: boolean }): string | null;
		}>;
	}
}
