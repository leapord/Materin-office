/** Augmentations for obsidian APIs missing from the bundled typings. */

// Side-effect import makes this file a module, which turns the `declare
// module "obsidian"` below into an AUGMENTATION instead of a shadowing
// ambient module declaration (which would hide the real typings).
import "obsidian";

declare module "obsidian" {
	interface App {
		/** Public since Obsidian 1.1 but absent from the bundled typings. */
		openWithDefaultApp(path: string): void;
	}
}
