/**
 * Minimal obsidian stand-in for vitest: only the symbols the tested source
 * files import at runtime. Types still come from the real obsidian package.
 */
export class PluginSettingTab {
	containerEl: HTMLElement = undefined as unknown as HTMLElement;
}
