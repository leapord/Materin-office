/*
 * Build-time replacement for the `setimmediate` package (YuzuJS), pulled in via
 * jszip/unzipper. The original polyfill's old-IE fallbacks inject <script>
 * elements and use `new Function` string coercion, which Obsidian's plugin
 * static review flags. Obsidian runs on modern Chromium with native
 * setImmediate, so this shim is effectively a no-op there.
 */

const win = window as unknown as Record<string, unknown>;

if (typeof win.setImmediate === "undefined") {
	// setTimeout with 0 delay is the universally available fallback.
	win.setImmediate = (callback: (...args: unknown[]) => void, ...args: unknown[]) =>
		window.setTimeout(callback, 0, ...args);
}
