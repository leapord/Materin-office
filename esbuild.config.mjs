import esbuild from "esbuild";
import process from "process";
import { cpSync, existsSync, watch } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { builtinModules } from "node:module";

const BANNER_TEXT = "Materin Office — built bundle. Do not edit manually.";

const rootDir = path.dirname(fileURLToPath(import.meta.url));
const debugVaultPluginDir = path.join(
	rootDir,
	".debug-vault/.obsidian/plugins/materin-office",
);

/** Auto-deploy built artifacts into the debug vault when it exists. */
const copyToDebugVault = {
	name: "copy-to-debug-vault",
	setup(build) {
		build.onEnd(() => {
			if (!existsSync(debugVaultPluginDir)) {
				return;
			}
			for (const file of ["main.js", "manifest.json", "styles.css"]) {
				cpSync(path.join(rootDir, file), path.join(debugVaultPluginDir, file));
			}
		});
	},
};

const prod = process.argv[2] === "production";

const context = await esbuild.context({
	banner: {
		js: `/* ${BANNER_TEXT}${prod ? "" : " (dev)"} */`,
	},
	entryPoints: ["src/main.ts"],
	// Replace the `immediate` package (jszip → lie) with a shim that keeps the
	// same queue semantics but drops the <script>-injection and `new Function`
	// fallbacks flagged by Obsidian's plugin review. See src/shims/immediate.ts.
	alias: {
		// jszip's and exceljs's main/browser entries are pre-bundled UMDs with
		// promise-scheduler polyfills (immediate, setimmediate) inlined — their
		// <script>-injection and `new Function` fallbacks get flagged by Obsidian's
		// plugin review. Point at their unbundled source so the shim aliases below
		// actually intercept, and swap the polyfills for clean shims.
		jszip: "jszip/lib/index.js",
		exceljs: "exceljs/lib/exceljs.browser.js",
		immediate: "./src/shims/immediate.ts",
		setimmediate: "./src/shims/setimmediate.ts",
	},
	bundle: true,
	external: ["obsidian", "electron", ...builtinModules],
	format: "cjs",
	target: "es2018",
	plugins: [copyToDebugVault],
	logLevel: "info",
	sourcemap: prod ? false : "inline",
	treeShaking: true,
	outfile: "main.js",
	minify: prod,
});

if (prod) {
	await context.rebuild();
	process.exit(0);
} else {
	await context.watch();
	// esbuild's watcher only tracks build inputs — styles.css/manifest.json
	// are not among them, so watch them directly and deploy on change
	// (otherwise style edits never reach the debug vault).
	for (const asset of ["styles.css", "manifest.json"]) {
		watch(path.join(rootDir, asset), { persistent: false }, () => {
			if (!existsSync(debugVaultPluginDir)) {
				return;
			}
			cpSync(path.join(rootDir, asset), path.join(debugVaultPluginDir, asset));
		});
	}
}
