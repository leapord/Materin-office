import { defineConfig } from "vitest/config";
import { fileURLToPath } from "node:url";

export default defineConfig({
	resolve: {
		alias: {
			// obsidian ships typings only — tests run against this stub.
			obsidian: fileURLToPath(
				new URL("./tests/obsidianStub.ts", import.meta.url),
			),
		},
	},
});
