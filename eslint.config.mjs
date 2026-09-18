// Mirrors the official community-plugin review lint (eslint-plugin-obsidianmd).
// Run `npx eslint .` — the goal is 0 errors, 0 warnings before resubmitting.
import { defineConfig } from "eslint/config";
import obsidianmd from "eslint-plugin-obsidianmd";

export default defineConfig([
	...obsidianmd.configs.recommended,
	{
		ignores: ["main.js", ".debug-vault/**"],
	},
	{
		languageOptions: {
			parserOptions: {
				projectService: {
					allowDefaultProject: ["eslint.config.mjs"],
				},
			},
		},
	},
]);
