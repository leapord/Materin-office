/**
 * Binary IO helpers against the vault (obsidian-aware).
 */
import { FileSystemAdapter, TFile, type App } from "obsidian";
import { toArrayBuffer } from "../core/bufferUtils";

export function readBinary(app: App, file: TFile): Promise<ArrayBuffer> {
	return app.vault.readBinary(file);
}

export function writeBinary(
	app: App,
	file: TFile,
	data: Uint8Array,
): Promise<void> {
	return app.vault.modifyBinary(file, toArrayBuffer(data));
}

/** Absolute filesystem path (desktop); null when the vault is not local. */
export function absolutePath(app: App, file: TFile): string | null {
	const adapter = app.vault.adapter;
	return adapter instanceof FileSystemAdapter
		? adapter.getFullPath(file.path)
		: null;
}
