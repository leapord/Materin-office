/**
 * BackupManager: before the first write-back of a session for a given file,
 * snapshot the on-disk bytes into the plugin config dir. Keeps at most
 * `keep` backups per file (newest retained).
 */
import { normalizePath, TFile, type App } from "obsidian";

const BACKUP_ROOT = "backups";

export class BackupManager {
	constructor(
		private readonly app: App,
		private readonly pluginDir: string,
		private readonly isEnabled: () => boolean,
		private readonly keep: () => number,
	) {}

	private dirFor(file: TFile): string {
		const slug = file.path.replace(/\//g, "__");
		return normalizePath(`${this.pluginDir}/${BACKUP_ROOT}/${slug}`);
	}

	/** True once any backup exists for this file. */
	async hasBackup(file: TFile): Promise<boolean> {
		const adapter = this.app.vault.adapter;
		try {
			const listed = await adapter.list(this.dirFor(file));
			return listed.files.length > 0;
		} catch {
			return false;
		}
	}

	/** Snapshots current disk bytes; no-op when disabled or already backed up. */
	async backupOncePerSession(file: TFile, alreadyBackedUp: boolean): Promise<boolean> {
		if (!this.isEnabled() || alreadyBackedUp) {
			return false;
		}
		const bytes = await this.app.vault.readBinary(file);
		const stamp = timestampStamp();
		const dir = this.dirFor(file);
		const adapter = this.app.vault.adapter;
		await adapter.mkdir(dir);
		const name = `${stamp}.${file.extension}`;
		await adapter.writeBinary(normalizePath(`${dir}/${name}`), bytes);
		await this.prune(file);
		return true;
	}

	private async prune(file: TFile): Promise<void> {
		const adapter = this.app.vault.adapter;
		const dir = this.dirFor(file);
		let listed;
		try {
			listed = await adapter.list(dir);
		} catch {
			return;
		}
		const keep = Math.max(this.keep(), 1);
		const names = [...listed.files].sort();
		while (names.length > keep) {
			const oldest = names.shift();
			if (oldest) {
				await adapter.remove(normalizePath(oldest)).catch(() => undefined);
			}
		}
	}
}

function timestampStamp(): string {
	const d = new Date();
	const pad = (n: number) => String(n).padStart(2, "0");
	return (
		`${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}` +
		`-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`
	);
}
