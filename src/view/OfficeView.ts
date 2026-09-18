/**
 * Office view shell: file state, save pipeline (mtime conflict guard,
 * backup-on-first-save), vault event handling, and the per-kind editor UI.
 */
import { ItemView, Notice, TFile, WorkspaceLeaf } from "obsidian";
import type { ViewStateResult } from "obsidian";
import type MaterinOfficePlugin from "../main";
import { detectKind, KIND_LABELS, isEditableKind, type OfficeKind } from "../core/detect";
import { OfficeFileError } from "../core/errors";
import { XlsxEngine } from "../formats/xlsx/xlsxEngine";
import { XlsEngine } from "../formats/xls/xlsEngine";
import { DocxEngine } from "../formats/docx/docxEngine";
import { PptxEngine } from "../formats/pptx/pptxEngine";
import { DocTextEngine } from "../formats/doc/docTextEngine";
import { createDomXmlAdapter } from "../core/ooxml/xmlAdapter";
import { readBinary, writeBinary } from "../io/vaultBinary";
import { BackupManager } from "../io/backup";
import { createKindUi, type KindUi } from "./kindUi";
import { showBanner, createStatusLine, type Banner } from "./common";
import type { OfficeEngine } from "../formats/engine";

export const VIEW_TYPE_MATERIN_OFFICE = "materin-office-view";

export class OfficeView extends ItemView {
	private readonly plugin: MaterinOfficePlugin;
	private file: TFile | null = null;
	private kind: OfficeKind | null = null;
	private dirty = false;
	private fileDeleted = false;
	private lastSavedMtime = 0;
	private engine: OfficeEngine | null = null;
	private kindUi: KindUi | null = null;
	private kindUiLoading = false;
	private conflictBanner: Banner | null = null;
	private deleteBanner: Banner | null = null;
	private riskBanner: Banner | null = null;
	private backupManager: BackupManager;
	private backedUpPaths = new Set<string>();

	constructor(leaf: WorkspaceLeaf, plugin: MaterinOfficePlugin) {
		super(leaf);
		this.plugin = plugin;
		this.navigation = true;
		this.backupManager = new BackupManager(
			this.app,
			plugin.manifest.dir ?? `${this.app.vault.configDir}/plugins/materin-office`,
			() => plugin.settings.backupEnabled,
			() => plugin.settings.backupKeepPerFile,
		);
	}

	getViewType(): string {
		return VIEW_TYPE_MATERIN_OFFICE;
	}

	getDisplayText(): string {
		return this.file?.name ?? "Office";
	}

	getIcon(): string {
		switch (this.kind) {
			case "docx":
			case "doc":
				return "file-text";
			case "pptx":
			case "ppt":
				return "presentation";
			default:
				return "file-spreadsheet";
		}
	}

	async setState(state: unknown, result: ViewStateResult): Promise<void> {
		const filePath = (state as { file?: unknown })?.file;
		if (typeof filePath === "string") {
			const file = this.app.vault.getAbstractFileByPath(filePath);
			if (file instanceof TFile && file !== this.file) {
				await this.loadFile(file);
			}
		}
		await super.setState(state, result);
	}

	getState(): Record<string, unknown> {
		return { ...super.getState(), file: this.file?.path };
	}

	async onOpen(): Promise<void> {
		this.contentEl.empty();
		this.contentEl.addClass("materin-office-view");
		this.registerDomEvent(
			this.contentEl,
			"keydown",
			(event: KeyboardEvent) => {
				if (
					(event.ctrlKey || event.metaKey) &&
					event.key.toLowerCase() === "s" &&
					!event.isComposing
				) {
					event.preventDefault();
					event.stopPropagation();
					void this.saveFile();
				}
			},
			{ capture: true },
		);
		this.registerVaultEvents();
	}

	async onClose(): Promise<void> {
		this.kindUi?.destroy();
		this.kindUi = null;
		this.contentEl.empty();
	}

	/** Called by the plugin when settings change. */
	applySettings(): void {
		this.kindUi?.applySettings();
	}

	/** Save command entry point (plugin command and Mod-s both land here). */
	async requestSave(): Promise<void> {
		await this.saveFile();
	}

	/** Reload command entry point. */
	requestReload(): void {
		void this.reloadFromDisk();
	}

	/** Opens the current file with the system default application. */
	openExternal(): void {
		if (!this.file) {
			return;
		}
		try {
			this.app.openWithDefaultApp(this.file.path);
		} catch (cause) {
			new Notice(
				`无法打开：${cause instanceof Error ? cause.message : String(cause)}`,
			);
		}
	}

	private async loadFile(file: TFile): Promise<void> {
		this.kindUi?.destroy();
		this.kindUi = null;
		this.engine = null;
		this.clearBanners();
		this.file = file;
		this.fileDeleted = false;
		this.dirty = false;
		this.kind = detectKind(file.path);
		this.lastSavedMtime = file.stat.mtime;
		this.renderShell();
		this.updateTitle();
		if (!this.kind) {
			return;
		}
		this.kindUiLoading = true;
		try {
			const bytes = await readBinary(this.app, file);
			const kindAtLoad = detectKind(file.path);
			this.kind = kindAtLoad;
			if (!kindAtLoad) {
				this.renderNotEditable();
				return;
			}
			if (!isEditableKind(kindAtLoad)) {
				this.renderNotEditable();
				return;
			}
			await this.mountEngine(kindAtLoad, bytes);
		} catch (cause) {
			if (this.file !== file) {
				return; // a newer load already started
			}
			const message = cause instanceof OfficeFileError
				? cause.message
				: "文件解析失败，文件可能已损坏。";
			this.renderError(message);
		} finally {
			this.kindUiLoading = false;
		}
	}

	private async mountEngine(kind: OfficeKind, bytes: ArrayBuffer): Promise<void> {
		if (kind === "xlsx") {
			this.engine = await XlsxEngine.load(new Uint8Array(bytes));
		} else if (kind === "xls") {
			this.engine = XlsEngine.load(new Uint8Array(bytes));
		} else if (kind === "docx") {
			this.engine = await DocxEngine.load(
				new Uint8Array(bytes),
				createDomXmlAdapter(),
			);
		} else if (kind === "pptx") {
			this.engine = await PptxEngine.load(
				new Uint8Array(bytes),
				createDomXmlAdapter(),
			);
		} else if (kind === "doc") {
			this.engine = await DocTextEngine.load(new Uint8Array(bytes));
		} else {
			this.engine = null;
		}
		if (!this.engine) {
			this.renderNotEditable();
			return;
		}
		this.kindUi = await createKindUi(kind, this.engine, new Uint8Array(bytes), {
			setDirty: () => this.setDirty(true),
			settings: () => this.plugin.settings,
			status: (text) => this.statusLine?.setText(text),
			openExternal: () => this.openExternal(),
		});
		this.kindUiLoading = false;
		if (!this.bodyEl) {
			return;
		}
		await this.kindUi.render(this.bodyEl);
		this.renderRiskBanner();
	}

	/**
	 * Save pipeline: guards, external-change detection, backup, write-back.
	 * Returns "saved" | "skipped" | "conflict" for the command layer.
	 */
	async saveFile(): Promise<"saved" | "skipped" | "conflict"> {
		if (!this.file || !this.engine || !this.dirty || this.fileDeleted) {
			return "skipped";
		}
		if (this.file.stat.mtime !== this.lastSavedMtime) {
			this.showConflictBanner();
			return "conflict";
		}
		try {
			await this.backupManager.backupOncePerSession(
				this.file,
				this.backedUpPaths.has(this.file.path),
			);
			this.backedUpPaths.add(this.file.path);
			const out = await this.engine.serialize();
			await writeBinary(this.app, this.file, out);
			this.lastSavedMtime = this.file.stat.mtime;
			this.setDirty(false);
			new Notice("已保存");
			return "saved";
		} catch (cause) {
			new Notice(`保存失败: ${cause instanceof Error ? cause.message : String(cause)}`);
			return "skipped";
		}
	}

	private setDirty(dirty: boolean): void {
		if (this.dirty === dirty) {
			return;
		}
		this.dirty = dirty;
		this.updateTitle();
	}

	private registerVaultEvents(): void {
		this.registerEvent(
			this.app.vault.on("modify", (file) => {
				if (!(file instanceof TFile) || file !== this.file) {
					return;
				}
				if (file.stat.mtime === this.lastSavedMtime) {
					return; // our own write-back echo
				}
				if (this.dirty) {
					this.showConflictBanner();
				} else {
					void this.reloadFromDisk();
				}
			}),
		);
		this.registerEvent(
			this.app.vault.on("delete", (file) => {
				if (file !== this.file) {
					return;
				}
				this.fileDeleted = true;
				this.showDeleteBanner();
			}),
		);
		this.registerEvent(
			this.app.vault.on("rename", (file, oldPath) => {
				if (this.file?.path !== oldPath) {
					return;
				}
				if (file instanceof TFile) {
					this.file = file;
					this.updateTitle();
					if (detectKind(file.path) !== this.kind) {
						void this.loadFile(file);
					}
				}
			}),
		);
	}

	private async reloadFromDisk(): Promise<void> {
		if (!this.file || this.kindUiLoading) {
			return;
		}
		await this.loadFile(this.file);
	}

	private showConflictBanner(): void {
		if (!this.bannerSlot) {
			return;
		}
		if (this.conflictBanner) {
			return;
		}
			this.conflictBanner = showBanner(
				this.bannerSlot,
				"文件已被外部修改。保存将覆盖外部改动。",
				"warning",
				[
					{ label: "重载（放弃我的修改）", onClick: () => void this.reloadFromDisk() },
					{ label: "仍然保存", onClick: () => void this.forceSave() },
				],
		);
	}

	private async forceSave(): Promise<void> {
		if (!this.file) {
			return;
		}
		this.lastSavedMtime = this.file.stat.mtime; // accept the external state
		this.conflictBanner?.destroy();
		this.conflictBanner = null;
		await this.saveFile();
	}

	private showDeleteBanner(): void {
		if (!this.bannerSlot) {
			return;
		}
		if (this.deleteBanner) {
			return;
		}
			this.deleteBanner = showBanner(
				this.bannerSlot,
				"文件已被删除，当前内容为只读展示。",
				"error",
				[
					{
						label: "关闭视图",
						onClick: () => this.leaf.detach(),
					},
				],
			);
	}

	private renderRiskBanner(): void {
		if (this.riskBanner || !this.bannerSlot || !(this.engine instanceof XlsxEngine)) {
			return;
		}
		const parts = this.engine.riskParts;
		if (parts.length === 0) {
			return;
		}
		this.riskBanner = showBanner(
			this.bannerSlot,
			`检测到 ${parts.join("、")}：写回可能丢失对应内容（ExcelJS 已知限制），首次保存已自动备份原文件。`,
			"warning",
		);
	}

	private clearBanners(): void {
		this.conflictBanner?.destroy();
		this.conflictBanner = null;
		this.deleteBanner?.destroy();
		this.deleteBanner = null;
		this.riskBanner?.destroy();
		this.riskBanner = null;
	}

	// ── rendering helpers ──────────────────────────────────────────────────

	private headerEl: HTMLElement | null = null;
	private bannerSlot: HTMLElement | null = null;
	private bodyEl: HTMLElement | null = null;
	private statusLine: { setText(text: string): void } | null = null;

	private renderShell(): void {
		const content = this.contentEl;
		content.empty();
		this.headerEl = content.createDiv("materin-office-header");
		this.bannerSlot = content.createDiv("materin-office-banner-slot");
		this.bodyEl = content.createDiv("materin-office-body");
		const footer = content.createDiv("materin-office-footer");
		this.statusLine = createStatusLine(footer);
		this.renderHeader();
	}

	private renderHeader(): void {
		if (!this.headerEl) {
			return;
		}
		this.headerEl.empty();
		const title = this.headerEl.createDiv("materin-office-title");
		title.createSpan({
			text: this.file?.name ?? "Office",
			cls: "materin-office-filename",
		});
		if (this.kind) {
			title.createSpan({ text: KIND_LABELS[this.kind], cls: "materin-office-badge" });
		}
		if (this.dirty) {
			title.createSpan({ text: "●", cls: "materin-office-dirty-dot" });
		}
		const actions = this.headerEl.createDiv("materin-office-actions");
		actions.createEl("button", { text: "重载", cls: "mod-muted" })
			.addEventListener("click", () => void this.reloadFromDisk());
		actions.createEl("button", { text: "保存", cls: "mod-cta" })
			.addEventListener("click", () => void this.requestSave());
	}

	private updateTitle(): void {
		this.renderHeader();

	}

	private renderNotEditable(): void {
		this.bodyEl?.empty();
		const box = this.bodyEl?.createDiv("materin-office-placeholder");
		if (!box || !this.file) {
			return;
		}
		box.createEl("h2", { text: this.file.name });
		box.createDiv({
			text: "此格式暂不支持内嵌编辑。",
			cls: "materin-office-muted",
		});
		box.createEl("button", {
			text: "用系统程序打开",
			cls: "mod-cta",
		}).addEventListener("click", () => this.openExternal());
	}

	private renderError(message: string): void {
		this.bodyEl?.empty();
		const box = this.bodyEl?.createDiv("materin-office-placeholder");
		if (!box) {
			return;
		}
		box.createEl("h2", { text: "无法打开文件" });
		box.createDiv({ text: message, cls: "materin-office-muted" });
	}
}
