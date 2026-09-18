import { Plugin, WorkspaceLeaf } from "obsidian";
import { OfficeView, VIEW_TYPE_MATERIN_OFFICE } from "./view/OfficeView";
import {
	OfficeSettingTab,
	DEFAULT_SETTINGS,
	type OfficeSettings,
} from "./settings";
import { OFFICE_EXTENSIONS } from "./core/detect";

export default class MaterinOfficePlugin extends Plugin {
	public settings: OfficeSettings = DEFAULT_SETTINGS;

	async onload(): Promise<void> {
		await this.loadSettings();
		this.registerView(
			VIEW_TYPE_MATERIN_OFFICE,
			(leaf: WorkspaceLeaf) => new OfficeView(leaf, this),
		);
		this.registerExtensions([...OFFICE_EXTENSIONS], VIEW_TYPE_MATERIN_OFFICE);
		this.addSettingTab(new OfficeSettingTab(this.app, this));
		this.registerCommands();
	}

	private registerCommands(): void {
		this.addCommand({
			id: "save",
			name: "保存文件",
			checkCallback: (checking: boolean) => {
				const view = this.activeOfficeView();
				if (!view) {
					return false;
				}
				if (!checking) {
					void view.requestSave();
				}
				return true;
			},
		});
		this.addCommand({
			id: "reload",
			name: "从磁盘重载",
			checkCallback: (checking: boolean) => {
				const view = this.activeOfficeView();
				if (!view) {
					return false;
				}
				if (!checking) {
					view.requestReload();
				}
				return true;
			},
		});
	}

	private activeOfficeView(): OfficeView | null {
		const leaf = this.app.workspace.getActiveViewOfType(OfficeView);
		return leaf;
	}

	async loadSettings(): Promise<void> {
		this.settings = {
			...DEFAULT_SETTINGS,
			...((await this.loadData()) as Partial<OfficeSettings> | null),
		};
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.settings);
	}

	/** Fans the current settings out to every open office view. */
	notifyViews(): void {
		for (const leaf of this.app.workspace.getLeavesOfType(
			VIEW_TYPE_MATERIN_OFFICE,
		)) {
			if (leaf.view instanceof OfficeView) {
				leaf.view.applySettings();
			}
		}
	}
}
