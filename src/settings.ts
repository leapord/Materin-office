import { App, PluginSettingTab } from "obsidian";
import type { SettingDefinitionItem } from "obsidian";
import type MaterinOfficePlugin from "./main";

export interface OfficeSettings {
	/** Copy the original file to the backup folder before the first save of each session. */
	backupEnabled: boolean;
	/** How many backups to keep per file (oldest pruned). */
	backupKeepPerFile: number;
	/** Mode a .docx opens in. */
	docxEditMode: "preview" | "edit";
	/** Show gridlines in the spreadsheet grid. */
	showXlsxGridLines: boolean;
	/** Show formula text instead of the cached value for formula cells. */
	showXlsxFormulas: boolean;
}

export const DEFAULT_SETTINGS: OfficeSettings = {
	backupEnabled: true,
	backupKeepPerFile: 10,
	docxEditMode: "preview",
	showXlsxGridLines: true,
	showXlsxFormulas: false,
};

export class OfficeSettingTab extends PluginSettingTab {
	private readonly plugin: MaterinOfficePlugin;

	constructor(app: App, plugin: MaterinOfficePlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	/**
	 * Declarative settings (1.13+): definitions drive both rendering and
	 * settings search; display() only clears the container, the framework
	 * renders the definitions itself.
	 */
	getSettingDefinitions(): SettingDefinitionItem[] {
		return [
			{
				name: "保存前备份",
				desc: "首次保存某文件前，自动把原始文件备份到插件目录，防止写回出错。",
				control: { type: "toggle", key: "backupEnabled" },
			},
			{
				name: "每文件保留备份数",
				desc: "超出数量的旧备份会被清理。",
				control: {
					type: "dropdown",
					key: "backupKeepPerFile",
					options: { "5": "5", "10": "10", "25": "25" },
				},
			},
			{
				name: "Word 默认打开模式",
				desc: "Word 文件打开时进入预览还是编辑模式。",
				control: {
					type: "dropdown",
					key: "docxEditMode",
					options: { preview: "预览", edit: "编辑" },
				},
			},
			{
				name: "显示网格线",
				control: { type: "toggle", key: "showXlsxGridLines" },
			},
			{
				name: "公式单元格显示公式",
				desc: "开启后公式单元格显示公式文本而非缓存计算值。",
				control: { type: "toggle", key: "showXlsxFormulas" },
			},
		];
	}

	/** Declarative controls read their current value through this hook. */
	getControlValue(key: string): unknown {
		return this.plugin.settings[key as keyof OfficeSettings];
	}

	/** Persists through the same path the old Setting-based UI used. */
	setControlValue(key: string, value: unknown): void {
		this.plugin.settings = { ...this.plugin.settings, [key]: value };
		void this.plugin.saveSettings();
		this.plugin.notifyViews();
	}

	display(): void {
		this.containerEl.empty();
	}
}
