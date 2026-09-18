# Materin Office

在 Obsidian 中直接读取、编辑并保存 Office 文件：`.docx`、`.xlsx`、`.pptx`、`.xls`、`.doc`、`.ppt`。

Materin 系列插件之三（兄弟插件：[materin-ctx](../Materin-ctx)、[materin-view](../Materin-view)）。

## 特性

- **xlsx**：网格编辑器，单元格值编辑、公式显示、工作表切换；样式读取完整。
- **xls**：与 xlsx 共享网格编辑器（值级编辑，样式写回有限）。
- **docx**：高保真预览 + 文本级编辑（多 run 文本替换算法），原格式 100% 保留。
- **pptx**：幻灯片画布 + 文本框编辑，仅重写被编辑的部件，其余字节不变。
- **doc**：只读文本提取（word-extractor），不可编辑，可一键用系统程序打开。
- **ppt**：暂不支持，规划提供外部打开入口。
- **安全网**：首次保存前自动备份原文件（可关）；外部修改冲突检测。

## 开发

```bash
npm install
npm run dev      # watch 模式，产物自动拷入 debug-vault
npm run build    # 类型检查 + 生产构建
npm run test     # vitest
npm run lint     # eslint（目标 0/0）
```

调试：仓库根目录的 `.debug-vault` 是指向 `../debug-vault` 的符号链接；vault 内已安装 hot-reload 插件，`npm run dev` 后在 Obsidian 中打开该 vault 即可热载。

## 已知限制

- xlsx 保存由 ExcelJS 重写整个工作簿：**图表、图片、透视缓存可能丢失**（打开含这些内容的文件会有警告横幅；首次保存自动备份）。
- xls（BIFF 格式）保存仅保留单元格值，样式基本丢失——建议转存为 xlsx。
- docx/pptx 为文本级编辑：改文字、查找替换；不支持增删改格式（加粗、字号等）。
- pptx 画布仅覆盖 p:sp 文本形状：表格/图表内的文字暂不可编辑；无本地 xfrm 的占位符几何从 slideLayout 继承（不再向上找 slideMaster，个别母版布局的占位符会落入"未定位文本框"区）。
- 公式由 Excel 打开时重算，本插件不重算公式。
