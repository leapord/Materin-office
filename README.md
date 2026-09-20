# Materin Office

Read, edit and save Office files directly in Obsidian: `.docx`, `.xlsx`, `.pptx`, `.xls`, `.doc`, `.ppt`.

Part of the Materin plugin series (sibling plugins: [materin-ctx](../Materin-ctx), [materin-view](../Materin-view)). [中文说明](README.zh-CN.md)

## Features

- **xlsx**: Grid editor with cell value editing, formula display, and sheet switching; full style reading.
- **xls**: Shares the grid editor with xlsx (value-level editing; limited style write-back).
- **docx**: High-fidelity preview + text-level editing (multi-run text replacement algorithm), preserving 100% of the original formatting.
- **pptx**: Slide canvas + text box editing; only edited parts are rewritten, all other bytes remain untouched.
- **doc**: Read-only text extraction (word-extractor); not editable, with one-click open in a system app.
- **ppt**: Not yet supported; an external-open entry point is planned.
- **Safety net**: Automatic backup of the original file before first save (can be disabled); external modification conflict detection.

## Development

```bash
npm install
npm run dev      # watch mode; output is copied into the debug vault
npm run build    # type check + production build
npm run test     # vitest
npm run lint     # eslint (target: 0/0)
```

Debugging: `.debug-vault` at the repo root is a symlink to `../debug-vault`. A hot-reload plugin is installed in that vault — after `npm run dev`, open the vault in Obsidian and changes hot-load.

## Known Limitations

- xlsx saving rewrites the entire workbook via ExcelJS: **charts, images, and pivot caches may be lost** (a warning banner appears when opening files that contain them; the first save creates an automatic backup).
- xls (BIFF format) saving keeps cell values only; styles are mostly lost — converting to xlsx is recommended.
- docx/pptx editing is text-level: change text, find & replace; formatting changes (bold, font size, etc.) are not supported.
- The pptx canvas only covers `p:sp` text shapes: text inside tables/charts is not yet editable; placeholders without a local `xfrm` inherit geometry from slideLayout (no longer searching slideMaster, so placeholders of some master layouts may fall into the "unpositioned text boxes" area).
- Formulas are recalculated by Excel when the file is opened; this plugin does not recalculate formulas.

## License

[MIT](LICENSE)
