# Introduction to Web Development

This folder contains a browser-based classroom presentation and its editable PowerPoint export.

## Present in a browser

Open `index.html` in a modern browser. Use the Previous and Next buttons, the left and right arrow keys, or swipe on a touchscreen. Press `P` or `F` to hide the control bars for presentation mode. Print from the browser for a slide-per-page print layout.

## Generate the PowerPoint file

The included `web-development-introduction.pptx` is ready to use. To create a new editable version, run the export script from the project folder with the bundled Node.js runtime used by Codex:

```powershell
& "C:\Users\sande\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" .\outputs\export-pptx.mjs
```

The script uses `@oai/artifact-tool`, a JavaScript PowerPoint authoring library. It creates a versioned `.pptx` in this folder, preserves editable text, native table content, editable diagram shapes, speaker notes, and the supplied cover illustration.

## Included files

- `index.html`, `style.css`, and `script.js`: the live web presentation
- `export-pptx.mjs`: the editable PowerPoint export script
- `web-development-introduction.pptx`: generated PowerPoint presentation
- `assets/web-development-network-hero.png`: presentation cover and closing artwork
