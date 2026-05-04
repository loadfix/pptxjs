# pptxjs

PPTX rendering library.

This repository is a fork of [pptxjs](https://github.com/meshesha/PPTXjs) by Meshesha. It builds on their original work with a from-scratch TypeScript reimplementation that shares the project's name and goal — rendering PPTX presentations in the browser — but follows the architecture of the sibling [docxjs](https://github.com/loadfix/docxjs) fork (TypeScript + rollup + karma, JSZip for package I/O). Credit for the foundational library and the PPTXjs name goes to the original author.

## Goal

Render/convert a PPTX presentation into HTML, keeping HTML semantic as much as possible.

## Installation

```
npm install pptx-preview
```

## Usage

```html
<!--lib uses jszip-->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="pptx-preview.min.js"></script>
<script>
    var pptxData = <presentation Blob>;

    pptx.renderAsync(pptxData, document.getElementById("container"))
        .then(x => console.log("pptx: finished"));
</script>
<body>
    ...
    <div id="container"></div>
    ...
</body>
```

## API

```ts
renderAsync(
    data: Blob | ArrayBuffer | Uint8Array,
    bodyContainer: HTMLElement,
    styleContainer?: HTMLElement,
    options?: Partial<Options>
): Promise<Presentation>

parseAsync(
    data: Blob | ArrayBuffer | Uint8Array,
    options?: Partial<Options>
): Promise<Presentation>

renderPresentation(
    presentation: Presentation,
    options?: Partial<Options>
): Promise<Node[]>
```

## Status

Early scaffolding. The public surface (`renderAsync`, `parseAsync`, `renderPresentation`) is in place but may still change in shape as the parser and renderer fill in. Treat the API as unstable for now.

## Contributing

This fork commits `dist/` alongside source changes so consumers can pull from git directly. If you open a PR, rebuild `dist/` (via `npm run build`) before committing so the bundled output stays in sync with `src/`. Run `npm run e2e` (Karma + Chrome) for the browser test suite.

## Related projects

Part of a family of document-rendering libraries:

- [docxjs](https://github.com/loadfix/docxjs) — browser-side DOCX → HTML renderer (TypeScript)
- [xlsxjs](https://github.com/loadfix/xlsxjs) — browser-side XLSX → HTML renderer (TypeScript)
- [python-docx](https://github.com/loadfix/python-docx) — Python DOCX parser/generator
- [python-pptx](https://github.com/loadfix/python-pptx) — Python PPTX parser/generator
- [python-xlsx](https://github.com/loadfix/python-xlsx) — Python XLSX parser/generator
