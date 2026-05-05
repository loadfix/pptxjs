# pptxjs

PPTX rendering library.

A browser-side PPTX → HTML renderer, written from scratch in TypeScript and following the architecture of the sibling [docxjs](https://github.com/loadfix/docxjs) library (TypeScript + rollup + karma, JSZip for package I/O). Not affiliated with other libraries that share the pptxjs name.

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

### Options

Public fields of `Options` (all optional via `Partial<Options>`):

- `className: string` — CSS class prefix on rendered elements. Defaults to `"pptx"`.
- `inWrapper: boolean` — wrap output in a container div. Defaults to `true`.
- `debug: boolean` — enables extra logging. Defaults to `false`.
- `trimXmlDeclaration: boolean` — strip `<?xml ... ?>` before parsing. Defaults to `true`.
- `useBase64URL: boolean` — emit images as data URLs instead of `blob:` URLs. Defaults to `false`.
- `showHidden: boolean` — include slides marked `<p:sldId show="0">`. Defaults to `false`.
- `renderNotes: boolean` — append a `.pptx-notes` block beneath each slide. Defaults to `false`.
- `renderComments: boolean` — render comment pins as absolutely-positioned markers. Defaults to `false`.
- `convertEmf: boolean` — convert embedded EMF / WMF images to PNG via the [`emf-converter`](https://www.npmjs.com/package/emf-converter) package so brand logos and other vector assets render instead of being skipped. The library is loaded lazily; decks without any EMF/WMF pay no bundle weight. Defaults to `true`.
- `onSlideError?: (slideIndex: number, err: unknown) => void` — called when a slide fails to parse. The slide still occupies its index in `presentation.slides` with `parseError` set and renders as a red error banner; this hook just lets hosts log / telemeter the failure.

### EMF / WMF support in UMD builds

The ESM build lazy-loads `emf-converter` automatically via a bundler-backed dynamic import. The UMD build (`dist/pptx-preview.js`) can't resolve the bare `emf-converter` specifier on its own — script-tag consumers who want metafile conversion must wire it in themselves:

```html
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="pptx-preview.min.js"></script>
<script type="module">
  import * as emfConverter from 'https://esm.sh/emf-converter';
  window.pptx.setEmfConverter(emfConverter);
</script>
```

Without this call, EMF/WMF images on UMD builds are silently skipped (their `<img>` slots stay empty).

## Status

Early scaffolding. The public surface (`renderAsync`, `parseAsync`, `renderPresentation`) is in place but may still change in shape as the parser and renderer fill in. Treat the API as unstable for now.

## Contributing

This project commits `dist/` alongside source changes so consumers can pull from git directly. If you open a PR, rebuild `dist/` (via `npm run build`) before committing so the bundled output stays in sync with `src/`. Run `npm run e2e` (Karma + Chrome) for the browser test suite.

## Related projects

Part of a family of document-rendering libraries:

- [docxjs](https://github.com/loadfix/docxjs) — browser-side DOCX → HTML renderer (TypeScript)
- [xlsxjs](https://github.com/loadfix/xlsxjs) — browser-side XLSX → HTML renderer (TypeScript)
- [python-docx](https://github.com/loadfix/python-docx) — Python DOCX parser/generator
- [python-pptx](https://github.com/loadfix/python-pptx) — Python PPTX parser/generator
- [python-xlsx](https://github.com/loadfix/python-xlsx) — Python XLSX parser/generator
