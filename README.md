# pptxjs

PPTX rendering library.

Goal
----
Render/convert a PPTX presentation into HTML, keeping HTML semantic as much as possible.

Installation
-----
```
npm install pptx-preview
```

Usage
-----
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

API
---
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

Status
------
Early scaffolding. Parser and renderer are stubs.
