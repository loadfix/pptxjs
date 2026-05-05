// Shared-class nomenclature — the `oox-*` family — emitted alongside
// the format-specific `pptx-*` classes so cross-format selectors in
// manifests can target the same logical concept across docxjs /
// pptxjs / xlsxjs without caring which renderer produced the DOM.
//
// The table below is the canonical mapping; keep it in sync with the
// equivalent file in the sibling renderers (docxjs/src/shared-classes.ts,
// xlsxjs/src/shared-classes.ts).
//
//   concept       | docxjs            | pptxjs            | xlsxjs
//   --------------|-------------------|-------------------|------------------
//   wrapper       | .docx-wrapper     | (outer host node) | (outer host node)
//   page/slide    | .docx             | .pptx-slide       | .xlsx
//   paragraph     | <p>               | <p>               | (n/a)
//   run           | <span>            | <span>            | (n/a)
//   heading       | <h1..h6>          | (n/a)             | (n/a)
//   table         | <table>           | .pptx-table       | <table>
//   table-row     | <tr>              | <tr>              | <tr>
//   table-cell    | <td>/<th>         | <td>              | <td>/<th>
//   image         | <img>             | .pptx-pic         | .xlsx-image

export type OoxConcept =
	| "wrapper"
	| "page"
	| "paragraph"
	| "run"
	| "heading"
	| "table"
	| "table-row"
	| "table-cell"
	| "image";

export const OOX_CLASSES: Record<OoxConcept, string> = {
	"wrapper":    "oox-wrapper",
	"page":       "oox-page",
	"paragraph":  "oox-paragraph",
	"run":        "oox-run",
	"heading":    "oox-heading",
	"table":      "oox-table",
	"table-row":  "oox-table-row",
	"table-cell": "oox-table-cell",
	"image":      "oox-image",
};

// Add the shared `.oox-<concept>` class to an element in addition to
// whatever format-specific class it already carries. Safe to call with
// `null` (does nothing).
export function addSharedClass(el: Element | null | undefined, concept: OoxConcept): void {
	if (!el) return;
	el.classList.add(OOX_CLASSES[concept]);
}
