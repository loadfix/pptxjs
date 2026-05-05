// Hyperlink-wrapping helpers, shared by text/shape/pic renderers. Lives in
// its own module so shape.ts / pic.ts / text.ts can all import it without
// creating cycles through dispatch.ts (which in turn imports those modules).

// Wrap `inner` in an `<a>` anchor if `rId` resolves to a URL on `urls`.
// External http(s)/mailto URLs open in a new tab with rel="noopener
// noreferrer"; intra-deck `#slide-N` links stay in-page. Returns `inner`
// unchanged when rId is null, the rel is missing, or the URL is empty.
//
// URL safety is enforced earlier — `urls` is built by `presentation.ts`
// which drops unsafe schemes (javascript:, data:, etc.) before they reach
// this map.
export function wrapInHyperlink(
	inner: HTMLElement,
	rId: string | null,
	urls: Map<string, string>,
): HTMLElement {
	if (!rId) return inner;
	const href = urls.get(rId);
	if (!href) return inner;
	const a = document.createElement("a");
	a.href = href;
	if (!href.startsWith("#")) {
		a.target = "_blank";
		a.rel = "noopener noreferrer";
	}
	// Wrappers shouldn't introduce visual styling — let the wrapped content
	// decide its own color/decoration.
	a.style.textDecoration = "inherit";
	a.style.color = "inherit";
	a.appendChild(inner);
	return a;
}
