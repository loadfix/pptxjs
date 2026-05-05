// Infer image MIME type from the file extension in the package path.
// PPTX packages may also carry this in [Content_Types].xml, but extension
// inference is sufficient for the common image formats.
const IMAGE_MIME: Record<string, string> = {
	png: "image/png",
	jpg: "image/jpeg",
	jpeg: "image/jpeg",
	gif: "image/gif",
	bmp: "image/bmp",
	webp: "image/webp",
	svg: "image/svg+xml",
	tiff: "image/tiff",
	tif: "image/tiff",
	// EMF/WMF are not natively renderable in browsers. We still emit a
	// MIME so `loadBlob` returns the bytes (hosts may want to handle them);
	// `isBrowserRenderableImage` distinguishes these from the natively
	// supported formats so the renderer can skip the <img> fallback.
	emf: "image/x-emf",
	wmf: "image/x-wmf",
};

export function imageMimeFromPath(path: string): string {
	const i = path.lastIndexOf(".");
	if (i < 0) return "application/octet-stream";
	const ext = path.slice(i + 1).toLowerCase();
	return IMAGE_MIME[ext] ?? "application/octet-stream";
}

// Returns true when the given MIME type is something the browser can render
// as an <img>. EMF/WMF/TIFF are returned false — browsers won't decode them,
// and an <img src> with an unrenderable blob shows its `alt` text (which for
// pptxjs is the PPTX `descr` attribute, often a source filename). Callers
// can use this to suppress the <img> entirely and emit a neutral placeholder.
const BROWSER_RENDERABLE_MIME = new Set([
	"image/png",
	"image/jpeg",
	"image/gif",
	"image/bmp",
	"image/webp",
	"image/svg+xml",
]);
export function isBrowserRenderableImage(mime: string): boolean {
	return BROWSER_RENDERABLE_MIME.has(mime);
}
