// Infer image MIME type from the file extension in the package path.
// PPTX packages may also carry this in [Content_Types].xml, but extension
// inference is sufficient for the common image formats.
//
// EMF / WMF are included here with their canonical Microsoft MIME strings
// (`image/x-emf`, `image/x-wmf`) even though browsers can't render them
// natively — the presentation loader runs an EMF/WMF → PNG conversion pass
// (see `src/emf.ts`) and the MIME flag drives that dispatch. Without the
// right MIME on the source Blob the converter can't tell EMF from WMF
// (distinct record formats, distinct decoders).
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
	emf: "image/x-emf",
	wmf: "image/x-wmf",
};

export function imageMimeFromPath(path: string): string {
	const i = path.lastIndexOf(".");
	if (i < 0) return "application/octet-stream";
	const ext = path.slice(i + 1).toLowerCase();
	return IMAGE_MIME[ext] ?? "application/octet-stream";
}
