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
};

export function imageMimeFromPath(path: string): string {
	const i = path.lastIndexOf(".");
	if (i < 0) return "application/octet-stream";
	const ext = path.slice(i + 1).toLowerCase();
	return IMAGE_MIME[ext] ?? "application/octet-stream";
}
