import { A_NS } from './namespaces';
import { imageMimeFromPath } from './mime';
import { OpenXmlPackage, resolveRelTarget } from './open-xml-package';
import { parseSlide, ShapeLike, SlideParseContext } from './presentation-parser';

const NOTES_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const IMAGE_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export interface NotesSlide {
	// The parsed notes-slide shapes. Structurally identical to Slide.shapes —
	// the notesSlide part carries the same <p:cSld><p:spTree> content as a
	// regular slide, just under a <p:notes> root. Prepended with the
	// notesMaster's static chrome (header/footer/date/slide-number/slide-image
	// placeholders) so the notes page renders like a full sub-slide.
	shapes: ShapeLike[];
	// Images embedded directly on this notes slide (rare but possible).
	// Keyed by the synthetic `${partPath}#${rId}` id used by BlipFill after
	// rewriting — merged into the Presentation's global embedUrls by the
	// loader so shape renderers can look them up.
	embedUrls: Map<string, string>;
}

// Load a slide's associated notesSlide part and return the parsed shapes.
// `notesCtx` is the SlideParseContext built from the notesMaster (or the
// slide master as a fallback); it carries theme/clrMap/placeholders so
// parseSlide can resolve inherited geometry and text styling. The caller
// is responsible for merging `embedUrls` into the presentation-wide map
// and rewriting blip rIds to their synthetic keys.
export async function loadNotesSlide(
	pkg: OpenXmlPackage,
	slidePath: string,
	notesCtx: SlideParseContext,
	notesStaticShapes: ShapeLike[],
): Promise<{ notes: NotesSlide; notesPath: string } | null> {
	const rels = await pkg.loadRelationships(slidePath);
	let notesPath: string | null = null;
	for (const r of rels.values()) {
		if (r.type === NOTES_REL_TYPE) {
			notesPath = resolveRelTarget(slidePath, r.target);
			break;
		}
	}
	if (!notesPath) return null;

	const doc = await pkg.loadXml(notesPath);
	if (!doc) return null;

	// Resolve the notes slide's own image rels. Rare — most notes slides
	// carry only placeholders — but parseSlide/parsePic still need the map
	// to resolve <a:blip r:embed="..."> references. Cache blobs per URL in
	// the same way the main pipeline does.
	const embedUrls = await loadNotesEmbeds(pkg, notesPath);

	// Parse with the notesMaster-derived context, but swap in this notes
	// slide's own embed map so any images on the notes part resolve.
	const localCtx: SlideParseContext = { ...notesCtx, embedUrls };
	// parseSlide only looks at `doc.documentElement`'s <p:cSld> child, so
	// it transparently handles both <p:sld> and <p:notes> roots.
	const parsed = parseSlide(doc, 0, localCtx);

	// Prepend the notesMaster chrome. Static shapes (e.g. the sldImg preview
	// border, the notesMaster's footer background) belong under the notes
	// slide's own content, same cascade as for a regular slide → layout →
	// master. The handoutMaster layer is skipped — notes slides don't cascade
	// through a layout.
	const shapes: ShapeLike[] = [...notesStaticShapes, ...parsed.shapes];

	return {
		notes: { shapes, embedUrls },
		notesPath,
	};
}

async function loadNotesEmbeds(pkg: OpenXmlPackage, notesPath: string): Promise<Map<string, string>> {
	const urls = new Map<string, string>();
	const rels = await pkg.loadRelationships(notesPath);
	for (const rel of rels.values()) {
		if (rel.type !== IMAGE_REL_TYPE) continue;
		const mediaPath = resolveRelTarget(notesPath, rel.target);
		const blob = await pkg.loadBlob(mediaPath, imageMimeFromPath(mediaPath));
		if (blob) urls.set(rel.id, URL.createObjectURL(blob));
	}
	return urls;
}
