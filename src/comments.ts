import { A_NS } from './namespaces';
import { OpenXmlPackage, resolveRelTarget } from './open-xml-package';
import { firstChildNS, childrenNS } from './xml-utils';

const COMMENTS_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const COMMENT_AUTHORS_PATH = "ppt/commentAuthors.xml";

export interface Comment {
	// Per-slide index from <p:cm idx="…"> — unique within the slide. Stored
	// as a string because the XML attribute is a free-form NMTOKEN; callers
	// usually just need a stable key.
	id: string;
	authorName: string;
	text: string;
	// ISO-8601 timestamp from <p:cm dt="…"> when present. Empty string when
	// the PPTX omitted the datetime (rare but valid per schema).
	createdISO: string;
	// Raw EMU coordinates relative to the slide's top-left, as emitted by
	// <p:pos x="…" y="…">. Callers that want pixels should divide by 9525
	// (EMU_PER_PX) — we deliberately keep the raw units so consumers can
	// choose their own unit conversion.
	x: number;
	y: number;
}

// Parse ppt/commentAuthors.xml into { authorId → displayName }. When the
// file doesn't exist (presentation has no comments) an empty map is
// returned, letting loadComments no-op gracefully.
export async function loadCommentAuthors(pkg: OpenXmlPackage): Promise<Map<string, string>> {
	const map = new Map<string, string>();
	const doc = await pkg.loadXml(COMMENT_AUTHORS_PATH);
	if (!doc) return map;
	for (const cmAuthor of childrenNS(doc.documentElement, A_NS.p, "cmAuthor")) {
		const id = cmAuthor.getAttribute("id");
		const name = cmAuthor.getAttribute("name") ?? "";
		if (id) map.set(id, name);
	}
	return map;
}

// Pull the per-slide <p:cm> entries out of the slide's comments part.
// Returns [] when the slide has no comments rel or the part is empty.
export async function loadComments(
	pkg: OpenXmlPackage,
	slidePath: string,
	authors: Map<string, string>,
): Promise<Comment[]> {
	const rels = await pkg.loadRelationships(slidePath);
	let commentsPath: string | null = null;
	for (const r of rels.values()) {
		if (r.type === COMMENTS_REL_TYPE) {
			commentsPath = resolveRelTarget(slidePath, r.target);
			break;
		}
	}
	if (!commentsPath) return [];

	const doc = await pkg.loadXml(commentsPath);
	if (!doc) return [];

	const out: Comment[] = [];
	// Document root is <p:cmLst>; each comment is a <p:cm> child.
	for (const cm of childrenNS(doc.documentElement, A_NS.p, "cm")) {
		const authorId = cm.getAttribute("authorId") ?? "";
		const idx = cm.getAttribute("idx") ?? "";
		const dt = cm.getAttribute("dt") ?? "";
		const pos = firstChildNS(cm, A_NS.p, "pos");
		const x = pos ? Number(pos.getAttribute("x")) || 0 : 0;
		const y = pos ? Number(pos.getAttribute("y")) || 0 : 0;
		const textEl = firstChildNS(cm, A_NS.p, "text");
		const text = textEl?.textContent ?? "";
		out.push({
			id: idx,
			authorName: authors.get(authorId) ?? "",
			text,
			createdISO: dt,
			x,
			y,
		});
	}
	return out;
}
