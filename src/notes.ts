import { A_NS } from './namespaces';
import { OpenXmlPackage, resolveRelTarget } from './open-xml-package';
import { firstChildNS, childrenNS } from './xml-utils';

const NOTES_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";

export interface NotesSlide {
	// Concatenated speaker-notes text from the body placeholder. Paragraphs
	// are joined with "\n"; runs within a paragraph are concatenated with no
	// separator. Rich-text rendering (bullets, fonts, hyperlinks) is a
	// follow-up — the notesSlide part shares the same text model as a regular
	// slide, but for now we surface only the plain string.
	text: string;
}

// Locate this slide's notesSlide part via its rels and pull out the body
// text. Returns null when the slide has no speaker notes (either no rel
// entry or an empty body placeholder).
export async function loadNotesSlide(pkg: OpenXmlPackage, slidePath: string): Promise<NotesSlide | null> {
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

	// Find <p:sp> whose <p:ph type="body"> (or no explicit type, defaulting
	// to body per ECMA-376) lives under <p:spTree>.
	const cSld = firstChildNS(doc.documentElement, A_NS.p, "cSld");
	const spTree = cSld ? firstChildNS(cSld, A_NS.p, "spTree") : null;
	if (!spTree) return null;

	let bodySp: Element | null = null;
	for (const sp of childrenNS(spTree, A_NS.p, "sp")) {
		const nvSpPr = firstChildNS(sp, A_NS.p, "nvSpPr");
		const nvPr = nvSpPr ? firstChildNS(nvSpPr, A_NS.p, "nvPr") : null;
		const ph = nvPr ? firstChildNS(nvPr, A_NS.p, "ph") : null;
		if (!ph) continue;
		const phType = ph.getAttribute("type");
		// body placeholders either carry type="body" or omit type entirely.
		if (phType === "body" || phType === null) {
			bodySp = sp;
			break;
		}
	}
	if (!bodySp) return null;

	const txBody = firstChildNS(bodySp, A_NS.p, "txBody");
	if (!txBody) return null;

	const paraTexts: string[] = [];
	for (const p of childrenNS(txBody, A_NS.a, "p")) {
		let line = "";
		for (const child of Array.from(p.children)) {
			if (child.namespaceURI !== A_NS.a) continue;
			if (child.localName === "r") {
				const t = firstChildNS(child, A_NS.a, "t");
				if (t) line += t.textContent ?? "";
			} else if (child.localName === "br") {
				line += "\n";
			} else if (child.localName === "fld") {
				const t = firstChildNS(child, A_NS.a, "t");
				if (t) line += t.textContent ?? "";
			}
		}
		paraTexts.push(line);
	}

	const text = paraTexts.join("\n");
	if (!text) return null;
	return { text };
}
