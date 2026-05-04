import type { Options } from './pptx-preview';
import { OpenXmlPackage, resolveRelTarget } from './open-xml-package';
import { imageMimeFromPath } from './mime';
import {
	parseSlide,
	parseStaticShapes,
	Slide,
	SlideSize,
	buildPlaceholderMap,
	emptyPlaceholderMap,
	emptyMasterTextStyles,
	parseMasterTextStyles,
	PlaceholderMap,
	MasterTextStyles,
} from './presentation-parser';
import {
	ThemeColors,
	ClrMap,
	emptyThemeColors,
	emptyClrMap,
	parseTheme,
	parseClrMap,
} from './theme';
import { parseSlideBackground } from './background';
import { A_NS } from './namespaces';

const LAYOUT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const THEME_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

interface MasterBundle {
	placeholders: PlaceholderMap;
	textStyles: MasterTextStyles;
	clrMap: ClrMap;
	theme: ThemeColors;
	path: string | null;
	doc: Document | null;
}

interface LayoutBundle {
	placeholders: PlaceholderMap;
	doc: Document | null;
}

export class Presentation {
	slides: Slide[] = [];
	slideSize: SlideSize = { cx: 9144000, cy: 6858000 };

	static async load(data: Blob | any, _parser: unknown, options: Options): Promise<Presentation> {
		const pkg = await OpenXmlPackage.load(data, { trimXmlDeclaration: options.trimXmlDeclaration });
		const pres = new Presentation();

		const presPath = "ppt/presentation.xml";
		const presDoc = await pkg.loadXml(presPath);
		if (!presDoc) return pres;

		const sldSzEl = presDoc.getElementsByTagNameNS(A_NS.p, "sldSz")[0];
		if (sldSzEl) {
			pres.slideSize = {
				cx: Number(sldSzEl.getAttribute("cx")) || pres.slideSize.cx,
				cy: Number(sldSzEl.getAttribute("cy")) || pres.slideSize.cy,
			};
		}

		const rels = await pkg.loadRelationships(presPath);
		const sldIdEls = Array.from(presDoc.getElementsByTagNameNS(A_NS.p, "sldId"));
		const slidePaths: string[] = [];
		for (const el of sldIdEls) {
			const rid = el.getAttributeNS(A_NS.r, "id");
			if (!rid) continue;
			const rel = rels.get(rid);
			if (!rel) continue;
			slidePaths.push(resolveRelTarget(presPath, rel.target));
		}

		const layoutCache = new Map<string, LayoutBundle>();
		const masterCache = new Map<string, MasterBundle>();
		const layoutToMasterPath = new Map<string, string | null>();

		async function getLayoutAndMaster(slidePath: string): Promise<{
			layout: LayoutBundle;
			master: MasterBundle;
			layoutPath: string | null;
			masterPath: string | null;
		}> {
			const emptyMaster: MasterBundle = {
				placeholders: emptyPlaceholderMap(),
				textStyles: emptyMasterTextStyles(),
				clrMap: emptyClrMap(),
				theme: emptyThemeColors(),
				path: null,
				doc: null,
			};
			const emptyLayout: LayoutBundle = { placeholders: emptyPlaceholderMap(), doc: null };

			const slideRels = await pkg.loadRelationships(slidePath);
			let layoutPath: string | null = null;
			for (const r of slideRels.values()) {
				if (r.type === LAYOUT_REL_TYPE) {
					layoutPath = resolveRelTarget(slidePath, r.target);
					break;
				}
			}
			if (!layoutPath) return { layout: emptyLayout, master: emptyMaster, layoutPath: null, masterPath: null };

			let layoutBundle = layoutCache.get(layoutPath);
			if (!layoutBundle) {
				const layoutDoc = await pkg.loadXml(layoutPath);
				layoutBundle = {
					placeholders: layoutDoc ? buildPlaceholderMap(layoutDoc) : emptyPlaceholderMap(),
					doc: layoutDoc,
				};
				layoutCache.set(layoutPath, layoutBundle);

				const layoutRels = await pkg.loadRelationships(layoutPath);
				let masterPath: string | null = null;
				for (const r of layoutRels.values()) {
					if (r.type === MASTER_REL_TYPE) {
						masterPath = resolveRelTarget(layoutPath, r.target);
						break;
					}
				}
				layoutToMasterPath.set(layoutPath, masterPath);
			}

			const masterPath = layoutToMasterPath.get(layoutPath) ?? null;
			let masterBundle = masterPath ? masterCache.get(masterPath) : undefined;
			if (masterPath && !masterBundle) {
				const masterDoc = await pkg.loadXml(masterPath);
				let theme = emptyThemeColors();
				if (masterDoc) {
					const masterRels = await pkg.loadRelationships(masterPath);
					for (const r of masterRels.values()) {
						if (r.type !== THEME_REL_TYPE) continue;
						const themePath = resolveRelTarget(masterPath, r.target);
						const themeDoc = await pkg.loadXml(themePath);
						if (themeDoc) theme = parseTheme(themeDoc);
						break;
					}
				}
				masterBundle = masterDoc
					? {
						placeholders: buildPlaceholderMap(masterDoc),
						textStyles: parseMasterTextStyles(masterDoc),
						clrMap: parseClrMap(masterDoc),
						theme,
						path: masterPath,
						doc: masterDoc,
					}
					: { ...emptyMaster, path: masterPath };
				masterCache.set(masterPath!, masterBundle);
			}

			return {
				layout: layoutBundle,
				master: masterBundle ?? emptyMaster,
				layoutPath,
				masterPath,
			};
		}

		// Cache image URLs across slides — the same embedded image may be
		// referenced by multiple slides via separate rIds, but there's one
		// media entry per unique file.
		const mediaUrlCache = new Map<string, string>();

		async function loadSlideEmbeds(slidePath: string): Promise<Map<string, string>> {
			const urls = new Map<string, string>();
			const slideRels = await pkg.loadRelationships(slidePath);
			for (const rel of slideRels.values()) {
				if (rel.type !== IMAGE_REL_TYPE) continue;
				const mediaPath = resolveRelTarget(slidePath, rel.target);
				let url = mediaUrlCache.get(mediaPath);
				if (!url) {
					const blob = await pkg.loadBlob(mediaPath, imageMimeFromPath(mediaPath));
					if (blob) {
						url = URL.createObjectURL(blob);
						mediaUrlCache.set(mediaPath, url);
					}
				}
				if (url) urls.set(rel.id, url);
			}
			return urls;
		}

		// Per-part embed caches keyed by package path, so layouts/masters with
		// their own images (e.g. a logo on the layout) resolve too.
		async function embedsFor(partPath: string): Promise<Map<string, string>> {
			return loadSlideEmbeds(partPath);
		}

		for (let i = 0; i < slidePaths.length; i++) {
			const path = slidePaths[i];
			const doc = await pkg.loadXml(path);
			if (!doc) continue;
			const { layout, master, layoutPath, masterPath } = await getLayoutAndMaster(path);

			// Shared parse context (clrMap/theme/etc.) that also needs the
			// correct embedUrls for whichever part is being parsed.
			const baseCtx = {
				layout: layout.placeholders,
				master: master.placeholders,
				masterTextStyles: master.textStyles,
				theme: master.theme,
				clrMap: master.clrMap,
			};

			// Static chrome from master → layout, parsed with their own embeds
			// so images stored at the master/layout level resolve.
			const staticShapes: ReturnType<typeof parseStaticShapes> = [];
			if (master.doc && masterPath) {
				const masterEmbeds = await embedsFor(masterPath);
				staticShapes.push(...parseStaticShapes(master.doc, { ...baseCtx, embedUrls: masterEmbeds }));
			}
			if (layout.doc && layoutPath) {
				const layoutEmbeds = await embedsFor(layoutPath);
				staticShapes.push(...parseStaticShapes(layout.doc, { ...baseCtx, embedUrls: layoutEmbeds }));
			}

			const slideEmbeds = await embedsFor(path);
			const slide = parseSlide(doc, i, { ...baseCtx, embedUrls: slideEmbeds });
			// Prepend static chrome so slide content renders on top.
			slide.shapes = [...staticShapes, ...slide.shapes];

			// Background cascade: slide → layout → master. First hit wins.
			slide.background =
				parseSlideBackground(doc, master.clrMap, master.theme)
				?? (layout.doc ? parseSlideBackground(layout.doc, master.clrMap, master.theme) : null)
				?? (master.doc ? parseSlideBackground(master.doc, master.clrMap, master.theme) : null);
			pres.slides.push(slide);
		}

		return pres;
	}
}
