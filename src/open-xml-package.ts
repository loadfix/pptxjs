import JSZip from "jszip";

export interface OpenXmlPackageOptions {
	trimXmlDeclaration: boolean;
}

export class OpenXmlPackage {
	constructor(private _zip: JSZip, public options: OpenXmlPackageOptions) {}

	static async load(input: Blob | any, options: OpenXmlPackageOptions): Promise<OpenXmlPackage> {
		const zip = await JSZip.loadAsync(input);
		return new OpenXmlPackage(zip, options);
	}

	async loadText(path: string): Promise<string | null> {
		const p = normalizePath(path);
		const entry = this._zip.files[p] ?? this._zip.files[p.replace(/\//g, "\\")];
		return entry ? await entry.async("string") : null;
	}

	async loadBlob(path: string, mimeType: string): Promise<Blob | null> {
		const p = normalizePath(path);
		const entry = this._zip.files[p] ?? this._zip.files[p.replace(/\//g, "\\")];
		if (!entry) return null;
		const buf = await entry.async("arraybuffer");
		return new Blob([buf], { type: mimeType });
	}

	parseXml(txt: string): Document {
		const cleaned = this.options.trimXmlDeclaration ? txt.replace(/^<\?xml[^?]*\?>/, "") : txt;
		return new DOMParser().parseFromString(cleaned, "application/xml");
	}

	async loadXml(path: string): Promise<Document | null> {
		const txt = await this.loadText(path);
		return txt ? this.parseXml(txt) : null;
	}

	async loadRelationships(partPath: string): Promise<Map<string, Relationship>> {
		const [dir, filename] = splitPath(partPath);
		const relsPath = `${dir}_rels/${filename}.rels`;
		const doc = await this.loadXml(relsPath);
		const out = new Map<string, Relationship>();
		if (!doc) return out;
		for (const el of Array.from(doc.documentElement.children)) {
			const id = el.getAttribute("Id");
			const type = el.getAttribute("Type");
			const target = el.getAttribute("Target");
			if (id && type && target) out.set(id, { id, type, target });
		}
		return out;
	}
}

export interface Relationship {
	id: string;
	type: string;
	target: string;
}

function normalizePath(p: string): string {
	return p.startsWith("/") ? p.slice(1) : p;
}

function splitPath(p: string): [string, string] {
	const i = p.lastIndexOf("/");
	return i < 0 ? ["", p] : [p.slice(0, i + 1), p.slice(i + 1)];
}

// Given a part path and a relationship Target (which may be relative like
// "../slideLayouts/slideLayout1.xml"), return the absolute package path.
export function resolveRelTarget(partPath: string, target: string): string {
	if (target.startsWith("/")) return target.slice(1);
	const [dir] = splitPath(partPath);
	const segments = (dir + target).split("/");
	const resolved: string[] = [];
	for (const s of segments) {
		if (s === "" || s === ".") continue;
		if (s === "..") resolved.pop();
		else resolved.push(s);
	}
	return resolved.join("/");
}
