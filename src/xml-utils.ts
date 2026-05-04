export function firstChildNS(parent: Element | null, ns: string, localName: string): Element | null {
	if (!parent) return null;
	for (const c of Array.from(parent.children)) {
		if (c.namespaceURI === ns && c.localName === localName) return c;
	}
	return null;
}

export function childrenNS(parent: Element, ns: string, localName: string): Element[] {
	const out: Element[] = [];
	for (const c of Array.from(parent.children)) {
		if (c.namespaceURI === ns && c.localName === localName) out.push(c);
	}
	return out;
}
