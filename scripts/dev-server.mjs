// Tiny static file server for the demo. Serves the repo root on port 8766
// (configurable via PORT env var). Used by `npm run dev` / `npm run serve`.

import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { resolve, extname, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(fileURLToPath(import.meta.url), '../..');
const port = Number(process.env.PORT) || 8766;

const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.js':   'application/javascript; charset=utf-8',
    '.mjs':  'application/javascript; charset=utf-8',
    '.css':  'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.svg':  'image/svg+xml',
    '.png':  'image/png',
    '.jpg':  'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif':  'image/gif',
    '.ico':  'image/x-icon',
    '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    '.map':  'application/json; charset=utf-8',
    '.woff': 'font/woff',
    '.woff2':'font/woff2',
    '.ttf':  'font/ttf',
};

const server = createServer(async (req, res) => {
    try {
        const url = new URL(req.url, `http://localhost:${port}`);
        let path = decodeURIComponent(url.pathname);
        if (path.endsWith('/')) path += 'index.html';
        const full = normalize(resolve(root, '.' + path));
        if (!full.startsWith(root)) {
            res.writeHead(403).end('Forbidden');
            return;
        }
        const info = await stat(full);
        if (info.isDirectory()) {
            res.writeHead(301, { Location: path + '/' }).end();
            return;
        }
        const body = await readFile(full);
        res.writeHead(200, {
            'Content-Type': MIME[extname(full)] ?? 'application/octet-stream',
            'Content-Length': body.length,
            'Cache-Control': 'no-store',
        });
        res.end(body);
    } catch (err) {
        if (err.code === 'ENOENT') res.writeHead(404).end('Not Found');
        else { res.writeHead(500).end('Server Error'); console.error(err); }
    }
});

server.listen(port, () => {
    console.log(`pptxjs demo → http://localhost:${port}/`);
});
