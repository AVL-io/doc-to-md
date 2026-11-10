// Build: bundles the app into dist/ (PWA) and dist/doc-to-md-offline.html (single file).
import * as esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { execSync } from 'node:child_process';

const root = path.dirname(new URL(import.meta.url).pathname);
const src = path.join(root, 'src');
const dist = path.join(root, 'dist');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
let sha = '';
try { sha = execSync('git rev-parse --short HEAD', { cwd: root, stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim(); } catch { /* not a git checkout */ }
const VERSION = `${pkg.version}${sha ? '+' + sha : ''} (${new Date().toISOString().slice(0, 10)})`;
const workerSrc = path.join(root, 'node_modules/pdfjs-dist/build/pdf.worker.min.mjs');

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(dist, { recursive: true });

const common = {
  bundle: true,
  minify: true,
  target: ['chrome109', 'edge109', 'firefox115', 'safari16'],
  define: { __APP_VERSION__: JSON.stringify(VERSION) },
  legalComments: 'none',
  logLevel: 'warning',
  external: ['buffer'], // only used by the docx library outside browsers
};

// ---------- 1. PWA build (code split so large libraries load only when needed)
await esbuild.build({
  ...common,
  entryPoints: [path.join(src, 'app.js')],
  outdir: dist,
  format: 'esm',
  splitting: true,
  chunkNames: 'chunks/[name]-[hash]',
});
for (const f of ['index.html', 'styles.css', 'manifest.webmanifest']) fs.copyFileSync(path.join(src, f), path.join(dist, f));
fs.cpSync(path.join(src, 'icons'), path.join(dist, 'icons'), { recursive: true });
fs.copyFileSync(workerSrc, path.join(dist, 'pdf.worker.min.mjs'));

const files = [];
const walk = (d) => { for (const f of fs.readdirSync(d)) { const p = path.join(d, f); if (fs.statSync(p).isDirectory()) walk(p); else files.push(path.relative(dist, p).split(path.sep).join('/')); } };
walk(dist);
const precache = ['./', ...files.filter(f => f !== 'index.html')];
const hash = crypto.createHash('sha256');
for (const f of files.sort()) hash.update(f).update(fs.readFileSync(path.join(dist, f)));
const cacheVersion = hash.digest('hex').slice(0, 12);
fs.writeFileSync(path.join(dist, 'sw.js'), fs.readFileSync(path.join(src, 'sw.js'), 'utf8')
  .replace('__VERSION__', cacheVersion)
  .replace('__PRECACHE__', JSON.stringify(precache)));

// ---------- 2. Single-file offline build
const one = await esbuild.build({
  ...common,
  entryPoints: [path.join(src, 'app.js')],
  format: 'esm',
  splitting: false,
  write: false,
  outdir: 'out',
});
const mainJs = one.outputFiles.find(f => f.path.endsWith('.js')).contents;
const b64 = (buf) => Buffer.from(buf).toString('base64');
const css = fs.readFileSync(path.join(src, 'styles.css'), 'utf8');
const iconData = 'data:image/svg+xml;base64,' + b64(fs.readFileSync(path.join(src, 'icons/icon.svg')));
const bootstrap = "(function(){var d=function(id){var s=atob(document.getElementById(id).textContent.trim()),a=new Uint8Array(s.length);for(var i=0;i<s.length;i++)a[i]=s.charCodeAt(i);return URL.createObjectURL(new Blob([a],{type:'text/javascript'}));};window.__PDF_WORKER_URL__=d('pdfw');var m=document.createElement('script');m.type='module';m.src=d('appjs');document.body.appendChild(m);})();";
const shaB = crypto.createHash('sha256').update(bootstrap).digest('base64');
let html = fs.readFileSync(path.join(src, 'index.html'), 'utf8');
html = html
  .replace(/<meta http-equiv="Content-Security-Policy"[^>]*>/, `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'sha256-${shaB}' blob: 'wasm-unsafe-eval'; style-src 'unsafe-inline'; img-src data: blob:; font-src data:; worker-src blob:; connect-src blob: data:; object-src 'none'; base-uri 'none'; form-action 'none'">`)
  .replace(/<link rel="manifest"[^>]*>\n?/, '')
  .replace(/<link rel="apple-touch-icon"[^>]*>\n?/, '')
  .replace(/<link rel="icon"[^>]*>/, `<link rel="icon" href="${iconData}">`)
  .replace(/<link rel="stylesheet"[^>]*>/, `<style>\n${css}\n</style>`)
  .replace(/<script type="module" src="app.js"><\/script>\n?/, '')
  .replace('src="icons/icon.svg"', `src="${iconData}"`)
  .replace('</body>', `<script type="application/octet-stream" id="pdfw">${b64(fs.readFileSync(workerSrc))}</script>\n<script type="application/octet-stream" id="appjs">${b64(mainJs)}</script>\n<script>${bootstrap}</script>\n</body>`);
fs.writeFileSync(path.join(dist, 'doc-to-md-offline.html'), html);

const size = (f) => (fs.statSync(path.join(dist, f)).size / 1048576).toFixed(2) + ' MB';
console.log(`Built ${VERSION}`);
console.log(`  PWA: ${files.length} files, app.js ${size('app.js')}`);
console.log(`  Offline single file: ${size('doc-to-md-offline.html')}`);
