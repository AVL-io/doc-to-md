// Renders the PNG app icons from src/icons/icon.svg (run once; output is committed).
import { chromium } from './playwright.mjs';
import fs from 'node:fs';
const svg = fs.readFileSync('src/icons/icon.svg', 'utf8');
const browser = await chromium.launch();
const page = await browser.newPage({ deviceScaleFactor: 1 });
for (const [name, size, pad] of [['icon-192.png', 192, 0], ['icon-512.png', 512, 0], ['icon-maskable-512.png', 512, 0.12]]) {
  const inner = Math.round(size * (1 - pad * 2));
  await page.setViewportSize({ width: size, height: size });
  await page.setContent(`<html><body style="margin:0;background:#1f4e79;display:flex;align-items:center;justify-content:center;width:${size}px;height:${size}px">${svg.replace('<svg ', `<svg width="${inner}" height="${inner}" `)}</body></html>`);
  fs.writeFileSync(`src/icons/${name}`, await page.screenshot({ omitBackground: pad === 0, clip: { x: 0, y: 0, width: size, height: size } }));
}
await browser.close();
console.log('icons written');
