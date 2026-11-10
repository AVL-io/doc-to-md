// Loads Playwright from the project or from the global npm folder (dev/test tooling only).
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
const require = createRequire(import.meta.url);
let pw;
try { pw = require('playwright'); }
catch {
  const g = execSync('npm root -g').toString().trim();
  pw = require(require.resolve('playwright', { paths: [g] }));
}
export const { chromium } = pw;
