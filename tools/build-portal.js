#!/usr/bin/env node
/**
 * Builds the standalone web portal (docs/index.html, served by GitHub Pages)
 * from the same ui/Index.html the Apps Script web app uses. The portal talks to
 * the script's JSON API (doPost) with fetch and no Google cookies.
 *
 *   node tools/build-portal.js <web app /exec URL>
 */
const fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..');
const url = process.argv[2] || (JSON.parse(fs.readFileSync(path.join(root, 'portal.json'), 'utf8')).apiUrl);
if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) throw new Error('Pass the web app /exec URL');
let html = fs.readFileSync(path.join(root, 'appsscript/ui/Index.html'), 'utf8');
const logo = fs.readFileSync(path.join(root, 'appsscript/ui/Logo.html'), 'utf8');
html = html
  .split("<?!= include('ui/Logo') ?>").join(logo)
  .split('<?= mode ?>').join('web')
  .replace(/<\?= hasPassword \? '([^']*)' :[\s\S]*?\?>/, '$1')
  .replace('<script>', '<script>window.ARETA_API_URL = ' + JSON.stringify(url) + ';</script>\n<script>');
if (/<\?/.test(html)) throw new Error('Unresolved template tag: ' + html.match(/<\?.{0,60}/)[0]);
html = html.replace('<head>', '<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<title>Areta Comparables Automation</title>\n<link rel="icon" type="image/svg+xml" href="assets/icon-brand.svg">\n<link rel="icon" type="image/png" sizes="32x32" href="assets/favicon-32.png">\n<link rel="apple-touch-icon" href="assets/favicon-180.png">');
fs.writeFileSync(path.join(root, 'docs/index.html'), html);
fs.writeFileSync(path.join(root, 'docs/.nojekyll'), '');
console.log('docs/index.html built for', url);
