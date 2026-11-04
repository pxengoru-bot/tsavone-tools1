// Checks that Qr.gs makes QR codes that a real decoder (jsQR) reads back correctly,
// and that the PNG it builds is a valid image.
// Run: NODE_PATH=<folder containing node_modules with jsqr> node register/tests/qr.test.js
const fs = require('fs');
const vm = require('vm');
const zlib = require('zlib');
const path = require('path');
const jsQR = require('jsqr');

const ctx = {};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Qr.gs'), 'utf8'), ctx);

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) pass++; else { fail++; console.log('FAIL', name, extra || ''); }
}

function decodeMatrix(matrix, scale = 4) {
  const quiet = 4, n = matrix.length, px = (n + quiet * 2) * scale;
  const rgba = new Uint8ClampedArray(px * px * 4);
  for (let y = 0; y < px; y++) for (let x = 0; x < px; x++) {
    const my = Math.floor(y / scale) - quiet, mx = Math.floor(x / scale) - quiet;
    const dark = my >= 0 && my < n && mx >= 0 && mx < n && matrix[my][mx];
    const v = dark ? 0 : 255, i = (y * px + x) * 4;
    rgba[i] = rgba[i + 1] = rgba[i + 2] = v; rgba[i + 3] = 255;
  }
  const r = jsQR(rgba, px, px, { inversionAttempts: 'dontInvert' });
  return r ? r.data : null;
}

// 1. Realistic verification links of many lengths (covers QR versions 1 to about 15)
const base = 'https://script.google.com/macros/s/AKfycbx' + 'Q'.repeat(65) + '/exec?id=';
const samples = [
  'TSV-2026-T1.1-0001',
  base + 'TSV-2026-T1.1-0001',
  base + 'TSV-2027-T8.5-0123',
  'https://www.tsavone.com/verify?id=TSV-2026-T3.3-0042',
  'A', 'Tsavone', 'https://www.tsavone.com',
  'Café Ntinda – Gulu City ✓'
];
for (let len = 1; len <= 330; len += 7) samples.push('x'.repeat(len).replace(/x/g, (_, i) => 'abcdefghijklmnopqrstuvwxyz0123456789-./?=&'[(i * 7 + len) % 42]));
const versions = new Set();
samples.forEach((s) => {
  const m = ctx.qrMatrix(s);
  versions.add((m.length - 17) / 4);
  const got = decodeMatrix(m);
  check('decode ' + s.slice(0, 40) + ' (len ' + s.length + ')', got === s, 'got ' + JSON.stringify(got && got.slice(0, 40)));
});

// 2. PNG structure: signature, chunk CRCs, pixel data inflates to the right size and decodes
const text = base + 'TSV-2026-T1.1-0001';
const png = Buffer.from(ctx.qrPngBytes(ctx.qrMatrix(text), 8));
check('png signature', png.slice(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])));
let off = 8, width = 0, idat = [];
const crcTable = [];
for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; crcTable.push(c >>> 0); }
const crc = (b) => { let c = 0xFFFFFFFF; for (const x of b) c = crcTable[(c ^ x) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
while (off < png.length) {
  const len = png.readUInt32BE(off), type = png.slice(off + 4, off + 8).toString();
  const data = png.slice(off + 8, off + 8 + len);
  check('crc ' + type, crc(png.slice(off + 4, off + 8 + len)) === png.readUInt32BE(off + 8 + len));
  if (type === 'IHDR') width = data.readUInt32BE(0);
  if (type === 'IDAT') idat.push(data);
  off += 12 + len;
}
const raw = zlib.inflateSync(Buffer.concat(idat)); // also verifies the Adler-32 checksum
check('png raw size', raw.length === width * (width + 1), raw.length + ' vs ' + width * (width + 1));
const rgba = new Uint8ClampedArray(width * width * 4);
for (let y = 0; y < width; y++) for (let x = 0; x < width; x++) {
  const v = raw[y * (width + 1) + 1 + x], i = (y * width + x) * 4;
  rgba[i] = rgba[i + 1] = rgba[i + 2] = v; rgba[i + 3] = 255;
}
const fromPng = jsQR(rgba, width, width);
check('png decodes', fromPng && fromPng.data === text);
fs.writeFileSync(path.join(require('os').tmpdir(), 'tsavone-qr-test.png'), png);

console.log(`QR versions exercised: ${[...versions].sort((a, b) => a - b).join(', ')}`);
console.log(`${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
