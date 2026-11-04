/**
 * Tsavone Training Register – QR code and PNG maker.
 *
 * Makes QR codes entirely inside Apps Script (no outside services), so certificate
 * links never leave Google. Byte mode, error correction level M (about 15% damage
 * tolerance), versions 1–40, automatic mask choice. Based on the public QR Code
 * specification (ISO/IEC 18004).
 *
 * You do not need to change anything in this file.
 */

var QR_ECL_M_ = 1; // index into the tables below: 0 = L, 1 = M, 2 = Q, 3 = H
var QR_ECL_FORMAT_BITS_ = [1, 0, 3, 2];

var QR_ECC_CODEWORDS_PER_BLOCK_ = [
  [-1, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
  [-1, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28],
  [-1, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
  [-1, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30]
];
var QR_NUM_EC_BLOCKS_ = [
  [-1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25],
  [-1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49],
  [-1, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68],
  [-1, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81]
];

/** Returns a square array of booleans (true = dark module) encoding the text. */
function qrMatrix(text) {
  var data = utf8Bytes_(String(text));
  var ecl = QR_ECL_M_;
  var version, capacityBits, ccBits;
  for (version = 1; version <= 40; version++) {
    capacityBits = qrNumDataCodewords_(version, ecl) * 8;
    ccBits = version < 10 ? 8 : 16;
    if (4 + ccBits + data.length * 8 <= capacityBits) break;
  }
  if (version > 40) throw new Error('Text is too long for a QR code.');

  // Bit stream: byte mode, length, data, terminator, padding
  var bits = [];
  var push = function (val, len) { for (var i = len - 1; i >= 0; i--) bits.push((val >>> i) & 1); };
  push(4, 4);
  push(data.length, ccBits);
  data.forEach(function (b) { push(b, 8); });
  push(0, Math.min(4, capacityBits - bits.length));
  push(0, (8 - bits.length % 8) % 8);
  for (var pad = 0xEC; bits.length < capacityBits; pad ^= 0xEC ^ 0x11) push(pad, 8);
  var codewords = [];
  for (var i = 0; i < bits.length; i += 8) {
    var b = 0;
    for (var j = 0; j < 8; j++) b = (b << 1) | bits[i + j];
    codewords.push(b);
  }

  var qr = new QrBuilder_(version, ecl);
  qr.drawFunctionPatterns();
  qr.drawCodewords(qr.addEccAndInterleave(codewords));

  // Pick the mask with the lowest penalty
  var bestMask = 0, bestPenalty = Infinity;
  for (var m = 0; m < 8; m++) {
    qr.applyMask(m);
    qr.drawFormatBits(m);
    var p = qr.penalty();
    if (p < bestPenalty) { bestPenalty = p; bestMask = m; }
    qr.applyMask(m); // undo (XOR)
  }
  qr.applyMask(bestMask);
  qr.drawFormatBits(bestMask);
  return qr.modules;
}

function QrBuilder_(version, ecl) {
  this.version = version;
  this.ecl = ecl;
  this.size = version * 4 + 17;
  this.modules = [];
  this.isFunction = [];
  for (var y = 0; y < this.size; y++) {
    this.modules.push(new Array(this.size).fill(false));
    this.isFunction.push(new Array(this.size).fill(false));
  }
}

QrBuilder_.prototype.setFunction = function (x, y, dark) {
  this.modules[y][x] = dark;
  this.isFunction[y][x] = true;
};

QrBuilder_.prototype.drawFunctionPatterns = function () {
  var size = this.size, i, j;
  for (i = 0; i < size; i++) {
    this.setFunction(6, i, i % 2 === 0);
    this.setFunction(i, 6, i % 2 === 0);
  }
  this.drawFinder(3, 3);
  this.drawFinder(size - 4, 3);
  this.drawFinder(3, size - 4);
  var pos = qrAlignmentPositions_(this.version);
  var n = pos.length;
  for (i = 0; i < n; i++) {
    for (j = 0; j < n; j++) {
      if ((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)) continue;
      this.drawAlignment(pos[i], pos[j]);
    }
  }
  this.drawFormatBits(0);
  this.drawVersion();
};

QrBuilder_.prototype.drawFinder = function (x, y) {
  for (var dy = -4; dy <= 4; dy++) {
    for (var dx = -4; dx <= 4; dx++) {
      var dist = Math.max(Math.abs(dx), Math.abs(dy));
      var xx = x + dx, yy = y + dy;
      if (xx >= 0 && xx < this.size && yy >= 0 && yy < this.size) this.setFunction(xx, yy, dist !== 2 && dist !== 4);
    }
  }
};

QrBuilder_.prototype.drawAlignment = function (x, y) {
  for (var dy = -2; dy <= 2; dy++) {
    for (var dx = -2; dx <= 2; dx++) this.setFunction(x + dx, y + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
  }
};

QrBuilder_.prototype.drawFormatBits = function (mask) {
  var data = (QR_ECL_FORMAT_BITS_[this.ecl] << 3) | mask;
  var rem = data;
  for (var i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
  var bits = ((data << 10) | rem) ^ 0x5412;
  var bit = function (k) { return ((bits >>> k) & 1) !== 0; };
  var size = this.size;
  for (i = 0; i <= 5; i++) this.setFunction(8, i, bit(i));
  this.setFunction(8, 7, bit(6));
  this.setFunction(8, 8, bit(7));
  this.setFunction(7, 8, bit(8));
  for (i = 9; i < 15; i++) this.setFunction(14 - i, 8, bit(i));
  for (i = 0; i < 8; i++) this.setFunction(size - 1 - i, 8, bit(i));
  for (i = 8; i < 15; i++) this.setFunction(8, size - 15 + i, bit(i));
  this.setFunction(8, size - 8, true);
};

QrBuilder_.prototype.drawVersion = function () {
  if (this.version < 7) return;
  var rem = this.version;
  for (var i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >>> 11) * 0x1F25);
  var bits = (this.version << 12) | rem;
  for (i = 0; i < 18; i++) {
    var dark = ((bits >>> i) & 1) !== 0;
    var a = this.size - 11 + i % 3, b = Math.floor(i / 3);
    this.setFunction(a, b, dark);
    this.setFunction(b, a, dark);
  }
};

QrBuilder_.prototype.addEccAndInterleave = function (data) {
  var ver = this.version, ecl = this.ecl;
  var numBlocks = QR_NUM_EC_BLOCKS_[ecl][ver];
  var blockEccLen = QR_ECC_CODEWORDS_PER_BLOCK_[ecl][ver];
  var rawCodewords = Math.floor(qrNumRawDataModules_(ver) / 8);
  var numShortBlocks = numBlocks - rawCodewords % numBlocks;
  var shortBlockLen = Math.floor(rawCodewords / numBlocks);
  var divisor = rsDivisor_(blockEccLen);
  var blocks = [];
  for (var i = 0, k = 0; i < numBlocks; i++) {
    var len = shortBlockLen - blockEccLen + (i < numShortBlocks ? 0 : 1);
    var dat = data.slice(k, k + len);
    k += len;
    var ecc = rsRemainder_(dat, divisor);
    if (i < numShortBlocks) dat.push(0);
    blocks.push(dat.concat(ecc));
  }
  var result = [];
  for (i = 0; i < blocks[0].length; i++) {
    for (var j = 0; j < blocks.length; j++) {
      if (i !== shortBlockLen - blockEccLen || j >= numShortBlocks) result.push(blocks[j][i]);
    }
  }
  return result;
};

QrBuilder_.prototype.drawCodewords = function (data) {
  var i = 0, size = this.size;
  for (var right = size - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    for (var vert = 0; vert < size; vert++) {
      for (var j = 0; j < 2; j++) {
        var x = right - j;
        var upward = ((right + 1) & 2) === 0;
        var y = upward ? size - 1 - vert : vert;
        if (!this.isFunction[y][x] && i < data.length * 8) {
          this.modules[y][x] = ((data[i >>> 3] >>> (7 - (i & 7))) & 1) !== 0;
          i++;
        }
      }
    }
  }
};

QrBuilder_.prototype.applyMask = function (mask) {
  for (var y = 0; y < this.size; y++) {
    for (var x = 0; x < this.size; x++) {
      var invert;
      switch (mask) {
        case 0: invert = (x + y) % 2 === 0; break;
        case 1: invert = y % 2 === 0; break;
        case 2: invert = x % 3 === 0; break;
        case 3: invert = (x + y) % 3 === 0; break;
        case 4: invert = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0; break;
        case 5: invert = (x * y) % 2 + (x * y) % 3 === 0; break;
        case 6: invert = ((x * y) % 2 + (x * y) % 3) % 2 === 0; break;
        default: invert = ((x + y) % 2 + (x * y) % 3) % 2 === 0;
      }
      if (!this.isFunction[y][x] && invert) this.modules[y][x] = !this.modules[y][x];
    }
  }
};

/** Standard QR penalty score: lower is easier to scan. */
QrBuilder_.prototype.penalty = function () {
  var size = this.size, m = this.modules, result = 0, x, y, run, dark = 0;
  var lines = [];
  for (y = 0; y < size; y++) {
    var row = [], col = [];
    for (x = 0; x < size; x++) { row.push(m[y][x]); col.push(m[x][y]); }
    lines.push(row, col);
  }
  lines.forEach(function (line) {
    // Rule 1: runs of five or more of the same colour
    run = 1;
    for (var i = 1; i <= line.length; i++) {
      if (i < line.length && line[i] === line[i - 1]) run++;
      else { if (run >= 5) result += run - 2; run = 1; }
    }
    // Rule 3: finder-like patterns 1:1:3:1:1 with four light modules on one side
    for (i = 0; i + 6 < line.length; i++) {
      if (line[i] && !line[i + 1] && line[i + 2] && line[i + 3] && line[i + 4] && !line[i + 5] && line[i + 6]) {
        var before = true, after = true;
        for (var k = 1; k <= 4; k++) {
          if (i - k >= 0 && line[i - k]) before = false;
          if (i + 6 + k < line.length && line[i + 6 + k]) after = false;
        }
        if (before || after) result += 40;
      }
    }
  });
  // Rule 2: 2x2 blocks of one colour
  for (y = 0; y < size - 1; y++) {
    for (x = 0; x < size - 1; x++) {
      var c = m[y][x];
      if (c === m[y][x + 1] && c === m[y + 1][x] && c === m[y + 1][x + 1]) result += 3;
    }
  }
  // Rule 4: balance of dark and light
  for (y = 0; y < size; y++) for (x = 0; x < size; x++) if (m[y][x]) dark++;
  var total = size * size;
  result += Math.max(0, Math.ceil(Math.abs(dark * 20 - total * 10) / total) - 1) * 10;
  return result;
};

function qrAlignmentPositions_(ver) {
  if (ver === 1) return [];
  var num = Math.floor(ver / 7) + 2;
  var step = ver === 32 ? 26 : Math.ceil((ver * 4 + 4) / (num * 2 - 2)) * 2;
  var result = [6];
  for (var pos = ver * 4 + 17 - 7; result.length < num; pos -= step) result.splice(1, 0, pos);
  return result;
}

function qrNumRawDataModules_(ver) {
  var result = (16 * ver + 128) * ver + 64;
  if (ver >= 2) {
    var numAlign = Math.floor(ver / 7) + 2;
    result -= (25 * numAlign - 10) * numAlign - 55;
    if (ver >= 7) result -= 36;
  }
  return result;
}

function qrNumDataCodewords_(ver, ecl) {
  return Math.floor(qrNumRawDataModules_(ver) / 8) - QR_ECC_CODEWORDS_PER_BLOCK_[ecl][ver] * QR_NUM_EC_BLOCKS_[ecl][ver];
}

function rsMultiply_(x, y) {
  var z = 0;
  for (var i = 7; i >= 0; i--) {
    z = (z << 1) ^ ((z >>> 7) * 0x11D);
    z ^= ((y >>> i) & 1) * x;
  }
  return z;
}

function rsDivisor_(degree) {
  var result = [];
  for (var i = 0; i < degree - 1; i++) result.push(0);
  result.push(1);
  var root = 1;
  for (i = 0; i < degree; i++) {
    for (var j = 0; j < result.length; j++) {
      result[j] = rsMultiply_(result[j], root);
      if (j + 1 < result.length) result[j] ^= result[j + 1];
    }
    root = rsMultiply_(root, 0x02);
  }
  return result;
}

function rsRemainder_(data, divisor) {
  var result = divisor.map(function () { return 0; });
  data.forEach(function (b) {
    var factor = b ^ result.shift();
    result.push(0);
    divisor.forEach(function (coef, i) { result[i] ^= rsMultiply_(coef, factor); });
  });
  return result;
}

function utf8Bytes_(str) {
  var out = [];
  for (var i = 0; i < str.length; i++) {
    var c = str.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF && i + 1 < str.length) {
      c = 0x10000 + ((c - 0xD800) << 10) + (str.charCodeAt(++i) - 0xDC00);
    }
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xC0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xF0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return out;
}

/* ---------- PNG maker (greyscale, no compression needed) ---------- */

/**
 * Turns a QR matrix into PNG bytes (numbers 0–255).
 * scale = pixels per module; a 4-module white border is added, as the QR standard requires.
 */
function qrPngBytes(matrix, scale) {
  scale = scale || 8;
  var quiet = 4;
  var n = matrix.length;
  var px = (n + quiet * 2) * scale;
  var raw = [];
  for (var y = 0; y < px; y++) {
    raw.push(0); // filter type: none
    var my = Math.floor(y / scale) - quiet;
    for (var x = 0; x < px; x++) {
      var mx = Math.floor(x / scale) - quiet;
      var dark = my >= 0 && my < n && mx >= 0 && mx < n && matrix[my][mx];
      raw.push(dark ? 0 : 255);
    }
  }
  var ihdr = [].concat(u32_(px), u32_(px), [8, 0, 0, 0, 0]); // 8-bit greyscale
  return [137, 80, 78, 71, 13, 10, 26, 10]
    .concat(pngChunk_('IHDR', ihdr))
    .concat(pngChunk_('IDAT', zlibStore_(raw)))
    .concat(pngChunk_('IEND', []));
}

/** Returns an Apps Script PNG blob of a QR code for the text. */
function qrPngBlob(text, name) {
  var bytes = qrPngBytes(qrMatrix(text), 8).map(function (b) { return b > 127 ? b - 256 : b; });
  return Utilities.newBlob(bytes, 'image/png', name || 'qr.png');
}

function u32_(v) { return [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255]; }

function pngChunk_(type, data) {
  var typeBytes = type.split('').map(function (ch) { return ch.charCodeAt(0); });
  var crcInput = typeBytes.concat(data);
  return [].concat(u32_(data.length), typeBytes, data, u32_(crc32_(crcInput)));
}

var CRC_TABLE_ = null;
function crc32_(bytes) {
  if (!CRC_TABLE_) {
    CRC_TABLE_ = [];
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE_.push(c >>> 0);
    }
  }
  var crc = 0xFFFFFFFF;
  for (var i = 0; i < bytes.length; i++) crc = CRC_TABLE_[(crc ^ bytes[i]) & 255] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

/** zlib wrapper using uncompressed ("stored") blocks. */
function zlibStore_(data) {
  var out = [0x78, 0x01];
  for (var i = 0; i < data.length || i === 0; i += 65535) {
    var chunk = data.slice(i, i + 65535);
    var last = i + 65535 >= data.length ? 1 : 0;
    var len = chunk.length;
    out.push(last, len & 255, (len >>> 8) & 255, ~len & 255, (~len >>> 8) & 255);
    for (var j = 0; j < chunk.length; j++) out.push(chunk[j]);
    if (last) break;
  }
  var a = 1, b = 0;
  for (i = 0; i < data.length; i++) { a = (a + data[i]) % 65521; b = (b + a) % 65521; }
  return out.concat(u32_(((b << 16) | a) >>> 0));
}
