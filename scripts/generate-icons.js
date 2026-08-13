#!/usr/bin/env node
/**
 * Generates the add-in icons (16/32/80 px PNGs) with zero dependencies.
 * Simple geometric design: army-green rounded block with a lighter band.
 */
"use strict";

const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const OUT = path.resolve(__dirname, "..", "assets");

// --- CRC32 (PNG chunk checksums) ---
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, "ascii");
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function makePng(size) {
  const w = size, h = size;
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 2;  // color type: RGB
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  const raw = Buffer.alloc(h * (1 + w * 3));
  for (let y = 0; y < h; y++) {
    raw[y * (1 + w * 3)] = 0; // filter: none
    for (let x = 0; x < w; x++) {
      const i = y * (1 + w * 3) + 1 + x * 3;
      const edge = Math.max(1, Math.round(size * 0.06)); // border thickness
      const inBorder = x < edge || y < edge || x >= w - edge || y >= h - edge;
      // corner rounding
      const corner =
        (x < edge * 2 && y < edge * 2) || (x >= w - edge * 2 && y < edge * 2) ||
        (x < edge * 2 && y >= h - edge * 2) || (x >= w - edge * 2 && y >= h - edge * 2);
      // horizontal band through the middle
      const band = y >= h * 0.38 && y <= h * 0.62;
      let r, g, b;
      if (corner) {
        r = 0x1d; g = 0x3a; b = 0x26; // dark corner (rounded feel)
      } else if (band) {
        r = 0xe8; g = 0xef; b = 0xe9; // light band
      } else if (inBorder) {
        r = 0x17; g = 0x2f; b = 0x1f; // border
      } else {
        r = 0x2f; g = 0x5d; b = 0x3a; // army green
      }
      raw[i] = r; raw[i + 1] = g; raw[i + 2] = b;
    }
  }

  const idat = zlib.deflateSync(raw, { level: 9 });
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", idat),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

fs.mkdirSync(OUT, { recursive: true });
for (const size of [16, 32, 80]) {
  const png = makePng(size);
  const file = path.join(OUT, `icon-${size}.png`);
  fs.writeFileSync(file, png);
  // sanity check: PNG signature + IHDR width/height
  const ok = png.length > 24 && png.readUInt32BE(16) === size && png.readUInt32BE(20) === size;
  console.log(`${path.basename(file)} ${size}x${size} (${png.length} bytes) ${ok ? "ok" : "SANITY FAIL"}`);
}
