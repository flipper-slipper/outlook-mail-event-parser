/**
 * Generates calendar-shaped PNG icons at all sizes required by the manifest.
 */
const fs   = require("fs");
const path = require("path");
const zlib = require("zlib");

const assetsDir = path.join(__dirname, "..", "assets");
if (!fs.existsSync(assetsDir)) fs.mkdirSync(assetsDir);

// ── Palette ───────────────────────────────────────────────────────────────────
const BLUE   = [0,  120, 212];
const DBLU   = [0,   84, 153];  // darker border / ring outline
const WHITE  = [255, 255, 255];
const LGRID  = [200, 225, 245];  // light-blue grid lines

// ── Pixel canvas ──────────────────────────────────────────────────────────────
function makeCanvas(size) {
  const data = new Array(size).fill(null).map(() =>
    new Array(size).fill(null).map(() => [...WHITE])
  );
  const set = (x, y, c) => {
    if (x >= 0 && x < size && y >= 0 && y < size) data[y][x] = c;
  };
  const fill = (x0, y0, x1, y1, c) => {
    for (let y = y0; y <= y1; y++)
      for (let x = x0; x <= x1; x++) set(x, y, c);
  };
  const hline = (y, x0, x1, c) => { for (let x = x0; x <= x1; x++) set(x, y, c); };
  const vline = (x, y0, y1, c) => { for (let y = y0; y <= y1; y++) set(x, y, c); };
  return { data, set, fill, hline, vline };
}

// ── Calendar icon renderer ────────────────────────────────────────────────────
function drawCalendar(size) {
  const { data, set, fill, hline, vline } = makeCanvas(size);

  const s   = (v) => Math.round(size * v);   // scale helper
  const pad  = Math.max(1, s(0.06));          // outer padding
  const bord = Math.max(1, s(0.05));          // border thickness
  const hdrH = Math.max(3, s(0.28));          // header height

  const left  = pad;
  const right = size - 1 - pad;
  const top   = pad;
  const bot   = size - 1 - pad;

  // ── White background (already filled) ──

  // ── Blue header strip ────────────────────────────────────────────────────────
  fill(left, top, right, top + hdrH, BLUE);

  // ── White calendar body ──────────────────────────────────────────────────────
  fill(left + bord, top + hdrH + 1, right - bord, bot - bord, WHITE);

  // ── Blue border ──────────────────────────────────────────────────────────────
  for (let b = 0; b < bord; b++) {
    hline(top + b,       left, right, DBLU);
    hline(bot - b,       left, right, DBLU);
    vline(left  + b,     top,  bot,   DBLU);
    vline(right - b,     top,  bot,   DBLU);
  }

  // ── Calendar binding rings (white tabs above header) ─────────────────────────
  if (size >= 16) {
    const rw  = Math.max(1, s(0.07));   // half-width of ring
    const rh  = Math.max(2, s(0.18));   // ring height (extends above icon top)
    const r1x = s(0.30);
    const r2x = s(0.70);
    const rTop = 0;
    const rBot = Math.min(top + Math.round(hdrH * 0.55), size - 1);

    for (const rx of [r1x, r2x]) {
      // white fill
      fill(rx - rw, rTop, rx + rw, rBot, WHITE);
      // dark outline
      for (let y = rTop; y <= rBot; y++) {
        set(rx - rw - 1, y, DBLU);
        set(rx + rw + 1, y, DBLU);
      }
      hline(rBot + 1, rx - rw - 1, rx + rw + 1, DBLU);
    }
  }

  // ── Grid lines in body (only for larger icons) ────────────────────────────────
  if (size >= 32) {
    const bodyL = left  + bord;
    const bodyR = right - bord;
    const bodyT = top   + hdrH + 1;
    const bodyB = bot   - bord;

    const cols = 7;
    const rows = size >= 64 ? 5 : 4;
    const cellW = (bodyR - bodyL) / cols;
    const cellH = (bodyB - bodyT) / rows;

    for (let c = 1; c < cols; c++) {
      const x = Math.round(bodyL + c * cellW);
      vline(x, bodyT, bodyB, LGRID);
    }
    for (let r = 1; r < rows; r++) {
      const y = Math.round(bodyT + r * cellH);
      hline(y, bodyL, bodyR, LGRID);
    }

    // ── Highlight one "today" cell ────────────────────────────────────────────
    if (size >= 48) {
      const cellX = Math.round(bodyL + 1 * cellW) + 1;
      const cellY = Math.round(bodyT + 0 * cellH) + 1;
      const cellXe = Math.round(bodyL + 2 * cellW) - 1;
      const cellYe = Math.round(bodyT + 1 * cellH) - 1;
      fill(cellX, cellY, cellXe, cellYe, [220, 240, 255]);
    }
  }

  return data;
}

// ── PNG encoder ───────────────────────────────────────────────────────────────
function toPng(pixelData, size) {
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(size, 0);
  ihdrData.writeUInt32BE(size, 4);
  ihdrData[8] = 8;  // bit depth
  ihdrData[9] = 2;  // RGB
  const ihdr = pngChunk("IHDR", ihdrData);

  // Build raw image bytes (filter 0 per row)
  const rows = [];
  for (let y = 0; y < size; y++) {
    const row = Buffer.alloc(1 + size * 3);
    row[0] = 0;  // filter None
    for (let x = 0; x < size; x++) {
      row[1 + x * 3]     = pixelData[y][x][0];
      row[1 + x * 3 + 1] = pixelData[y][x][1];
      row[1 + x * 3 + 2] = pixelData[y][x][2];
    }
    rows.push(row);
  }
  const idat = pngChunk("IDAT", zlib.deflateSync(Buffer.concat(rows)));
  const iend = pngChunk("IEND", Buffer.alloc(0));

  return Buffer.concat([sig, ihdr, idat, iend]);
}

function pngChunk(type, data) {
  const len  = Buffer.alloc(4); len.writeUInt32BE(data.length, 0);
  const tBuf = Buffer.from(type, "ascii");
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([tBuf, data])) >>> 0, 0);
  return Buffer.concat([len, tBuf, data, crcBuf]);
}

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let j = 0; j < 8; j++) c = (c & 1) ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
  }
  return c ^ 0xffffffff;
}

// ── Generate all sizes ────────────────────────────────────────────────────────
const sizes = [16, 32, 64, 80, 128];
for (const size of sizes) {
  const pixels = drawCalendar(size);
  const png    = toPng(pixels, size);
  const out    = path.join(assetsDir, `icon-${size}.png`);
  fs.writeFileSync(out, png);
  console.log(`Created ${out} (${size}x${size})`);
}
console.log("\nAll calendar icons created in assets/");
