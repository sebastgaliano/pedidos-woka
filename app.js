const SIZES = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"];

function n(id) {
  const el = document.getElementById(id);
  if (!el) return 0;
  const v = Number(el.value);
  return Number.isFinite(v) && v >= 0 ? v : 0;
}

function normalizeSize(v) {
  return String(v ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/^XXS$/, "XS");
}

function findRowContaining(ws, text) {
  if (!ws || !ws["!ref"]) return null;
  const t = String(text).toLowerCase();
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell || cell.v == null) continue;
      if (String(cell.v).toLowerCase().includes(t)) return r;
    }
  }
  return null;
}

function loadSizeMapFromSheetStartingAt(ws, startRow) {
  const map = new Map();
  if (!ws || !ws["!ref"]) return map;

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const r0 = Math.max(startRow, range.s.r);

  for (let r = r0; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell || cell.v == null) continue;

      const s = normalizeSize(cell.v);
      if (!SIZES.includes(s)) continue;

      if (!map.has(s)) map.set(s, { r, c });
    }
  }

  return map;
}

function setNumericCell(ws, r, c, value) {
  const addr = XLSX.utils.encode_cell({ r, c });
  ws[addr] = ws[addr] || {};
  ws[addr].v = value;
  ws[addr].t = "n";
}

function applyBlockFrom(ws, prefix, startRow) {
  const sizeMap = loadSizeMapFromSheetStartingAt(ws, startRow);
  for (const size of SIZES) {
    const pos = sizeMap.get(size);
    if (!pos) continue;
    const key = `${prefix}_${size.toLowerCase()}`;
    setNumericCell(ws, pos.r, pos.c + 1, n(key));
  }
}

function applySudaderaTwoBlocks(ws) {
  const rowTec = findRowContaining(ws, "SUDADERA TECNICA");
  const rowPas = findRowContaining(ws, "SUDADERA PASEO");

  if (rowTec == null || rowPas == null) {
    applyBlockFrom(ws, "sudt", 0);
    return;
  }

  applyBlockFrom(ws, "sudt", rowTec + 1);
  applyBlockFrom(ws, "sudp", rowPas + 1);
}

async function loadTemplate() {
  const res = await fetch("./Hoja de pedido_actualizada.xlsx", { cache: "no-store" });
  if (!res.ok) throw new Error("No se pudo cargar la plantilla.");
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type: "array" });
}

function buildFilename() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `PEDIDO_WOKA_${yyyy}-${mm}-${dd}.xlsx`;
}

function downloadWorkbook(wb, filename) {
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();

  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

async function generate() {
  const wb = await loadTemplate();

  const wsCam = wb.Sheets["CAMISETA"];
  if (wsCam) applyBlockFrom(wsCam, "cam", 0);

  const wsSud = wb.Sheets["SUDADERA"];
  if (wsSud) applySudaderaTwoBlocks(wsSud);

  const wsPch = wb.Sheets["PANTALON CHANDAL"];
  if (wsPch) applyBlockFrom(wsPch, "pch", 0);

  downloadWorkbook(wb, buildFilename());
}

function bind() {
  const btn = document.getElementById("btn");
  const btnMobile = document.getElementById("btnMobile");
  const btnClear = document.getElementById("btnClear");
  const btnClearMobile = document.getElementById("btnClearMobile");

  btn?.addEventListener("click", async () => {
    try { await generate(); } catch (e) { alert(e?.message || String(e)); }
  });

  btnMobile?.addEventListener("click", () => btn?.click());

  const clear = () => location.reload();
  btnClear?.addEventListener("click", clear);
  btnClearMobile?.addEventListener("click", clear);
}

bind();
