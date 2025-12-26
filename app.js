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

function loadSizeMapFromSheet(ws) {
  const map = new Map(); // size -> { row, col }
  if (!ws || !ws["!ref"]) return map;

  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell || cell.v == null) continue;

      const s = normalizeSize(cell.v);
      if (!SIZES.includes(s)) continue;

      // Guardamos la primera aparición de cada talla (normalmente es la columna de tallas)
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

function applyBlock(ws, prefix) {
  const sizeMap = loadSizeMapFromSheet(ws);

  // Rellena en la columna de la derecha de donde esté la talla
  for (const size of SIZES) {
    const pos = sizeMap.get(size);
    if (!pos) continue;

    const key = `${prefix}_${size.toLowerCase()}`; // ej: cam_xs
    const val = n(key);
    setNumericCell(ws, pos.r, pos.c + 1, val);
  }
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
  if (wsCam) applyBlock(wsCam, "cam");

  const wsSud = wb.Sheets["SUDADERA"];
  if (wsSud) {
    // Para diferenciar técnica y paseo, buscamos tallas y rellenamos todas;
    // si tu hoja tiene 2 bloques con tallas repetidas, la función por defecto
    // rellenará la primera aparición de cada talla.
    // Para hacerlo 100% exacto con 2 bloques, usamos un método por “anclaje”:
    applySudaderaTwoBlocks(wsSud);
  }

  const wsPch = wb.Sheets["PANTALON CHANDAL"];
  if (wsPch) applyBlock(wsPch, "pch");

  downloadWorkbook(wb, buildFilename());
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
      const s = String(cell.v).toLowerCase();
      if (s.includes(t)) return r;
    }
  }
  return null;
}

function loadSizeMapFromSheetStartingAt(ws, startRow) {
  const map = new Map();
  if (!ws || !ws["!ref"]) return map;

  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let r = Math.max(startRow, range.s.r); r <= range.e.r; r++) {
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

function applyBlockFrom(ws, prefix, startRow) {
  const sizeMap = loadSizeMapFromSheetStartingAt(ws, startRow);
  for (const size of SIZES) {
    const pos = sizeMap.get(size);
    if (!pos) continue;
    const key = `${prefix}_${size.toLowerCase()}`;
    const val = n(key);
    setNumericCell(ws, pos.r, pos.c + 1, val);
  }
}

function applySudaderaTwoBlocks(ws) {
  const rowTec = findRowContaining(ws, "SUDADERA TECNICA");
  const rowPas = findRowContaining(ws, "SUDADERA PASEO");

  // Si no encontramos títulos, caemos al modo simple
  if (rowTec == null || rowPas == null) {
    applyBlock(ws, "sudt");
    return;
  }

  // Rellenamos desde justo debajo del título
  applyBlockFrom(ws, "sudt", rowTec + 1);
  applyBlockFrom(ws, "sudp", rowPas + 1);
}

document.getElementById("btn")?.addEventListener("click", async () => {
  try {
    await generate();
  } catch (err) {
    alert(err?.message || String(err));
  }
});

document.getElementById("btnClear")?.addEventListener("click", () => location.reload());
