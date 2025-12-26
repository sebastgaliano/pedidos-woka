const CONFIG = {
  templatePath: "./Hoja de pedido_actualizada.xlsx",
  logoPath: "./Logo.png",
  sizes: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"],
  products: [
    { key: "camiseta", title: "Camiseta", sheet: "CAMISETA", sections: [{ key: "base", title: null, anchorText: null }] },
    { key: "sudadera", title: "Sudadera", sheet: "SUDADERA", sections: [
      { key: "tecnica", title: "Sudadera técnica", anchorText: "SUDADERA TECNICA" },
      { key: "paseo", title: "Sudadera paseo", anchorText: "SUDADERA PASEO" },
    ]},
    { key: "pantalon_chandal", title: "Pantalón chándal", sheet: "PANTALON CHANDAL", sections: [{ key: "base", title: null, anchorText: null }] },

    // Ejemplos futuros (solo alta aquí, sin tocar HTML):
    // { key: "equipacion_juego", title: "Equipación de juego", sheet: "EQUIPACION JUEGO", sections: [{ key:"base", title:null, anchorText:null }] },
    // { key: "cubre_juego", title: "Cubre de juego", sheet: "CUBRE JUEGO", sections: [{ key:"base", title:null, anchorText:null }] },
    // { key: "pantalon_juego", title: "Pantalón de juego", sheet: "PANTALON JUEGO", sections: [{ key:"base", title:null, anchorText:null }] },
  ],
};
function shiftMerges(ws, fromRow, delta) {
  const merges = ws["!merges"];
  if (!Array.isArray(merges) || delta === 0) return;

  for (const m of merges) {
    if (m.s.r >= fromRow) m.s.r += delta;
    if (m.e.r >= fromRow) m.e.r += delta;
  }
}
function findHeaderRow(ws, startRow, endRow) {
  if (!ws || !ws["!ref"]) return null;

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const rStart = Math.max(startRow, range.s.r);
  const rEnd = Math.min(endRow, range.e.r);

  const required = ["nombre", "talla", "cantidad"]; // mínimo fiable

  for (let r = rStart; r <= rEnd; r++) {
    let hit = 0;

    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;

      const v = String(cell.v).toLowerCase().trim();
      if (required.some(k => v === k)) hit++;
    }

    if (hit >= 2) return r; // con 2/3 ya lo damos por bueno
  }

  return null;
}


function clampMergesToRef(ws) {
  const merges = ws["!merges"];
  if (!Array.isArray(merges) || !ws["!ref"]) return;

  const range = XLSX.utils.decode_range(ws["!ref"]);
  ws["!merges"] = merges.filter(m =>
    m.s.r >= range.s.r && m.e.r <= range.e.r &&
    m.s.c >= range.s.c && m.e.c <= range.e.c
  );
}

function autoFitColumns(ws, minWch = 8, maxWch = 70) {
  if (!ws || !ws["!ref"]) return;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const cols = new Array(range.e.c + 1).fill(0);

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell || cell.v == null) continue;

      const s = String(cell.v);
      // penaliza menos números, más texto largo
      const len = s.length;
      cols[c] = Math.max(cols[c], len);
    }
  }

  ws["!cols"] = cols.map(len => {
    const wch = Math.max(minWch, Math.min(maxWch, Math.ceil(len * 1.1)));
    return { wch };
  });
}

function idFor(productKey, sectionKey, size) {
  return `${productKey}__${sectionKey}__${String(size).toLowerCase()}`;
}

function n(id) {
  const el = document.getElementById(id);
  if (!el) return 0;
  const v = Number(el.value);
  return Number.isFinite(v) && v >= 0 ? v : 0;
}

function normalizeSize(v) {
  return String(v ?? "").trim().toUpperCase().replace(/\s+/g, "");
}

function render() {
  const root = document.getElementById("app");
  root.innerHTML = "";

  for (const p of CONFIG.products) {
    const card = document.createElement("div");
    card.className = "card";

    const head = document.createElement("div");
    head.className = "card-head";

    const h2 = document.createElement("h2");
    h2.textContent = p.title;

    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = `Hoja: ${p.sheet}`;

    head.appendChild(h2);
    head.appendChild(chip);
    card.appendChild(head);

    for (const s of p.sections) {
      if (s.title) {
        const sub = document.createElement("div");
        sub.className = "subhead";
        sub.textContent = s.title;
        card.appendChild(sub);
      }

      const sizes = document.createElement("div");
      sizes.className = "sizes";

      for (const size of CONFIG.sizes) {
        const wrap = document.createElement("div");

        const label = document.createElement("label");
        label.textContent = size;

        const input = document.createElement("input");
        input.type = "number";
        input.min = "0";
        input.value = "0";
        input.inputMode = "numeric";
        input.id = idFor(p.key, s.key, size);

        wrap.appendChild(label);
        wrap.appendChild(input);
        sizes.appendChild(wrap);
      }

      card.appendChild(sizes);
    }

    root.appendChild(card);
  }
}

function buildFilename() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `PEDIDO_WOKA_${yyyy}-${mm}-${dd}.xlsx`;
}

async function loadTemplate() {
  const res = await fetch(CONFIG.templatePath, { cache: "no-store" });
  if (!res.ok) throw new Error("No se pudo cargar la plantilla.");
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, {
    type: "array",
    cellStyles: true
  });
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

function decodeRef(ws) {
  if (!ws["!ref"]) ws["!ref"] = "A1:A1";
  return XLSX.utils.decode_range(ws["!ref"]);
}

function setRef(ws, range) {
  ws["!ref"] = XLSX.utils.encode_range(range);
}

function cellAddr(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}

function getCell(ws, r, c) {
  return ws[cellAddr(r, c)];
}

function setCell(ws, r, c, cellObj) {
  ws[cellAddr(r, c)] = cellObj;
}

function deleteCell(ws, r, c) {
  delete ws[cellAddr(r, c)];
}

function findRowContaining(ws, text) {
  const range = decodeRef(ws);
  const t = String(text).toLowerCase();
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = getCell(ws, r, c);
      if (!cell || cell.v == null) continue;
      if (String(cell.v).toLowerCase().includes(t)) return r;
    }
  }
  return null;
}

function findSizeColumnAndRows(ws, startRow, endRow, sizesSet) {
  const range = decodeRef(ws);
  const rStart = Math.max(startRow, range.s.r);
  const rEnd = Math.min(endRow, range.e.r);

  for (let c = range.s.c; c <= range.e.c; c++) {
    const rows = [];
    for (let r = rStart; r <= rEnd; r++) {
      const cell = getCell(ws, r, c);
      if (!cell || cell.v == null) continue;
      const v = normalizeSize(cell.v);
      if (sizesSet.has(v)) rows.push(r);
    }
    if (rows.length) return { sizeCol: c, rows };
  }
  return null;
}

function shiftRows(ws, fromRow, delta) {
  if (delta === 0) return;

  const range = decodeRef(ws);
  const newRange = { s: { ...range.s }, e: { ...range.e } };

  const cells = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = cellAddr(r, c);
      const cell = ws[addr];
      if (!cell) continue;
      cells.push({ r, c, cell });
    }
  }

  for (const { r, c } of cells) deleteCell(ws, r, c);

  for (const { r, c, cell } of cells) {
    const nr = (r >= fromRow) ? r + delta : r;
    setCell(ws, nr, c, cell);
  }

  // Ajusta merges (celdas combinadas) para que no se rompa el formato
  shiftMerges(ws, fromRow, delta);

  // Actualiza rango
  if (delta > 0) {
    if (range.e.r >= fromRow) newRange.e.r = range.e.r + delta;
  } else {
    if (range.e.r >= fromRow) newRange.e.r = range.e.r + delta;
  }

  setRef(ws, newRange);
  clampMergesToRef(ws);
}


function cloneCell(cell) {
  if (!cell) return undefined;
  return JSON.parse(JSON.stringify(cell));
}

function cloneRow(ws, srcRow, dstRow) {
  const range = decodeRef(ws);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const src = getCell(ws, srcRow, c);
    if (src) setCell(ws, dstRow, c, cloneCell(src));
    else deleteCell(ws, dstRow, c);
  }
}

function clearRow(ws, row) {
  const range = decodeRef(ws);
  for (let c = range.s.c; c <= range.e.c; c++) deleteCell(ws, row, c);
}

function setNumber(ws, r, c, value) {
  const addr = cellAddr(r, c);
  ws[addr] = ws[addr] || {};
  ws[addr].v = value;
  ws[addr].t = "n";
}

function setText(ws, r, c, value) {
  const addr = cellAddr(r, c);
  ws[addr] = ws[addr] || {};
  ws[addr].v = value;
  ws[addr].t = "s";
}

function applySection(ws, productKey, sectionKey, anchorText) {
  const sizesSet = new Set(CONFIG.sizes.map(s => normalizeSize(s)));

  const qtyBySize = new Map();
  for (const size of CONFIG.sizes) {
    const val = n(idFor(productKey, sectionKey, size));
    if (val > 0) qtyBySize.set(normalizeSize(size), val);
  }

  const range = decodeRef(ws);

  let startRow = range.s.r;
  let endRow = range.e.r;

  if (anchorText) {
    const anchorRow = findRowContaining(ws, anchorText);
    if (anchorRow != null) {
      startRow = anchorRow + 1;
      endRow = range.e.r;
    }
  }

  const found = findSizeColumnAndRows(ws, startRow, endRow, sizesSet);

  let sizeCol = null;
  let existingRows = [];
  let templateRow = null;

  if (found) {
    sizeCol = found.sizeCol;
    existingRows = found.rows.slice().sort((a,b)=>a-b);
    templateRow = existingRows[0];
  } else {
    // Si no hay tallas en la plantilla para esa sección:
    // intentamos deducir una fila "plantilla" cercana (la primera fila después del anchor)
    sizeCol = null;
    templateRow = startRow;
    existingRows = [];
  }

  // Si no hay ninguna talla con cantidad, eliminamos todas las filas de tallas existentes y listo.
  if (qtyBySize.size === 0) {
    if (existingRows.length) {
      // borrar todas las filas de tallas detectadas
      for (let i = existingRows.length - 1; i >= 0; i--) {
        const row = existingRows[i];
        // borrar esa fila desplazando arriba
        clearRow(ws, row);
        shiftRows(ws, row + 1, -1);
      }
    }
    return;
  }

  // 1) Borramos filas existentes de tallas (todas) para reconstruir limpio
  if (existingRows.length) {
    for (let i = existingRows.length - 1; i >= 0; i--) {
      const row = existingRows[i];
      clearRow(ws, row);
      shiftRows(ws, row + 1, -1);
    }
  }

  // Recalcular tras borrar
  const r2 = decodeRef(ws);
  let insertAt = startRow;
  if (anchorText) {
    const anchorRow2 = findRowContaining(ws, anchorText);
    if (anchorRow2 != null) insertAt = anchorRow2 + 1;
  }

  // Encontrar de nuevo sizeCol por plantilla (si existía antes)
  // Si no lo tenemos, intentamos encontrar una columna donde en templateRow haya algo tipo "M/S/..."
  if (sizeCol == null) {
    // buscamos en la fila templateRow la primera celda tipo texto para usar como col talla (heurística)
    const rr = decodeRef(ws);
    const rowCheck = Math.min(Math.max(templateRow, rr.s.r), rr.e.r);
    for (let c = rr.s.c; c <= rr.e.c; c++) {
      const cell = getCell(ws, rowCheck, c);
      if (!cell || cell.v == null) continue;
      const v = normalizeSize(cell.v);
      if (sizesSet.has(v)) { sizeCol = c; break; }
    }
    // si sigue null, asumimos que en tu plantilla la talla está en columna E (4) (fallback)
    if (sizeCol == null) sizeCol = 4;
  }

  const qtyCol = sizeCol + 1;

  // Insertamos exactamente las filas necesarias (tallas con qty > 0) en el orden del array CONFIG.sizes
  const wanted = [];
  for (const size of CONFIG.sizes) {
    const ns = normalizeSize(size);
    if (qtyBySize.has(ns)) wanted.push(ns);
  }

  // Aseguramos que hay sitio insertando filas
  shiftRows(ws, insertAt, wanted.length);

  // Copiamos formato y escribimos talla/cantidad
  const refNow = decodeRef(ws);
  const modelRow = Math.min(Math.max(templateRow, refNow.s.r), refNow.e.r);

  for (let i = 0; i < wanted.length; i++) {
    const row = insertAt + i;
    cloneRow(ws, modelRow, row);
    setText(ws, row, sizeCol, wanted[i]);
    setNumber(ws, row, qtyCol, qtyBySize.get(wanted[i]));
  }
}

async function generate() {
  const wb = await loadTemplate();

  for (const p of CONFIG.products) {
    const ws = wb.Sheets[p.sheet];
    if (!ws) continue;

    for (const sec of p.sections) {
      applySection(ws, p.key, sec.key, sec.anchorText);
    }
  }

  for (const p of CONFIG.products) {
    const ws = wb.Sheets[p.sheet];
    if (!ws) continue;
    autoFitColumns(ws, 8, 70);
  }

  downloadWorkbook(wb, buildFilename());
}

function bind() {
  const btn = document.getElementById("btn");
  const btnMobile = document.getElementById("btnMobile");
  const btnClear = document.getElementById("btnClear");
  const btnClearMobile = document.getElementById("btnClearMobile");

  const run = async () => {
    try { await generate(); }
    catch (e) { alert(e?.message || String(e)); }
  };

  btn?.addEventListener("click", run);
  btnMobile?.addEventListener("click", run);

  const clear = () => location.reload();
  btnClear?.addEventListener("click", clear);
  btnClearMobile?.addEventListener("click", clear);
}

render();
bind();




