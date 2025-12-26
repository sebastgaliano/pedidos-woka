const CONFIG = {
  templatePath: "./Plantilla.xlsx",
  sizes: ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"],
  products: [
    {
      key: "equipación",
      title: "Equipación",
      sheet: "EQUIPACIÓN DE JUEGO",
      sections: [
        { key: "camiseta_juego", title: "Camiseta de juego", anchorText: "CAMISETA DE JUEGO" },
        { key: "pantalón_juego", title: "Pantalón de juego", anchorText: "PANTALÓN DE JUEGO" },
        { key: "cubre_juego", title: "Cubre de juego", anchorText: "CUBRE DE JUEGO" },
      ],
    },
    {
      key: "camiseta_paseo",
      title: "Camiseta de paseo",
      sheet: "CAMISETA",
      sections: [{ key: "base", title: null, anchorText: null }],
    },
    {
      key: "sudadera",
      title: "Sudadera",
      sheet: "SUDADERA",
      sections: [
        { key: "tecnica", title: "Sudadera técnica", anchorText: "SUDADERA TECNICA" },
        { key: "paseo", title: "Sudadera paseo", anchorText: "SUDADERA PASEO" },
      ],
    },
    {
      key: "pantalon_chandal",
      title: "Pantalón chándal",
      sheet: "PANTALON CHANDAL",
      sections: [{ key: "base", title: null, anchorText: null }],
    },

    // Futuro: solo añade aquí (sin tocar HTML)
    // { key:"equipacion_juego", title:"Equipación de juego", sheet:"EQUIPACION JUEGO", sections:[{key:"base", title:null, anchorText:null}] },
    // { key:"cubre_juego", title:"Cubre de juego", sheet:"CUBRE JUEGO", sections:[{key:"base", title:null, anchorText:null}] },
    // { key:"pantalon_juego", title:"Pantalón de juego", sheet:"PANTALON JUEGO", sections:[{key:"base", title:null, anchorText:null}] },
    // { key:"abrigo", title:"Abrigo", sheet:"ABRIGO", sections:[{key:"base", title:null, anchorText:null}] },
  ],
};

function collectSummary() {
  const lines = [];
  let grandTotal = 0;

  for (const p of CONFIG.products) {
    for (const sec of p.sections) {
      const entries = [];
      let subtotal = 0;

      for (const size of CONFIG.sizes) {
        const qty = n(idFor(p.key, sec.key, size));
        if (qty > 0) {
          entries.push({ size, qty });
          subtotal += qty;
        }
      }

      if (entries.length) {
        grandTotal += subtotal;

        const title = sec.title ? `${p.title} · ${sec.title}` : p.title;
        const items = entries.map(e => `${e.size}: ${e.qty}`).join(" · ");

        lines.push({ title, subtotal, items });
      }
    }
  }

  return { lines, grandTotal };
}

function renderSummaryHTML(summary) {
  if (!summary.lines.length) {
    return `
      <div style="font-weight:800;">No has añadido ninguna prenda.</div>
      <div style="color:#a7a7b3; font-size:13px; margin-top:6px;">Añade cantidades y vuelve a intentarlo.</div>
    `;
  }

  const blocks = summary.lines.map(l => `
    <div style="padding:10px 0; border-bottom:1px solid rgba(255,255,255,.08);">
      <div style="display:flex; justify-content:space-between; gap:10px; align-items:baseline;">
        <div style="font-weight:900;">${escapeHtml(l.title)}</div>
        <div style="color:#ffd8c2; font-weight:900;">${l.subtotal}</div>
      </div>
      <div style="color:#a7a7b3; font-size:13px; margin-top:4px; line-height:1.35;">
        ${escapeHtml(l.items)}
      </div>
    </div>
  `).join("");

  return `
    <div style="display:flex; justify-content:space-between; align-items:baseline; gap:12px; margin-bottom:10px;">
      <div style="color:#a7a7b3; font-size:13px;">Resumen de prendas</div>
      <div style="font-weight:900;">Total: ${summary.grandTotal}</div>
    </div>
    <div>${blocks}</div>
  `;
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
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

    for (const sec of p.sections) {
      if (sec.title) {
        const sub = document.createElement("div");
        sub.className = "subhead";
        sub.textContent = sec.title;
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
        input.id = idFor(p.key, sec.key, size);

        wrap.appendChild(label);
        wrap.appendChild(input);
        sizes.appendChild(wrap);
      }

      card.appendChild(sizes);
    }

    root.appendChild(card);
  }
}

async function loadTemplate() {
  const res = await fetch(CONFIG.templatePath, { cache: "no-store" });
  if (!res.ok) throw new Error("No se pudo cargar la plantilla.");
  const ab = await res.arrayBuffer();
  return XLSX.read(ab, { type: "array", cellStyles: true });
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

function decodeRef(ws) {
  if (!ws["!ref"]) ws["!ref"] = "A1:A1";
  return XLSX.utils.decode_range(ws["!ref"]);
}

function addr(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}

function getCell(ws, r, c) {
  return ws[addr(r, c)];
}

function normalizeSize(v) {
  return String(v ?? "").trim().toUpperCase().replace(/\s+/g, "");
}

function findRowContaining(ws, text) {
  if (!ws || !ws["!ref"]) return null;
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

function findHeaderRow(ws, startRow, endRow) {
  if (!ws || !ws["!ref"]) return null;

  const range = decodeRef(ws);
  const rStart = Math.max(startRow, range.s.r);
  const rEnd = Math.min(endRow, range.e.r);

  const required = ["nombre", "talla", "cantidad"];

  for (let r = rStart; r <= rEnd; r++) {
    let hit = 0;

    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = getCell(ws, r, c);
      if (!cell || cell.v == null) continue;

      const v = String(cell.v).toLowerCase().trim();
      if (required.includes(v)) hit++;
    }

    if (hit >= 2) return r;
  }

  return null;
}

function findHeaderCols(ws, headerRow) {
  const range = decodeRef(ws);
  let sizeCol = null;
  let qtyCol = null;

  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = getCell(ws, headerRow, c);
    if (!cell || cell.v == null) continue;

    const v = String(cell.v).toLowerCase().trim();
    if (v === "talla") sizeCol = c;
    if (v === "cantidad") qtyCol = c;
  }

  return { sizeCol, qtyCol };
}

function cloneCell(cell) {
  if (!cell) return undefined;
  return JSON.parse(JSON.stringify(cell));
}

function cloneRow(ws, srcRow, dstRow) {
  const range = decodeRef(ws);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const src = getCell(ws, srcRow, c);
    if (src) ws[addr(dstRow, c)] = cloneCell(src);
    else delete ws[addr(dstRow, c)];
  }
}

function clearRowValues(ws, r) {
  const range = decodeRef(ws);
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = getCell(ws, r, c);
    if (!cell) continue;
    delete cell.v;
    delete cell.w;
    delete cell.f;
    delete cell.t;
  }
}

function setText(ws, r, c, value) {
  const a = addr(r, c);
  ws[a] = ws[a] || {};
  ws[a].v = value;
  ws[a].t = "s";
}

function setNumber(ws, r, c, value) {
  const a = addr(r, c);
  ws[a] = ws[a] || {};
  ws[a].v = value;
  ws[a].t = "n";
}

function applyStandardCols(ws) {
  // Ajusta si tu proveedor usa más columnas. A..I cubre tu ejemplo.
  ws["!cols"] = [
    { wch: 24 }, // A Nombre
    { wch: 16 }, // B Imprimir Nombre
    { wch: 10 }, // C Número
    { wch: 16 }, // D Imprimir Número
    { wch: 8  }, // E Talla
    { wch: 10 }, // F Cantidad
    { wch: 18 }, // G Equipo
    { wch: 10 }, // H GENERO
    { wch: 14 }, // I COLEGIO
  ];
}

function applySection(ws, productKey, sectionKey, anchorText) {
  const wanted = [];
  for (const size of CONFIG.sizes) {
    const qty = n(idFor(productKey, sectionKey, size));
    if (qty > 0) wanted.push({ size: normalizeSize(size), qty });
  }

  const range = decodeRef(ws);

  let startRow = range.s.r;
  let endRow = range.e.r;

  if (anchorText) {
    const anchorRow = findRowContaining(ws, anchorText);
    if (anchorRow != null) startRow = anchorRow + 1;
  }

  const headerRow = findHeaderRow(ws, startRow, endRow);
  if (headerRow == null) return;

  const { sizeCol, qtyCol } = findHeaderCols(ws, headerRow);
  if (sizeCol == null || qtyCol == null) return;

  const dataStart = headerRow + 1;

  // Detectar hasta donde llega el bloque actual (sin mover filas ni merges)
  let dataEnd = dataStart - 1;
  let emptyStreak = 0;

  for (let r = dataStart; r <= range.e.r; r++) {
    const sCell = getCell(ws, r, sizeCol);
    const qCell = getCell(ws, r, qtyCol);
    const sVal = sCell?.v == null ? "" : String(sCell.v).trim();
    const qVal = qCell?.v == null ? "" : String(qCell.v).trim();

    if (sVal === "" && qVal === "") {
      emptyStreak++;
      if (emptyStreak >= 2) break;
    } else {
      emptyStreak = 0;
      dataEnd = r;
    }
  }

  const templateRow = dataStart; // fila “modelo” del bloque de datos

  // Limpiar bloque (dejamos formato intacto)
  const rowsNeeded = Math.max(wanted.length, 1);
  const clearTo = Math.max(dataEnd, dataStart + rowsNeeded + 12);

  for (let r = dataStart; r <= clearTo; r++) {
    if (r !== templateRow) cloneRow(ws, templateRow, r);
    clearRowValues(ws, r);
  }

  // Si no hay nada que escribir, dejamos el bloque vacío
  if (wanted.length === 0) return;

  // Escribir filas (se “añaden” tallas simplemente ocupando filas libres del bloque)
  for (let i = 0; i < wanted.length; i++) {
    const r = dataStart + i;
    if (r !== templateRow) cloneRow(ws, templateRow, r);
    setText(ws, r, sizeCol, wanted[i].size);
    setNumber(ws, r, qtyCol, wanted[i].qty);
  }
}

function openConfirmDialog(onConfirm) {
  const dlg = document.getElementById("confirmDialog");
  const content = document.getElementById("dlgContent");
  const btnConfirm = document.getElementById("dlgConfirm");
  const btnCancel = document.getElementById("dlgCancel");
  const btnClose = document.getElementById("dlgClose");

  const summary = collectSummary();
  content.innerHTML = renderSummaryHTML(summary);

  if (!summary.lines.length) {
    btnConfirm.disabled = true;
    btnConfirm.style.opacity = "0.6";
    btnConfirm.style.cursor = "not-allowed";
  } else {
    btnConfirm.disabled = false;
    btnConfirm.style.opacity = "";
    btnConfirm.style.cursor = "";
  }

  const close = () => dlg.close();
  btnCancel.onclick = close;
  btnClose.onclick = close;

  btnConfirm.onclick = async () => {
    close();
    await onConfirm();
  };

  dlg.showModal();
}

function confirmClear() {
  const ok = confirm("¿Seguro que quieres limpiar el formulario? Se perderán las cantidades.");
  if (ok) location.reload();
}

async function generate() {
  const wb = await loadTemplate();

  for (const p of CONFIG.products) {
    const ws = wb.Sheets[p.sheet];
    if (!ws) continue;

    for (const sec of p.sections) {
      applySection(ws, p.key, sec.key, sec.anchorText);
    }

    applyStandardCols(ws);
  }

  downloadWorkbook(wb, buildFilename());
}

function bind() {
  const run = async () => {
    openConfirmDialog(async () => {
      try { await generate(); }
      catch (e) { alert(e?.message || String(e)); }
    });
  };

  document.getElementById("btn")?.addEventListener("click", run);
  document.getElementById("btnMobile")?.addEventListener("click", run);

  document.getElementById("btnClear")?.addEventListener("click", confirmClear);
  document.getElementById("btnClearMobile")?.addEventListener("click", confirmClear);
}


render();
bind();




