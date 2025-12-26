function n(id) {
  const v = Number(document.getElementById(id).value);
  return Number.isFinite(v) && v >= 0 ? v : 0;
}

// Set value safely in a worksheet cell
function setCell(ws, addr, value) {
  ws[addr] = ws[addr] || {};
  ws[addr].v = value;
  ws[addr].t = "n"; // numeric
}

async function loadTemplate() {
  const res = await fetch("./Hoja de pedido_actualizada.xlsx");
  if (!res.ok) throw new Error("No se pudo cargar la plantilla. ¿Está en la misma carpeta?");
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

document.getElementById("btn").addEventListener("click", async () => {
  try {
    const wb = await loadTemplate();

    // ======================
    // CAMISETA (columna F)
    // En tu plantilla: tallas en E3..E6 (M,L,XL,2XL) y cantidades en F3..F6
    // ======================
    const wsCam = wb.Sheets["CAMISETA"];
    setCell(wsCam, "F3", n("cam_m"));
    setCell(wsCam, "F4", n("cam_l"));
    setCell(wsCam, "F5", n("cam_xl"));
    setCell(wsCam, "F6", n("cam_2xl"));

    // ======================
    // SUDADERA
    // Técnica: filas 3..6 (M,L,XL,2XL) cantidades F3..F6
    // Paseo:   filas 7..10 (M,L,XL,2XL) cantidades F7..F10
    // ======================
    const wsSud = wb.Sheets["SUDADERA"];
    setCell(wsSud, "F3", n("sudt_m"));
    setCell(wsSud, "F4", n("sudt_l"));
    setCell(wsSud, "F5", n("sudt_xl"));
    setCell(wsSud, "F6", n("sudt_2xl"));

    setCell(wsSud, "F7", n("sudp_m"));
    setCell(wsSud, "F8", n("sudp_l"));
    setCell(wsSud, "F9", n("sudp_xl"));
    setCell(wsSud, "F10", n("sudp_2xl"));

    // ======================
    // PANTALON CHANDAL (columna F, filas 3..6)
    // ======================
    const wsPch = wb.Sheets["PANTALON CHANDAL"];
    setCell(wsPch, "F3", n("pch_m"));
    setCell(wsPch, "F4", n("pch_l"));
    setCell(wsPch, "F5", n("pch_xl"));
    setCell(wsPch, "F6", n("pch_2xl"));

    // PANTALON (hoja vacía en tu ejemplo): de momento no la tocamos.
    // Si quieres, luego la rellenamos con el mismo patrón.

    XLSX.writeFile(wb, buildFilename());
  } catch (err) {
    alert(err.message || String(err));
  }
});
