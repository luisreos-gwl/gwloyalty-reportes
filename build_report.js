const pptxgen = require("pptxgenjs");
const D = JSON.parse(require("fs").readFileSync(process.argv[3] || "metrics.json","utf8"));
const money = (n) => "$" + Math.round(n).toLocaleString("en-US");
const pct = (a, b) => Math.round(((a - b) / b) * 100);
const arrow = (a, b) => (Math.abs(a - b) / b < 0.01 ? "≈" : a > b ? "▲" : "▼");
const MES = D.mes_label;
const AGENDA = "https://calendar.google.com/calendar/u/0/appointments/schedules/AcZssZ2Y2mHbXK6jK_EHNdrOq764-Zwt0vF6LvMCcMYQY-MkNA6eAmurZ_ypamPolruRobVqcBJrD6q_";
const WA = "https://api.whatsapp.com/message/GJ45VRAUFEHIG1?autoload=1&app_absent=0";
const PORTAL = "https://portal.gwloyalty.com/#paso-1";
// Reglas de la lámina 4 y 5 (el hallazgo que más pesa). Umbral: caída > 15%.
const CM = D.canje_mode;
const CLABEL = { both: "PUNTOS Y CASHBACK CANJEADOS", puntos: "PUNTOS CANJEADOS", cashback: "CASHBACK CANJEADO", recompensas: "RECOMPENSAS CANJEADAS" }[CM];
const CROW = { both: "Puntos canjeados", puntos: "Puntos canjeados", cashback: "Cashback canjeado", recompensas: "Recompensas canjeadas" }[CM];
const cfmt = (v) => (CM === "cashback" ? money(v) : Math.round(v).toLocaleString("en-US") + (CM === "recompensas" ? "" : " pts"));
const dNew = pct(D.nuevos, D.nuevos_prev), dCan = pct(D.canje, D.canje_prev), dIng = pct(D.ing, D.ing_prev);
const PUSH = ["Programa 1 push por semana", "Recuérdales volver con una promo corta de tu temporada.", "Video 2.2 · Notificaciones push"];
const REF = ["Invita a tus mejores clientes a recomendarte", "Los referidos son la forma más barata de traer clientes.", "Video 1.1 · Define tu estrategia"];
const REC = ["Revisa que tu recompensa sea alcanzable", "Si casi nadie la canjea, puede estar muy lejos o poco visible.", "Video 1.1 · Define tu estrategia"];
const EQ = ["Refuerza cómo ofrecen la tarjeta en caja", "Un buen pitch en el mostrador trae más registros.", "Video 2.1 · Lanza tu programa"];
let FX;
if (dNew <= -15) FX = { find: "Llegaron " + D.nuevos + " clientes nuevos, " + Math.abs(dNew) + "% menos que en " + D.mes_prev + (dCan <= -15 ? ", y los canjes bajaron " + Math.abs(dCan) + "%." : "."), hook: "Tus clientes nuevos bajaron " + Math.abs(dNew) + "% vs " + D.mes_prev + ". Te muestro cómo recuperarlos.", sug: [PUSH, REF] };
else if (dCan <= -15) FX = { find: "Los canjes bajaron " + Math.abs(dCan) + "% vs " + D.mes_prev + " (" + cfmt(D.canje) + " este mes).", hook: "Tus canjes bajaron " + Math.abs(dCan) + "%. Revisemos si tu recompensa está bien calibrada.", sug: [REC, PUSH] };
else if (dIng <= -15) FX = { find: "Tus ingresos de lealtad bajaron " + Math.abs(dIng) + "% vs " + D.mes_prev + ".", hook: "Tus ingresos bajaron " + Math.abs(dIng) + "% vs " + D.mes_prev + ". Veamos qué lo movió y cómo recuperarlo.", sug: [PUSH, EQ] };
else FX = { find: "Todo se mantiene estable vs " + D.mes_prev + ": hay espacio para crecer con más clientes nuevos.", hook: "Tu programa va estable. Te muestro cómo sumar más clientes nuevos el próximo mes.", sug: [REF, PUSH] };
const OUT = process.argv[2] || "reporte.pptx";
const fs = require("fs");


// ---------- Marca (README del sistema de diseño GW Loyalty) ----------
const INK = "121D1F", MINT = "6CCAA8", DEEP = "243F3A", FOREST = "3C6E5F", GREY = "616161";
const WHITE = "FFFFFF", CREAM = "F7F9F4", COOL = "F1FAF7", MUTED_DARK = "ACB0B1"; // blanco 65% sobre tinta
const F = "Poppins", FM = "Poppins Medium";
const px = (v) => v / 96;          // 1080 px = 11.25 in
const pt = (v) => v * 0.75;        // px -> pt

const pres = new pptxgen();
pres.defineLayout({ name: "IG45", width: px(1080), height: px(1350) });
pres.layout = "IG45";
pres.title = "Reporte mensual";
pres.author = "GW Loyalty";

const LOGO_WHITE = "logo_blanco.png";
const LOGO_DARK = "logo_tinta.png";

// ---------- Helpers ----------
function logo(slide, path, color) {
  if (fs.existsSync(path)) slide.addImage({ path, x: px(60), y: px(80), w: px(190), h: px(39.5), objectName: "logo", altText: "GW Loyalty" });
  else text(slide, "GW LOYALTY", { x: 60, y: 80, w: 300, h: 40, size: 30, bold: true, color, valign: "middle" });
}
function text(slide, str, o) {
  slide.addText(str, {
    x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h),
    fontFace: o.font || F, fontSize: pt(o.size), color: o.color || INK, bold: !!o.bold,
    align: o.align || "left", valign: o.valign || "top", margin: 0, isTextBox: true,
    charSpacing: o.cs || 0, lineSpacingMultiple: o.lh || 1.0, fit: "none", wrap: true,
    objectName: o.name,
    ...(o.link ? { hyperlink: { url: o.link } } : {}),
  });
}
function box(slide, o) {
  const shape = o.radius ? pres.ShapeType.roundRect : pres.ShapeType.rect;
  const opts = { x: px(o.x), y: px(o.y), w: px(o.w), h: px(o.h), objectName: o.name };
  if (o.radius) opts.rectRadius = px(Math.min(o.radius, o.h / 2));
  opts.fill = o.fill ? { color: o.fill, transparency: o.ft || 0 } : { type: "none" };
  opts.line = o.line ? { color: o.line, width: pt(o.lw || 1), transparency: o.lt || 0 } : { type: "none" };
  slide.addShape(shape, opts);
}
function pill(slide, label, o) {
  box(slide, { x: o.x, y: o.y, w: o.w, h: o.h, radius: o.h / 2, fill: o.fill, ft: o.ft, line: o.line, lt: o.lt, lw: 1, name: "pill-" + label });
  text(slide, label, { x: o.x, y: o.y, w: o.w, h: o.h, size: o.size || 20, font: FM, color: o.color, align: "center", valign: "middle", cs: 1.3 });
}
function progress(slide, n, dark) {
  const active = dark ? MINT : INK;
  for (let i = 1; i <= 5; i++) {
    box(slide, {
      x: 60 + (i - 1) * 74, y: 1240, w: 64, h: 6, radius: 3,
      fill: i <= n ? active : (dark ? WHITE : INK), ft: i <= n ? 0 : (dark ? 80 : 85), name: "progreso-" + i,
    });
  }
  if (n < 5) text(slide, "DESLIZA →", { x: 760, y: 1226, w: 260, h: 34, size: 22, font: FM, color: dark ? MINT : INK, align: "right", valign: "middle", cs: 1.3 });
}
function topRow(slide, n, tag, dark) {
  const c = dark ? MUTED_DARK : GREY;
  text(slide, "0" + n + " / 05", { x: 820, y: 88, w: 200, h: 30, size: 22, font: FM, color: c, align: "right", valign: "middle", cs: 1.3 });
  if (tag) {
    const w = Math.round(tag.length * 17 + 56);
    pill(slide, tag, { x: 60, y: 80, w, h: 48, size: 20, color: dark ? WHITE : INK, line: dark ? WHITE : INK, lt: dark ? 82 : 85 });
  }
}

// =====================================================================
// LÁMINA 1 — Tu mes de un vistazo (infografía; se envía también como imagen)
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: INK };
  logo(s, LOGO_WHITE, WHITE);
  text(s, "01 / 05", { x: 820, y: 88, w: 200, h: 30, size: 22, font: FM, color: MUTED_DARK, align: "right", valign: "middle", cs: 1.3 });
  pill(s, "REPORTE MENSUAL · " + MES, { x: 60, y: 150, w: 520, h: 48, size: 20, color: WHITE, line: WHITE, lt: 82 });
  text(s, D.cliente, { x: 60, y: 218, w: 960, h: 44, size: 34, bold: true, color: WHITE, valign: "middle" });

  text(s, "POR CADA $1 QUE HAS PAGADO A GW LOYALTY", { x: 60, y: 292, w: 960, h: 34, size: 26, font: FM, color: MUTED_DARK, valign: "middle", cs: 1.3 });
  text(s, "$" + Math.round(D.roi_rep), { x: 60, y: 336, w: 430, h: 210, size: 170, bold: true, color: MINT, valign: "middle", cs: -4 });
  text(s, "gastaron tus clientes que regresaron", { x: 500, y: 372, w: 520, h: 140, size: 40, color: WHITE, valign: "middle", lh: 1.1 });

  const tiles = [
    ["CLIENTES REGISTRADOS", D.clientes.toLocaleString("en-US"), "▲ +" + D.nuevos + " este mes"],
    ["INGRESOS DE LEALTAD (MES)", money(D.ing), money(D.acum) + " acumulado"],
    ["TICKET PROMEDIO", money(D.ticket), arrow(D.ticket, D.ticket_prev) + " " + Math.abs(pct(D.ticket, D.ticket_prev)) + "% vs mes anterior"],
    [CLABEL, cfmt(D.canje), CM === "both" ? "+ " + money(D.canje_cash) + " en cashback" : arrow(D.canje, D.canje_prev) + " " + Math.abs(pct(D.canje, D.canje_prev)) + "% vs mes anterior"],
    ["MEJOR CLIENTE · GASTO", money(D.mejor_gasto.gasto), D.mejor_gasto.nombre],
    ["MEJOR CLIENTE · VISITAS", D.mejor_visitas.n + " visitas", D.mejor_visitas.nombre],
  ];
  tiles.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 60 + col * (468 + 24), y = 600 + row * (150 + 16);
    box(s, { x, y, w: 468, h: 150, radius: 30, fill: WHITE, ft: 90, line: WHITE, lt: 82, name: "tile-" + (i + 1) });
    text(s, t[0], { x: x + 28, y: y + 18, w: 412, h: 28, size: 20, font: FM, color: MUTED_DARK, valign: "middle", cs: 1.0 });
    text(s, t[1], { x: x + 28, y: y + 50, w: 412, h: 56, size: 46, bold: true, color: WHITE, valign: "middle" });
    text(s, t[2], { x: x + 28, y: y + 108, w: 412, h: 30, size: 24, color: MINT, valign: "middle" });
  });

  box(s, { x: 60, y: 1116, w: 960, h: 64, radius: 32, fill: MINT, name: "chip-alertas" });
  text(s, (D.n_big + D.n_dob) + " transacciones por revisar · lámina 3", { x: 60, y: 1116, w: 960, h: 64, size: 26, bold: true, color: INK, align: "center", valign: "middle" });
  progress(s, 1, true);
}

// =====================================================================
// LÁMINA 2 — Así vas vs. antes
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: CREAM };
  topRow(s, 2, "COMPARATIVO", false);
  text(s, [
    { text: "ASÍ VAS ", options: { color: INK } },
    { text: "FRENTE", options: { color: FOREST, breakLine: true } },
    { text: "A MESES ANTERIORES", options: { color: INK } },
  ], { x: 60, y: 160, w: 960, h: 150, size: 60, bold: true, lh: 1.05 });

  // encabezados de columna
  const cx = [540, 700, 850]; const cw = [170, 140, 150];
  ["ESTE MES", "MES ANT.", "PROM. 3 M."].forEach((h, i) => text(s, h, { x: cx[i], y: 352, w: cw[i], h: 28, size: 20, font: FM, color: GREY, align: "right", cs: 1.0 }));
  const A2 = (x, y) => arrow(x, y);
  const rows = [
    ["Clientes nuevos", String(D.nuevos), String(D.nuevos_prev), String(Math.round(D.nuevos_avg)), A2(D.nuevos, D.nuevos_prev)],
    ["Ingresos de lealtad", money(D.ing), money(D.ing_prev), money(D.ing_avg), A2(D.ing, D.ing_prev)],
    ["Ticket promedio", money(D.ticket), money(D.ticket_prev), money(D.ticket_avg), A2(D.ticket, D.ticket_prev)],
    [CROW, cfmt(D.canje), cfmt(D.canje_prev), cfmt(D.canje_avg), A2(D.canje, D.canje_prev)],
  ];
  rows.forEach((r, i) => {
    const y = 392 + i * (120 + 14);
    box(s, { x: 60, y, w: 960, h: 120, radius: 30, fill: WHITE, line: INK, lt: 90, name: "fila-" + (i + 1) });
    text(s, r[0], { x: 90, y, w: 330, h: 120, size: 28, font: FM, color: INK, valign: "middle", lh: 1.1 });
    text(s, [{ text: r[4] + " ", options: { color: FOREST, fontSize: pt(22) } }, { text: r[1], options: { color: INK, bold: true, fontSize: pt(36) } }],
      { x: cx[0] - 70, y, w: cw[0] + 70, h: 120, size: 36, align: "right", valign: "middle" });
    text(s, r[2], { x: cx[1], y, w: cw[1], h: 120, size: 30, color: GREY, align: "right", valign: "middle" });
    text(s, r[3], { x: cx[2], y, w: cw[2] - 20, h: 120, size: 30, color: GREY, align: "right", valign: "middle" });
  });
  text(s, "Ingresos " + (D.ing >= D.ing_prev ? "+" : "") + pct(D.ing, D.ing_prev) + "% vs " + D.mes_prev + "; clientes nuevos " + pct(D.nuevos, D.nuevos_prev) + "%.", { x: 60, y: 960, w: 960, h: 44, size: 30, color: INK, valign: "middle" });

  // ROI acumulado (barras con rectángulos nativos)
  box(s, { x: 60, y: 1030, w: 960, h: 170, radius: 30, fill: WHITE, line: INK, lt: 90, name: "card-roi" });
  text(s, "CLIENTES QUE REGRESARON", { x: 90, y: 1052, w: 400, h: 26, size: 20, font: FM, color: GREY, cs: 1.0 });
  text(s, "$" + Math.round(D.roi_rep), { x: 90, y: 1078, w: 250, h: 84, size: 66, bold: true, color: FOREST, valign: "middle", cs: -2 });
  text(s, "por cada $1 que has pagado", { x: 90, y: 1160, w: 420, h: 30, size: 24, color: GREY, valign: "middle" });
  const vals = Object.values(D.rep_mes), months = Object.keys(D.rep_mes).map((k) => ["E","F","M","A","M","J","J","A","S","O","N","D"][parseInt(k.slice(5), 10) - 1]);
  vals.forEach((v, i) => {
    const h = Math.round((v / Math.max(...vals)) * 92), x = 600 + i * 66, base = 1140;
    box(s, { x, y: base - h, w: 40, h, radius: 8, fill: i === 5 ? INK : MINT, name: "barra-" + (i + 1) });
    text(s, months[i], { x, y: 1146, w: 40, h: 28, size: 22, font: FM, color: GREY, align: "center", valign: "middle" });
  });
  progress(s, 2, false);
}

// =====================================================================
// LÁMINA 3 — Revisión de transacciones
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  topRow(s, 3, "TRANSACCIONES", false);
  text(s, [
    { text: "REVISAMOS TUS", options: { color: INK, breakLine: true } },
    { text: "TRANSACCIONES", options: { color: FOREST } },
  ], { x: 60, y: 160, w: 960, h: 150, size: 60, bold: true, lh: 1.05 });
  text(s, "Buscamos montos atípicos y registros dobles. Estas son las 2 más relevantes.", { x: 60, y: 340, w: 900, h: 110, size: 32, color: GREY, lh: 1.3 });
  pill(s, D.visitas + " REVISADAS · " + (D.n_big + D.n_dob) + " POR REVISAR", { x: 60, y: 480, w: 600, h: 56, size: 22, color: INK, fill: MINT, line: null });

  const alerts = [];
  if (D.top_big) alerts.push(["MONTO ALTO", D.top_big.fecha, "Una compra de " + money(D.top_big.monto) + ": " + Math.round(D.top_big.monto / D.mediana) + " veces tu ticket mediano (" + money(D.mediana) + ").", "Confirma con tu equipo que el monto sea correcto."]);
  if (D.top_dob) alerts.push(["POSIBLE DOBLE REGISTRO", D.top_dob.fecha, "Dos registros de " + money(D.top_dob.monto) + " en la misma tarjeta con " + Math.max(1, Math.round(D.top_dob.dt * 60)) + " segundos de diferencia.", "Revisa que no se haya escaneado dos veces."]);
  if (!alerts.length) alerts.push(["SIN ALERTAS", "", "No encontramos montos atípicos ni registros dobles este mes.", "Sigue así: tu operación en caja se ve limpia."]);
  alerts.forEach((a, i) => {
    const y = 580 + i * (290 + 24);
    box(s, { x: 60, y, w: 960, h: 290, radius: 30, fill: CREAM, line: INK, lt: 90, name: "alerta-" + (i + 1) });
    pill(s, a[0], { x: 92, y: y + 30, w: 360, h: 44, size: 19, color: WHITE, fill: INK });
    text(s, a[1], { x: 700, y: y + 30, w: 290, h: 44, size: 22, font: FM, color: GREY, align: "right", valign: "middle", cs: 1.0 });
    text(s, a[2], { x: 92, y: y + 96, w: 896, h: 100, size: 32, bold: true, color: INK, lh: 1.15 });
    text(s, [{ text: "Qué hacer:  ", options: { color: FOREST, bold: true } }, { text: a[3], options: { color: GREY } }], { x: 92, y: y + 204, w: 896, h: 70, size: 28, lh: 1.2 });
  });
  progress(s, 3, false);
}

// =====================================================================
// LÁMINA 4 — Lo que más moverá tu siguiente mes
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: CREAM };
  topRow(s, 4, "SIGUIENTE PASO", false);
  text(s, [
    { text: "LO QUE MÁS ", options: { color: INK } },
    { text: "MOVERÁ", options: { color: FOREST, breakLine: true } },
    { text: "TU SIGUIENTE MES", options: { color: INK } },
  ], { x: 60, y: 160, w: 960, h: 150, size: 60, bold: true, lh: 1.05 });

  box(s, { x: 60, y: 350, w: 960, h: 200, radius: 30, fill: MINT, name: "hallazgo" });
  text(s, "LO QUE VIMOS", { x: 92, y: 376, w: 400, h: 26, size: 20, font: FM, color: DEEP, cs: 1.0 });
  text(s, FX.find, { x: 92, y: 410, w: 896, h: 120, size: 32, bold: true, color: INK, lh: 1.2 });

  const sug = FX.sug.map((x, i) => [String(i + 1), x[0], x[1], x[2]]);
  sug.forEach((a, i) => {
    const y = 580 + i * (300 + 20);
    box(s, { x: 60, y, w: 960, h: 300, radius: 30, fill: WHITE, line: INK, lt: 90, name: "sugerencia-" + (i + 1) });
    box(s, { x: 92, y: y + 30, w: 56, h: 56, radius: 28, fill: INK, name: "num-" + a[0] });
    text(s, a[0], { x: 92, y: y + 30, w: 56, h: 56, size: 26, bold: true, color: WHITE, align: "center", valign: "middle" });
    text(s, a[1], { x: 168, y: y + 24, w: 820, h: 70, size: 34, bold: true, color: INK, valign: "middle", lh: 1.1 });
    text(s, a[2], { x: 92, y: y + 110, w: 896, h: 70, size: 30, color: GREY, lh: 1.2 });
    box(s, { x: 92, y: y + 208, w: 640, h: 60, radius: 30, fill: INK, name: "boton-video-" + a[0] });
    s.addShape(pres.ShapeType.triangle, { x: px(122), y: px(y + 226), w: px(20), h: px(24), rotate: 90, fill: { color: MINT }, line: { type: "none" }, objectName: "icono-play-" + a[0] });
    text(s, a[3], { x: 160, y: y + 208, w: 560, h: 60, size: 24, font: FM, color: WHITE, valign: "middle", link: PORTAL });
  });
  progress(s, 4, false);
}

// =====================================================================
// LÁMINA 5 — Hablemos (cierre)
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: MINT };
  logo(s, LOGO_DARK, INK);
  text(s, "05 / 05", { x: 820, y: 88, w: 200, h: 30, size: 22, font: FM, color: DEEP, align: "right", valign: "middle", cs: 1.3 });
  text(s, [
    { text: "¿VEMOS JUNTOS", options: { breakLine: true } },
    { text: "CÓMO MEJORARLO?" },
  ], { x: 60, y: 260, w: 960, h: 190, size: 64, bold: true, color: INK, lh: 1.05 });
  text(s, "Una llamada de 20 minutos, sin costo.", { x: 60, y: 470, w: 960, h: 44, size: 32, color: DEEP, valign: "middle" });

  box(s, { x: 60, y: 560, w: 960, h: 270, radius: 30, fill: WHITE, ft: 20, line: WHITE, lt: 40, name: "card-gancho" });
  text(s, "UN TEMA PARA LA LLAMADA", { x: 92, y: 588, w: 700, h: 26, size: 20, font: FM, color: DEEP, cs: 1.0 });
  text(s, FX.hook, { x: 92, y: 630, w: 896, h: 170, size: 38, bold: true, color: INK, lh: 1.2 });

  box(s, { x: 60, y: 880, w: 520, h: 84, radius: 42, fill: INK, name: "boton-agendar" });
  text(s, "Agendar mi llamada", { x: 60, y: 880, w: 520, h: 84, size: 30, bold: true, color: WHITE, align: "center", valign: "middle", link: AGENDA });
  box(s, { x: 60, y: 990, w: 520, h: 84, radius: 42, line: INK, lw: 2, name: "boton-whatsapp" });
  text(s, "Escribir por WhatsApp", { x: 60, y: 990, w: 520, h: 84, size: 30, bold: true, color: INK, align: "center", valign: "middle", link: WA });
  text(s, "Un solo número para dudas, cambios y soporte.", { x: 60, y: 1100, w: 960, h: 40, size: 26, color: DEEP, valign: "middle" });
  text(s, "gwloyalty.com · @gwloyalty", { x: 60, y: 1180, w: 700, h: 34, size: 24, font: FM, color: DEEP, valign: "middle", cs: 1.0 });
  for (let i = 1; i <= 5; i++) box(s, { x: 60 + (i - 1) * 74, y: 1240, w: 64, h: 6, radius: 3, fill: INK, name: "progreso-" + i });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("ok", f));
