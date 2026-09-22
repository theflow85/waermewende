const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");

// ---------- Design tokens (aus der Papier-Palette) ----------
const C = {
  paper: "F7F5EF", paperAlt: "EFEBE1", ink: "1A1A17", inkSoft: "56534B", rule: "C9C3B4",
  wp: "0F5257",      // Wärmepumpe (belegt-teal)
  fw: "1F5673",      // Fernwärme (accent-blau)
  gas: "7A2E2E",     // Gas (warn)
  setz: "9A5B12",    // Setzung / Unsicherheit
  green: "28633B", greenBg: "E4EEE3", redBg: "F5E4E0", yellowBg: "F4EBD3",
  blueBg: "E2E8ED", tealBg: "DEE9E8",
  grey: "8C877A",    // WP heute (DE) in Charts
  dark: "1A1A17", white: "FFFFFF",
};
const FH = "Cambria", FB = "Calibri";

const _iconCache = new Map();
async function icon(name, color, px = 256) {
  const key = name + "|" + color;
  if (_iconCache.has(key)) return _iconCache.get(key);
  if (!lu[name]) throw new Error("Unbekanntes Icon: " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(lu[name], { color: "#" + color, size: px, strokeWidth: 1.75 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  const out = "image/png;base64," + buf.toString("base64");
  _iconCache.set(key, out);
  return out;
}

function frame(slide, kicker, title, page, sub) {
  slide.background = { color: C.paper };
  slide.addText(kicker, { x: 0.6, y: 0.35, w: 8, h: 0.3, fontFace: FB, fontSize: 11, color: C.inkSoft, charSpacing: 2, isTextBox: true, margin: 0 });
  slide.addText(title, { x: 0.6, y: 0.65, w: 12.1, h: 0.95, fontFace: FH, fontSize: 28, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
  if (sub) slide.addText(sub, { x: 0.6, y: 1.35, w: 12.1, h: 0.55, fontFace: FB, fontSize: 9, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
  slide.addText("WP/FW-Vergleich · BAG Energie · 20. Sept. 2026", { x: 0.6, y: 7.0, w: 8, h: 0.25, fontFace: FB, fontSize: 9, color: C.inkSoft, isTextBox: true, margin: 0 });
  slide.addText(String(page), { x: 12.2, y: 7.0, w: 0.5, h: 0.25, fontFace: FB, fontSize: 9, color: C.inkSoft, align: "right", isTextBox: true, margin: 0 });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.author = "Florian Schulte";
  pres.title = "Wärmepumpen in Deutschland — Kostenproblem, Reformhebel, Systemvergleich";

  // ---------- Bausteine ----------
  async function iconRow(s, o) {
    const d = o.d || 0.5;
    s.addShape(pres.shapes.OVAL, { x: o.x, y: o.y + 0.02, w: d, h: d, fill: { color: o.bg }, line: { color: o.bg } });
    s.addImage({ data: await icon(o.ic, o.col), x: o.x + d * 0.2, y: o.y + 0.02 + d * 0.2, w: d * 0.6, h: d * 0.6 });
    const tx = o.x + d + 0.14, tw = o.w - d - 0.14;
    if (o.head) {
      s.addText(o.head, { x: tx, y: o.y, w: tw, h: o.headH || 0.32, fontFace: FB, fontSize: o.headSize || 13, bold: true, color: o.headColor || C.ink, isTextBox: true, margin: 0, valign: "middle" });
      s.addText(o.body, { x: tx, y: o.y + (o.headH || 0.32), w: tw, h: o.bodyH || 0.9, fontFace: FB, fontSize: o.bodySize || 11, color: o.bodyColor || C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    } else {
      s.addText(o.body, { x: tx, y: o.y, w: tw, h: o.bodyH || 0.5, fontFace: FB, fontSize: o.bodySize || 11, color: o.bodyColor || C.ink, isTextBox: true, margin: 0, valign: "top" });
    }
  }

  function fazit(s, lead, text, leadColor) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.95, w: 12.1, h: 0.9, rectRadius: 0.08, fill: { color: C.paperAlt }, line: { color: C.rule, width: 0.75 } });
    s.addText([
      { text: lead, options: { bold: true, color: leadColor || C.fw } },
      { text: text },
    ], { x: 0.8, y: 5.95, w: 11.7, h: 0.9, fontFace: FB, fontSize: 12.5, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
  }

  // gemeinsame Chart-Optionen (frisches Objekt je Aufruf – pptxgenjs mutiert in place)
  function chartBase() {
    return {
      barGapWidthPct: 55,
      showLegend: false, showTitle: true, titleFontFace: FB, titleFontSize: 12, titleColor: C.ink, titleAlign: "left",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: FB, dataLabelFontSize: 11, dataLabelColor: C.ink, dataLabelFontBold: true,
      catAxisLabelFontFace: FB, catAxisLabelFontSize: 10, catAxisLabelColor: C.inkSoft, catAxisLineShow: false, catAxisLabelRotate: 0,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      legendFontFace: FB, legendFontSize: 10, legendColor: C.inkSoft,
      plotArea: { fill: { color: C.paper } }, chartArea: { fill: { color: C.paper } },
    };
  }
  const FMT_DEC = "0\".\"00";      // Ganzzahl ×100 -> 17.75
  const FMT_TSD = "#\"'\"##0";     // 30000 -> 30'000

  const GAS_SUB = "* Genauer: Vollfaktoriell über 27 Kombinationen (Kesseltausch 6'000/9'000/12'000 € · Gas 9/11/13 ct/kWh · Nutzungsgrad 85/88/92 %) liegt der Kessel in 22 von 27 unter dem heutigen deutschen WP-Preis von 23.73 ct und in 1 von 27 unter der EU-Referenz von 17.75 ct — bei 6'000 €, 9 ct und 92 % mit 17.35 ct/kWh. Grüngaspflicht nach § 43 GModG barwertgewichtet enthalten (2.31 ct); ohne sie 19.20 ct, gegen die Referenz weiterhin teurer. Spannweite 17.35–25.73 ct/kWh; Abschnitt 4.2, Anhang A.10.";
  const GAS_NOTES = "Hintergrund für die Diskussion: Pipeline-Lobby und Fernwärmeseite berichten dasselbe Problem — außer alten Gasheizungen rechnet sich derzeit nichts, weil Haushalte und Planer gegen den aktuellen Fossilgaspreis und die Gastherme rechnen und implizit davon ausgehen, dass 2045 nicht durchgesetzt wird. Den Wandel erzeugt nur die bessere Wirtschaftlichkeit der Erneuerbaren, und die glaubt man nur, wenn die Regulierung nicht wieder kassiert wird. Genau deshalb ist Teil I keine Nebensache: Der Vergleichsmaßstab entscheidet über das Vorzeichen — beim Kessel wie beim Netz. Zahlen: Abschnitt 4.2 und Anhang A.10 (Gasheizung), Abschnitt 7 und Anhang A.7 (Gasvorlauf im Netz). 27 Kombinationen = Investition 6/9/12 T€ × Gaspreis 9/11/13 ct × Nutzungsgrad.";
  const HEAT_NOTES = "Rasterfolie zu Befund 4 — hier wird sichtbar, was die Balken auf der Vorfolie nur als Faktor behaupten. Drei Kacheln, dasselbe Raster, nur der Vergleichsmaßstab wechselt. Links gegen die europäisch normalisierte Referenz von 17.75 ct: Die Abwärmespalte trägt ab dem Reihenhaus, alle drei anderen Spalten verlieren, Gasvorlauf vor Großwärmepumpe ist durchgängig „nie“. Mitte gegen die dezentrale Gasheizung mit 21.51 ct: Anschluss an die Gas-Folie — solange die Wärmepumpe 30'000 € kostet, ist Gas der billigste dezentrale Pfad, und schon gegen diesen schwächeren Maßstab trägt das Netz in drei von vier Spalten. Rechts gegen die heutigen deutschen 23.73 ct: fast vollständig grün, einschließlich des Pfads, der links gar nicht existiert. Kern: Nicht einzelne Schwellen verschieben sich, sondern ein ganzer Technologiepfad entsteht oder verschwindet mit dem Vergleichsmaßstab. Zum EFH-Gebiet: erreichte Dichte aus der Bebauungsgeometrie hergeleitet — 0.86 aus 17 Trassenmetern, plausibel 13 bis 30 m und damit 0.49 bis 1.13; die EFH-Zeile liegt in allen drei Kacheln nah an der Kippkante. Konsequenz trotzdem eindeutig: keine gesicherte Grundlage heißt kein Netz mit fünfzig Jahren Bindung. Alle Zellen hier gegen die Bestandsstraße gerechnet, daher Gründerzeitblock 6.37 statt 3.04 wie auf Folie 6 (dort Innenstadtstraße). Fundstellen: Abschnitt 5.1, Tabelle in 5.2, Anhang A.11.";
  function HEATMAP_TABLE(s) {
    const FW_KLAR = "C0DD97", FW_KNAPP = "EAF3DE", WP_KNAPP = "E6F1FB", WP_KLAR = "85B7EB";
    const TXT = { C0DD97: "173404", EAF3DE: "3B6D11", E6F1FB: "185FA5", "85B7EB": "042C53" };
    const z = (v, bg) => ({ text: v === "nie" ? "×" : "", options: { fill: { color: bg }, color: TXT[bg], align: "center", valign: "middle", fontFace: FB, fontSize: 12, bold: true } });
    const kopf = (t, span) => ({ text: t, options: { colspan: span, align: "center", fontFace: FB, fontSize: 11, bold: true, color: C.ink, valign: "middle" } });
    const subh = (t) => ({ text: t, options: { align: "center", fontFace: FB, fontSize: 9, color: C.inkSoft, valign: "middle" } });
    const zeile = (t) => ({ text: t, options: { fontFace: FB, fontSize: 10.5, color: C.ink, valign: "middle" } });
    const q = [subh("Abwärme"), subh("GroßWP"), subh("Gas→Abw"), subh("Gas→GWP")];
    const rows = [
      [{ text: "", options: {} }, kopf("gegen WP-Referenz (EU) · 17.75 ct", 4), kopf("gegen Gasheizung · 21.51 ct", 4), kopf("gegen WP heute (DE) · 23.73 ct", 4)],
      [{ text: "", options: {} }, ...q, ...q, ...q],
      [zeile("EFH-Gebiet"), z("0.80", WP_KNAPP), z("0.16", WP_KLAR), z("0.21", WP_KLAR), z("nie", WP_KLAR), z("1.19", FW_KNAPP), z("0.55", WP_KNAPP), z("0.60", WP_KNAPP), z("0.31", WP_KLAR), z("1.43", FW_KNAPP), z("0.78", WP_KNAPP), z("0.83", WP_KNAPP), z("0.54", WP_KNAPP)],
      [zeile("Reihenhaus"), z("1.99", FW_KNAPP), z("0.38", WP_KLAR), z("0.50", WP_KNAPP), z("nie", WP_KLAR), z("2.96", FW_KLAR), z("1.35", FW_KNAPP), z("1.48", FW_KNAPP), z("0.77", WP_KNAPP), z("3.54", FW_KLAR), z("1.93", FW_KNAPP), z("2.05", FW_KLAR), z("1.34", FW_KNAPP)],
      [zeile("Kleines MFH, 5 WE"), z("3.61", FW_KLAR), z("0.67", WP_KNAPP), z("0.90", WP_KNAPP), z("nie", WP_KLAR), z("5.39", FW_KLAR), z("2.45", FW_KLAR), z("2.68", FW_KLAR), z("1.38", FW_KNAPP), z("6.44", FW_KLAR), z("3.50", FW_KLAR), z("3.73", FW_KLAR), z("2.43", FW_KLAR)],
      [zeile("Gründerzeitblock"), z("6.37", FW_KLAR), z("0.78", WP_KNAPP), z("1.21", FW_KNAPP), z("nie", WP_KLAR), z("9.75", FW_KLAR), z("4.17", FW_KLAR), z("4.60", FW_KLAR), z("2.13", FW_KLAR), z("11.75", FW_KLAR), z("6.16", FW_KLAR), z("6.60", FW_KLAR), z("4.13", FW_KLAR)],
      [zeile("Plattenbau-Zeile"), z("3.18", FW_KLAR), z("nie", WP_KLAR), z("0.25", WP_KLAR), z("nie", WP_KLAR), z("5.11", FW_KLAR), z("1.93", FW_KNAPP), z("2.17", FW_KLAR), z("0.77", WP_KNAPP), z("6.24", FW_KLAR), z("3.06", FW_KLAR), z("3.31", FW_KLAR), z("1.91", FW_KNAPP)],
    ];
    s.addTable(rows, { x: 0.6, y: 1.95, w: 12.1, colW: [1.78, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86, 0.86], rowH: [0.34, 0.26, 0.42, 0.42, 0.42, 0.42, 0.42], border: { type: "solid", color: "FFFFFF", pt: 1.5 }, fontFace: FB });
  }

  // ============ 0. Titelfolie (dunkel) ============
  {
    const s = pres.addSlide();
    s.background = { color: C.dark };
    s.addText("Diskussionspapier · BAG Energie · Stuttgart, 20. September 2026", { x: 0.8, y: 0.7, w: 11, h: 0.3, fontFace: FB, fontSize: 12, color: "B8B3A6", charSpacing: 2, isTextBox: true, margin: 0 });
    s.addText("Europäische Wärmepumpe und Fernwärme im Kostenvergleich", { x: 0.8, y: 1.9, w: 11.7, h: 2.0, fontFace: FH, fontSize: 40, bold: true, color: C.white, isTextBox: true, margin: 0 });
    const stats = [["30'000 €", "kostet eine Wärmepumpe in Deutschland"], ["14'000 €", "dieselbe Anlage in UK, FR, AT"], ["Ist Fernwärme wirtschaftlich?", "Kommt darauf an."]];
    stats.forEach(([n, l], i) => {
      const x = 0.8 + i * 3.9;
      s.addText(n, { x, y: 4.30, w: 3.7, h: 1.05, fontFace: FH, fontSize: i === 2 ? 22 : 30, bold: true, color: i === 2 ? "7FB7B2" : C.white, isTextBox: true, margin: 0, valign: "bottom" });
      s.addText(l, { x, y: 5.4, w: 3.5, h: 0.7, fontFace: FB, fontSize: 12, color: "B8B3A6", isTextBox: true, margin: 0, valign: "top" });
    });
    s.addText("Florian Schulte, Gießen · Version 2.11.1 · Papier und Folien CC BY 4.0 · Skript MIT", { x: 0.8, y: 6.7, w: 11, h: 0.3, fontFace: FB, fontSize: 11, color: "B8B3A6", isTextBox: true, margin: 0 });
  }


  // ============ V7 — Methodik kompakt (Seite 1) ============
  {
    const s = pres.addSlide();
    frame(s, "ZUR ENTSTEHUNG", "Ein Quereinsteiger, eine KI und ein offenes Skript", 1);
    const cols = [
      ["LuBot", C.fw, "KI", "Recherche, Kostenmodell, Simulation, Text"],
      ["LuUserCheck", C.wp, "Autor", "Fragestellung, Einwände, Prüfung, Lektorat"],
      ["LuFileCode", C.green, "Nachprüfbar", "Annahmen markiert, Quellen mit Interessenlage, Skript offen (MIT)"],
      ["LuShieldAlert", C.setz, "Angreifbar", "Die Kalibrierung ist der schwache Punkt – im Papier durchgehend so markiert"],
    ];
    for (let i = 0; i < cols.length; i++) {
      const [ic, col, h, b] = cols[i]; const x = 0.6 + i * 3.1;
      s.addShape(pres.shapes.OVAL, { x, y: 2.5, w: 0.7, h: 0.7, fill: { color: C.paperAlt }, line: { color: C.paperAlt } });
      s.addImage({ data: await icon(ic, col), x: x + 0.17, y: 2.67, w: 0.36, h: 0.36 });
      s.addText(h, { x, y: 3.35, w: 2.8, h: 0.4, fontFace: FH, fontSize: 20, bold: true, color: col, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x, y: 3.85, w: 2.8, h: 1.2, fontFace: FB, fontSize: 13, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    fazit(s, "Der Anspruch: ", "Ob eine KI-gestützte Analyse taugt, entscheidet sich nicht an ihrer Entstehung, sondern daran, ob die Zahlen halten.", C.setz);
    s.addNotes("Eigene Einordnung als Newcomer und Nicht-Fachmann offen ansprechen — das ist der Grund, warum es hochgezogene Augenbrauen gibt, skeptische wie beeindruckte, und beides ist gewollt. Kernbotschaft: Die Methodik macht das Papier nicht glaubwürdiger, sondern angreifbarer — und genau das ist der Anspruch. Wer die drei Fragen beantworten kann, sitzt bei Agora, Fraunhofer IEE/ISE, ifeu, in der Kalkulation eines Stadtwerks — oder in der BAG Energie. Quelle: Abschnitt „Zur Entstehung dieses Papiers“ und Abschnitt 10.");
  }

  // ============ V3-1 — Doppelter Preis ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL I · AUSGANGSLAGE", "Deutschland zahlt das Doppelte – für dieselbe Anlage", 2);
    const x0 = 2.3, W = 6.0, kmax = 65;
    const px = (k) => x0 + W * k / kmax;
    s.addText("Luft-Wasser-Wärmepumpe, komplett installiert (€)", { x: 0.6, y: 1.95, w: 8.2, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const rows = [["Deutschland", 20, 63, "20'000–63'000", C.gas], ["Österreich", 13.5, 20, "13'500–20'000", C.wp], ["Frankreich", 12, 20, "12'000–20'000", C.wp], ["Großbritannien", 10, 14, "10'000–14'000", C.wp]];
    rows.forEach(([n, a, b, lbl, col], i) => {
      const y = 2.5 + i * 0.7;
      s.addText(n, { x: 0.6, y, w: 1.65, h: 0.45, fontFace: FB, fontSize: 12, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: px(a), y: y + 0.08, w: px(b) - px(a), h: 0.3, fill: { color: col }, line: { color: col } });
      s.addText(lbl, { x: px(b) + 0.1, y, w: 1.8, h: 0.45, fontFace: FB, fontSize: 11, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
    });
    s.addShape(pres.shapes.LINE, { x: px(14), y: 2.4, w: 0, h: 3.0, line: { color: C.setz, width: 1.25, dashType: "dash" } });
    s.addText("Referenz 14'000 netto · 16'660 brutto", { x: px(14) + 0.1, y: 5.35, w: 4, h: 0.3, fontFace: FB, fontSize: 10, bold: true, color: C.setz, isTextBox: true, margin: 0 });
    s.addText("Nach Förderung", { x: 9.6, y: 1.95, w: 3.1, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    s.addText("3'200 €", { x: 9.6, y: 2.4, w: 3.1, h: 0.7, fontFace: FH, fontSize: 36, bold: true, color: C.wp, isTextBox: true, margin: 0 });
    s.addText("britischer Haushalt", { x: 9.6, y: 3.1, w: 3.1, h: 0.3, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addText("14'900 €", { x: 9.6, y: 3.7, w: 3.1, h: 0.7, fontFace: FH, fontSize: 36, bold: true, color: C.gas, isTextBox: true, margin: 0 });
    s.addText("deutscher Haushalt, 50 % BEG", { x: 9.6, y: 4.4, w: 3.1, h: 0.3, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addText("Faktor 2 bis 2.5 vor, fast 5 nach Förderung. Kein Naturgesetz.", { x: 0.6, y: 6.2, w: 12.1, h: 0.45, fontFace: FH, fontSize: 16, italic: true, color: C.ink, isTextBox: true, margin: 0 });
    s.addNotes("Verbraucherzentrale RLP 2025 (DE: Spanne 20'000–63'000, Median 35'000); HuG, Octopus-RWTH (UK, FR, AT). Referenz 14'000 € netto = Mitte UK/AT-Korridor → 16'660 € mit 19 % USt = Referenzfall in Teil II. Eigenanteil nach Förderung: UK ~3'200, DE ~14'900.");
  }

  // ============ V3-2 — Treiber und Hebel ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL I · DIAGNOSE UND REFORM", "Vier Kostentreiber, fünf Hebel, ein Systemfehler", 3);
    s.addText("Mehrkosten gegenüber UK", { x: 0.6, y: 1.9, w: 5.6, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const drv = [
      ["LuPackage", "Geräte & Markt", "+2'000–4'000 €", "Premiumgeräte"],
      ["LuHardHat", "Fundament", "+1'000–2'200 €", "Normenwesen"],
      ["LuPlug", "Elektroanschluss", "+1'500–2'500 €", "Heizstab, § 14a"],
      ["LuClock", "Arbeitszeit", "+1'000–2'000 €", "110 h statt 40 h"],
      ["LuPercent", "Prozentuale Förderung", "× alles", "Moral Hazard"],
      ["LuReceipt", "Mehrwertsteuer", "+5'700 €", "UK: 0 %"],
    ];
    for (let i = 0; i < drv.length; i++) {
      const [ic, h, k, b] = drv[i]; const y = 2.3 + i * 0.56; const hot = i === 4;
      if (hot) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: y - 0.04, w: 5.8, h: 0.5, rectRadius: 0.06, fill: { color: C.redBg }, line: { color: C.gas, width: 1 } });
      s.addShape(pres.shapes.OVAL, { x: 0.65, y: y + 0.02, w: 0.38, h: 0.38, fill: { color: hot ? C.gas : C.paperAlt }, line: { color: hot ? C.gas : C.paperAlt } });
      s.addImage({ data: await icon(ic, hot ? C.white : C.ink), x: 0.73, y: y + 0.1, w: 0.22, h: 0.22 });
      s.addText(h, { x: 1.15, y, w: 2.2, h: 0.42, fontFace: FB, fontSize: 12.5, bold: true, color: hot ? C.gas : C.ink, isTextBox: true, margin: 0, valign: "middle" });
      s.addText(k, { x: 3.3, y, w: 1.4, h: 0.42, fontFace: FB, fontSize: 12.5, bold: true, color: hot ? C.gas : C.ink, align: "right", isTextBox: true, margin: 0, valign: "middle" });
      s.addText(b, { x: 4.8, y, w: 1.5, h: 0.42, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0, valign: "middle" });
    }
    s.addText("Fünf Reformhebel", { x: 6.9, y: 1.9, w: 5.8, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 2.22, w: 6.0, h: 1.2, rectRadius: 0.06, fill: { color: C.yellowBg }, line: { color: C.yellowBg } });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.68, y: 2.22, w: 0.06, h: 1.2, fill: { color: C.setz }, line: { color: C.setz } });
    s.addText("gleichzeitig drehen", { x: 10.9, y: 2.25, w: 1.8, h: 0.22, fontFace: FB, fontSize: 9, bold: true, color: C.setz, align: "right", isTextBox: true, margin: 0 });
    const heb = [
      ["Festbetrag statt Prozent-Förderung", "★★★"],
      ["Safe Harbour: Normen entkoppeln", "★★★"],
      ["Heizstab ≤ 3 kW als Standard", "★★"],
      ["§ 14a: Tarif statt Hardware-Pflicht", "★★"],
      ["EU-Nullsteuersatz für Heiztechnik", "★★"],
    ];
    for (let i = 0; i < heb.length; i++) {
      const [h, st] = heb[i]; const y = 2.3 + i * 0.6;
      s.addText(String(i + 1), { x: 6.9, y, w: 0.35, h: 0.42, fontFace: FH, fontSize: 18, bold: true, color: C.wp, isTextBox: true, margin: 0, valign: "middle" });
      s.addText([{ text: h + "   ", options: { bold: true, color: C.ink } }, { text: st, options: { color: C.setz, fontSize: 10 } }], { x: 7.3, y, w: 5.2, h: 0.42, fontFace: FB, fontSize: 14, isTextBox: true, margin: 0, valign: "middle" });
    }
    s.addNotes("Treiber: Geräte/Markt +2'000–4'000 (Premium-Bias); Fundament +1'000–2'200 (VDI 4645 faktisch bindend); Elektro +1'500–2'500 (400 V wegen Heizstab, § 14a-Hardware); Arbeitszeit +1'000–2'000 (110 h DE vs ~40 h UK); prozentuale Förderung multipliziert alles (Moral Hazard); USt +5'700 vs UK. ~70 % quantifizierbar. Hebel 1+2 gleichzeitig: Förderdesign, Haftung, Normenwesen – Einzelmaßnahmen verpuffen. H1 Vorbild UK Boiler Upgrade Scheme, FR einkommensgestaffelt. H2: DIN finanziert ~59 % aus Normenverkauf, Koalitionsvertrag 2025 kündigt Rückführung an. H3: Heizstab läuft in 1.9 % der Heizarbeit (ISE), unter § 14a-Schwelle 4.2 kW. H4: Vorbild PAS 1878/1879, OpenADR. H5: Anhang III Nr. 22 MwSt-Richtlinie, Einstimmigkeit.");
  }

  // ============ V3-3 — Gas-Warnung ============
  {
    const s = pres.addSlide();
    frame(s, "WARNUNG · WARUM TEIL I DRINGEND IST", "Solange die Wärmepumpe 30'000 € kostet, rechnet sich nur Gas*", 4,
      GAS_SUB);
    s.addChart(pres.charts.BAR, [{ name: "Vollkosten", labels: ["Gaskessel", "WP heute (DE)", "WP Referenz (EU)"], values: [2151, 2373, 1775] }], {
      ...chartBase(), chartColors: [C.gas, C.grey, C.wp], barDir: "col", barGapWidthPct: 70,
      x: 0.6, y: 2.05, w: 6.6, h: 3.7, title: "Vollkosten Wärme, Einfamilienhaus (ct/kWh)", titleFontSize: 13, dataLabelFontSize: 14, catAxisLabelFontSize: 12,
      dataLabelFormatCode: FMT_DEC, valAxisMinVal: 0, valAxisMaxVal: 2800,
    });
    s.addShape(pres.shapes.OVAL, { x: 7.9, y: 2.3, w: 0.6, h: 0.6, fill: { color: C.redBg }, line: { color: C.redBg } });
    s.addImage({ data: await icon("LuFlame", C.gas), x: 8.03, y: 2.43, w: 0.34, h: 0.34 });
    s.addText("Alle rechnen gegen „Gastherme forever“", { x: 8.65, y: 2.25, w: 4.1, h: 0.7, fontFace: FB, fontSize: 15, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
    s.addText("Haushalte wie Planer: heutiger Gaspreis, Kessel für immer – und die Wette, dass 2045 verschoben wird.", { x: 8.65, y: 3.0, w: 4.1, h: 0.9, fontFace: FB, fontSize: 12, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    s.addText("Investition: Kessel 9'000 € · WP 30'000 € (DE) · 16'660 € (Referenz)", { x: 8.65, y: 4.1, w: 4.1, h: 0.6, fontFace: FB, fontSize: 10.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    fazit(s, "Was das heißt: ", "Teil I entscheidet, ob sich überhaupt etwas anderes als Gas rechnet – die Wärmepumpe wie die Fernwärme. Ohne Wirtschaftlichkeit heute kein Hochlauf, kein Momentum.", C.gas);
    s.addNotes(GAS_NOTES);
  }

  // ============ V3-4 — Systemfrage ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · DIE SYSTEMFRAGE", "Erst der Sockel, dann die Dichte", 5);
    s.addText("Wärmegestehungskosten vor dem ersten Meter Netz (ct/kWh, Vollkosten)", { x: 0.6, y: 1.95, w: 8.4, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const x0 = 3.5, W = 5.0, cmax = 24, ref = 17.75;
    const px = (c) => x0 + W * c / cmax;
    const src = [["Abwärme / Müllverbrennung", 10.02, "+7.72", true], ["Großwärmepumpe / Geothermie", 16.23, "+1.51", true], ["Gas-KWK, fossil", 20.30, "−2.55", false]];
    src.forEach(([n, v, b, pos], i) => {
      const y = 2.5 + i * 0.8;
      const col = n.startsWith("Gas") ? C.gas : C.fw;
      const a = Math.min(px(v), px(ref)), w = Math.abs(px(ref) - px(v));
      s.addShape(pres.shapes.RECTANGLE, { x: a, y: y + 0.12, w, h: 0.26, fill: { color: pos ? C.greenBg : C.redBg }, line: { color: pos ? C.green : C.gas, width: 0.5 } });
      s.addText(n, { x: 0.6, y, w: 2.8, h: 0.5, fontFace: FB, fontSize: 12, color: C.ink, isTextBox: true, margin: 0, valign: "middle", align: "right" });
      s.addShape(pres.shapes.RECTANGLE, { x: x0, y: y + 0.08, w: px(v) - x0, h: 0.34, fill: { color: col }, line: { color: col } });
      s.addText(v.toFixed(2), { x: px(v) - 0.85, y, w: 0.8, h: 0.5, fontFace: FB, fontSize: 12, bold: true, color: C.white, align: "right", isTextBox: true, margin: 0, valign: "middle" });
      s.addText(b + " ct", { x: 8.9, y, w: 1.3, h: 0.5, fontFace: FH, fontSize: 18, bold: true, color: pos ? C.green : C.gas, isTextBox: true, margin: 0, valign: "middle" });
    });
    s.addShape(pres.shapes.LINE, { x: px(ref), y: 2.35, w: 0, h: 2.6, line: { color: C.wp, width: 2.5 } });
    s.addText("Wärmepumpe 17.75", { x: px(ref) - 1.0, y: 4.95, w: 2.0, h: 0.3, fontFace: FB, fontSize: 11, bold: true, color: C.wp, align: "center", isTextBox: true, margin: 0 });
    s.addText("Budget fürs Netz", { x: 8.9, y: 2.1, w: 1.6, h: 0.3, fontFace: FB, fontSize: 10, bold: true, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addText("Gasvorlauf: verschlechtert jede Quelle.", { x: 0.6, y: 5.3, w: 8, h: 0.3, fontFace: FB, fontSize: 11, italic: true, color: C.gas, isTextBox: true, margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.4, y: 1.95, w: 2.3, h: 3.6, rectRadius: 0.08, fill: { color: C.tealBg }, line: { color: C.tealBg } });
    s.addText([
      { text: "Budget", options: { bold: true, fontSize: 14, color: C.wp, breakLine: true } },
      { text: "= Wärmepumpe − Sockel", options: { bold: true, breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 6 } },
      { text: "Was das Netz je kWh für Trasse, Anschluss, Verluste ausgeben darf.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 6 } },
      { text: "Negativ = verloren vor dem ersten Meter.", options: { bold: true, color: C.gas } },
    ], { x: 10.55, y: 2.1, w: 2.0, h: 3.3, fontFace: FB, fontSize: 11.5, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
    fazit(s, "Die Setzung: ", "Nicht die Dichte ist das Axiom, sondern das Vorzeichen des Budgets. Erst danach fragt das Modell, ob die Bebauung reicht.", C.wp);
    s.addNotes("Abschnitt 4.3. Sockel = Erzeugung, Speicher, Pumpstrom, Netzverluste, Betrieb – Vollkosten, Annuität, ohne einen Meter Verteilnetz. Gegen WP-Referenz 17.75 → Budget: Abwärme +7.72, GroßWP/Geo +1.51, Gas-KWK fossil −2.55. Gasvorlauf: barwertgewichtet mit Nachfolgequelle, kostet 5.7 ct vor Abwärme (Budget bleibt +2.0) und 2.3 ct vor GroßWP (Budget −0.8, nie). Grünes Gas müsste ab 2045 für 6.4 ct/kWh liefern (Befund 2).");
  }

  // ============ V3-5 — Methode ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · METHODE", "Drei Setzungen und eine Regel", 6);
    const meth = [
      ["LuCalculator", "Vollkosten beider Pfade", "gleiche Systemgrenze, Annuität, ohne Förderung"],
      ["LuTarget", "Wärmepumpe auf EU-Niveau", "16'660 € → 17.75 ct/kWh"],
      ["LuGauge", "Dichteschwelle berechnet, nicht gesetzt", "aus Trassenkosten und Budget"],
    ];
    for (let i = 0; i < meth.length; i++) {
      const [ic, h, b] = meth[i]; const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.OVAL, { x, y: 1.95, w: 0.6, h: 0.6, fill: { color: C.greenBg }, line: { color: C.greenBg } });
      s.addImage({ data: await icon(ic, C.green), x: x + 0.14, y: 2.09, w: 0.32, h: 0.32 });
      s.addText(h, { x, y: 2.65, w: 3.8, h: 0.5, fontFace: FB, fontSize: 15, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x, y: 3.15, w: 3.8, h: 0.5, fontFace: FB, fontSize: 12, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    // Regel als Grafik
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 3.9, w: 12.1, h: 2.2, rectRadius: 0.08, fill: { color: C.blueBg }, line: { color: C.blueBg } });
    s.addText("Wärmeliniendichte: wie viel Wärme ein Meter Straße im Jahr verkauft (MWh/m·a)", { x: 0.8, y: 4.0, w: 11.7, h: 0.35, fontFace: FB, fontSize: 12, color: C.fw, isTextBox: true, margin: 0 });
    const parts = [["erreichte Dichte", C.fw], [">", C.ink], ["Netzkosten je Meter und Jahr", C.setz], ["÷", C.ink], ["Budget", C.wp], ["→ Netz trägt", C.green]];
    let x = 0.8;
    const widths = [2.6, 0.45, 3.35, 0.45, 1.5, 2.45];
    parts.forEach(([t, col], i) => {
      const isOp = t === ">" || t === "÷";
      if (!isOp && i < 5) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.5, w: widths[i], h: 0.85, rectRadius: 0.1, fill: { color: C.white }, line: { color: col, width: 1.5 } });
      s.addText(t, { x, y: 4.5, w: widths[i], h: 0.85, fontFace: FH, fontSize: isOp ? 26 : 17, bold: true, color: col, align: "center", isTextBox: true, margin: 0, valign: "middle" });
      x += widths[i] + 0.15;
    });
    // Klammer: das ist der berechnete Schwellenwert
    const bx = 0.8 + widths[0] + 0.15 + widths[1] + 0.15;
    const bw = widths[2] + 0.15 + widths[3] + 0.15 + widths[4];
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: 5.38, w: bw, h: 0.02, fill: { color: C.setz }, line: { color: C.setz } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: 5.32, w: 0.02, h: 0.06, fill: { color: C.setz }, line: { color: C.setz } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw, y: 5.32, w: 0.02, h: 0.06, fill: { color: C.setz }, line: { color: C.setz } });
    s.addShape(pres.shapes.RECTANGLE, { x: bx + bw / 2, y: 5.38, w: 0.02, h: 0.08, fill: { color: C.setz }, line: { color: C.setz } });
    s.addText("= Schwellenwert, den das Modell ausrechnet", { x: bx - 1.0, y: 5.46, w: bw + 2.0, h: 0.28, fontFace: FB, fontSize: 11.5, bold: true, color: C.setz, align: "center", isTextBox: true, margin: 0 });
    s.addText("Kleines Budget → hohe Schwelle. 36 Parameterkombinationen, Skript offen.", { x: 0.8, y: 5.78, w: 11.7, h: 0.3, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addNotes("Annuität VDI 2067, 3 % Realzins, Ersatzinvestitionen eingepreist. Trassenkosten AGFW 677 / 1'692 / 3'554 €/m (unbefestigt / Bestandsstraße / Innenstadt), Baupreisindex bis Mai 2026. Fünf Gebäudetypen, Netzausbau und vorhandenes Netz. Typische Dichten: EFH 0.3–0.8, Reihenhaus 0.8–1.5, Blockrand 2–4, Innenstadt 3–6; Faustregel Planung ab ~1.5. Bei Anschlussquote 60 % statt 100 % steigen alle Schwellen um Faktor 1.67.");
  }

  // ============ V3-6 — Wärmequelle entscheidet ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · BEFUND 1", "Die Wärmequelle entscheidet, nicht die Dichte", 7);
    const cats = ["EFH-Gebiet", "Reihenhaus", "Kleines MFH", "Gründerzeitblock", "Plattenbau-Zeile"];
    s.addChart(pres.charts.BAR, [
      { name: "erreichte Wärmeliniendichte", labels: cats, values: [86, 214, 391, 744, 423] },
      { name: "Schwelle Abwärmenetz", labels: cats, values: [107, 108, 108, 245, 133] },
    ], {
      ...chartBase(), x: 0.5, y: 1.9, w: 12.3, h: 3.9, barDir: "col", barGapWidthPct: 45, barGrouping: "clustered",
      chartColors: [C.fw, "D9A441"], showLegend: true, legendPos: "t", legendFontSize: 12,
      title: "Netzausbau, MWh je Trassenmeter und Jahr", titleFontSize: 13, dataLabelFontSize: 12, catAxisLabelFontSize: 12,
      dataLabelFormatCode: FMT_DEC, valAxisMinVal: 0, valAxisMaxVal: 850,
    });
    fazit(s, "Abwärme trägt fast überall, Großwärmepumpe nirgends: ", "Schwelle GroßWP/Geothermie 5.4 · 5.6 · 5.8 · 20 · >50 – kein Quartier erreicht sie.", C.fw);
    s.addNotes("Tabelle 5.2. Erreichte Dichte aus Bebauungsgeometrie: EFH 17 m Trasse → 0.86; RH 5 m → 2.14; kl. MFH 9 m → 3.91; Gründerzeit 12 m → 7.44 (Innenstadtstraße); Platte 80 m → 4.23. Abwärme-Schwellen 1.07 / 1.08 / 1.08 / 2.45 / 1.33 – EFH knapp, kippt bei 13.7 statt 17 m. In keiner von 36 Kombinationen dreht sich das Abwärme-Ergebnis. GroßWP/Geo: Befund 3, unentschieden – Vorzeichen wechselt in einem Sechstel der Kombinationen; Reserve R2.");
  }

  // ============ V3-7 — Heatmap + Kernbotschaft ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · BEFUND 4", "Der Vergleichsmaßstab erzeugt einen ganzen Pfad", 8);
    HEATMAP_TABLE(s);
    s.addText([
      { text: "Grün: ", options: { bold: true, color: "3B6D11" } }, { text: "das Netz trägt.   " },
      { text: "Blau: ", options: { bold: true, color: "185FA5" } }, { text: "die Wärmepumpe gewinnt.   " },
      { text: "×: ", options: { bold: true, color: "042C53" } }, { text: "negatives Budget – bei keiner Dichte erreichbar.   Je kräftiger der Ton, desto eindeutiger." },
    ], { x: 0.6, y: 4.85, w: 12.1, h: 0.3, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addText("Faktor 1.8 bis 5 zwischen links und rechts. Der Hochlauf verschärft das.", { x: 0.6, y: 5.25, w: 12.1, h: 0.4, fontFace: FH, fontSize: 15, italic: true, color: C.ink, isTextBox: true, margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.95, w: 12.1, h: 0.9, rectRadius: 0.08, fill: { color: C.paperAlt }, line: { color: C.rule, width: 0.75 } });
    s.addText([
      { text: "Die Kernbotschaften:", options: { bold: true, color: C.gas, breakLine: true } },
      { text: "– ", options: { bold: true, color: C.gas } },
      { text: "Das deutsche WP-Preisniveau subventioniert die Fernwärme – und macht beides so teuer, dass viele sich aktuell noch für fossiles Gas entscheiden.", options: { breakLine: true } },
      { text: "– Senken wir WP-Preise, könnten manche Fernwärmenetze zu „sunk costs“ werden oder einen politisch teuren Anschluss- und Benutzungszwang erzwingen." },
    ], { x: 0.8, y: 5.95, w: 11.7, h: 0.9, fontFace: FB, fontSize: 12.5, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
    s.addNotes(HEAT_NOTES);
  }

  // ============ V3-8 — Gebäudetyp ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · BEFUND 5", "1 Mrd. € = 26'000 oder 413'000 Haushalte, je nach Gebäudetyp", 9);
    const cats = ["WP-Festbetrag 9'000 €", "EFH-Gebiet", "Reihenhausgebiet", "Kleines MFH, 5 WE", "Gründerzeitblock, 12 WE", "Plattenbau-Zeile, 60 WE"];
    s.addChart(pres.charts.BAR, [{ name: "WE", labels: cats, values: [111000, 26000, 54000, 198000, 228000, 413000] }], {
      ...chartBase(), x: 1.4, y: 1.9, w: 11.3, h: 3.9, barDir: "bar", barGapWidthPct: 40, catAxisOrientation: "maxMin",
      chartColors: [C.wp, C.fw, C.fw, C.fw, C.fw, C.fw],
      title: "Wohneinheiten, die 1 Mrd. € umstellt", titleFontSize: 13, dataLabelFontSize: 13, catAxisLabelFontSize: 12, catAxisLabelColor: C.ink,
      dataLabelFormatCode: FMT_TSD, valAxisMinVal: 0, valAxisMaxVal: 480000,
    });
    s.addShape(pres.shapes.LEFT_BRACKET, { x: 1.05, y: 3.0, w: 0.15, h: 2.65, line: { color: C.fw, width: 1.5 } });
    s.addText("Netzausbau", { x: 0.1, y: 4.1, w: 1.5, h: 0.4, fontFace: FB, fontSize: 11, bold: true, color: C.fw, isTextBox: true, margin: 0, valign: "middle", align: "center", rotate: 270 });
    fazit(s, "Was das heißt: ", "Im EFH-Bestand ist Netzausbau die teuerste, im Geschosswohnungsbau die billigste Mittelverwendung. Ein Schwellenwert für die ganze Stadt entscheidet in beide Richtungen falsch.", C.fw);
    s.addNotes("Abschnitt 6, Tabelle 5 des Skripts. Alle Netzbalken: 1 Mrd. € als Netzausbau in genau diesen Quartierstyp. Je WE: 9'000 / 38'770 / 18'460 / 5'050 / 4'390 / 2'420 €. Faktoren zum Festbetrag: EFH 4.3× (3.6–6.8), RH 2.1×, kl. MFH 0.56×, Gründerzeit 0.49× (0.39–0.68), Platte 0.27×. Fairness: Festbetrag hebelt privates Kapital; Rechnung zählt Haushalte in Bewegung je Euro. Prognos/VKU 2024: 74 Mrd. € bis 2045, 3.4–3.5 Mrd. €/a Förderbedarf – Verbandsgutachten. Befund 6 (vorhandenes Netz, ~6 Mio. WE): Trasse bezahlt, Frage ist Quellenwechsel; mit Abwärme gewinnt der Anschluss in jedem Gebäudetyp.");
  }

  // ============ V12 — Politik + Instrumente ============
  {
    const s = pres.addSlide();
    frame(s, "SCHLUSSFOLGERUNGEN", "Was daraus folgt", 10);
    s.addText("Drei Botschaften", { x: 0.6, y: 1.85, w: 6.2, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.inkSoft, charSpacing: 1, isTextBox: true, margin: 0 });
    const msgs = [
      [C.gas, "Anschluss- und Benutzungszwang ist ein politisches Risiko.", "Je besser die WP-Reform wirkt, desto öfter zwingt er Bürger ans teurere Netz – oder das Netz wird zum stranded asset."],
      [C.gas, "Der WP-Preis ist das einzige Preissignal im Wärmesektor.", "Das deutsche Preisniveau subventioniert die Fernwärme – und treibt Haushalte zu fossilem Gas."],
      [C.setz, "2045 ist nichts wert, solange nicht gebaut wird.", "Alle wetten auf ein Verschieben des Ziels. Nur Wirtschaftlichkeit heute erzeugt Hochlauf."],
    ];
    for (let i = 0; i < msgs.length; i++) {
      const [col, h, b] = msgs[i]; const y = 2.3 + i * 1.25;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.04, w: 0.08, h: 1.0, fill: { color: col }, line: { color: col } });
      s.addText(h, { x: 0.9, y, w: 5.9, h: 0.5, fontFace: FB, fontSize: 13.5, bold: true, color: col, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x: 0.9, y: y + 0.5, w: 5.9, h: 0.65, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    s.addText("Drei Instrumente", { x: 7.2, y: 1.85, w: 5.5, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.inkSoft, charSpacing: 1, isTextBox: true, margin: 0 });
    const inst = [
      ["LuEuro", "BEG-Festbetrag + Safe Harbour", "= Fernwärme-Preisdisziplin. Kapital ist nicht der Engpass: Der Bund bürgt schon für Wärmenetze – niemand ruft es ab.", "Teil I · 9.5"],
      ["LuCircleHelp", "Benchmark in der Wärmeplanung", "nach Gebäudetyp, gegen die EU-Referenz – kein Gebietsmittel", "WPG § 18 · 9.1"],
      ["LuTarget", "Förderung quellenscharf", "Abwärme zuerst, Gasvorlauf wie seine Nachfolgequelle, im Bestand Quellenwechsel", "BEW · 9.3"],
    ];
    for (let i = 0; i < inst.length; i++) {
      const [ic, h, b, src] = inst[i]; const y = 2.3 + i * 1.25;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y, w: 5.5, h: 1.05, rectRadius: 0.06, fill: { color: i === 0 ? C.tealBg : C.white }, line: { color: i === 0 ? C.wp : C.rule, width: 0.75 } });
      s.addText(String(i + 1), { x: 7.3, y, w: 0.4, h: 1.05, fontFace: FH, fontSize: 22, bold: true, color: C.wp, isTextBox: true, margin: 0, valign: "middle", align: "center" });
      s.addImage({ data: await icon(ic, C.wp), x: 7.78, y: y + 0.38, w: 0.3, h: 0.3 });
      s.addText(h, { x: 8.2, y: y + 0.12, w: 4.4, h: 0.4, fontFace: FB, fontSize: 13.5, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
      s.addText(b, { x: 8.2, y: y + 0.52, w: 4.4, h: 0.5, fontFace: FB, fontSize: 10.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
      s.addText(src, { x: 11.4, y: y + 0.05, w: 1.2, h: 0.25, fontFace: FB, fontSize: 8.5, color: C.setz, align: "right", isTextBox: true, margin: 0 });
    }
    s.addNotes("Botschaften: Anschlusszwang 9.7 (stranded asset vs. Zwang gegen Bürgerpräferenz, beides wahrscheinlicher, je besser Teil I wirkt); Preissignal 9.5 (Kartellamt: Fernwärme faktisch nicht regulierbar); 2045: Gas-Folie und Hintergrundgespräche. Nicht auf der Folie, in Reserve: falscher Maßstab in Wärmeplänen (Befund 4), Gasvorlauf als Wette auf Politikversagen (6.4 ct/kWh ab 2045, Befund 2), Schwellenwert für die ganze Stadt entscheidet in beide Richtungen falsch (Befund 5). Instrumente gestrichen, bei Nachfrage: Wärmenetzregister (9.2, AGFW erfasst ~ein Fünftel der Trassenlänge) und öffentliche Kostenbenchmark Trassenbau (9.4 – Papier rechnet mit AGFW-Punktwolken, Preisstand 2021, Baupreisindex). Was die BAG damit tun könnte: nicht vorgeben – R1 nur auf Nachfrage.");
  }

  // ============ Linkfolie (Seite 11) ============
  {
    const s = pres.addSlide();
    frame(s, "ZUM NACHLESEN", "Das vollständige Papier", 11);
    s.addText("Alle Zahlen dieses Vortrags mit Herleitung, Quellen und Interessenlage,\nelf Anhänge und das vollständige Simulationsskript.\n\nPapier und Folien stehen unter CC BY 4.0, das Skript unter MIT –\nWeiterverwendung mit Namensnennung ausdrücklich erwünscht.",
      { x: 0.6, y: 2.2, w: 6.6, h: 2.0, fontFace: FB, fontSize: 14, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
    s.addText("theflow85.github.io/waermewende", { x: 0.6, y: 3.6, w: 6.6, h: 0.6, fontFace: FB, fontSize: 22, bold: true, color: C.wp, isTextBox: true, margin: 0, valign: "middle" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.9, w: 6.6, h: 1.15, rectRadius: 0.08, fill: { color: C.paperAlt }, line: { color: C.rule, width: 0.75 } });
    s.addText([
      { text: "Drei Fragen ans Fachreview: ", options: { bold: true, color: C.setz } },
      { text: "Erzeugungskosten je Wärmequelle, Verlegekosten je Meter, Trassenmeter je Gebäudetyp. Wer dazu etwas sagen kann: gern melden." },
    ], { x: 0.8, y: 4.9, w: 6.2, h: 1.15, fontFace: FB, fontSize: 12, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
    s.addImage({ path: "/home/claude/qr.png", x: 8.3, y: 2.2, w: 3.4, h: 3.4 });
    s.addText("Kamera aufs Bild halten", { x: 8.3, y: 5.7, w: 3.4, h: 0.3, fontFace: FB, fontSize: 11, color: C.inkSoft, align: "center", isTextBox: true, margin: 0 });
    s.addNotes("Abschlussfolie. Kurz-URL leitet im Repo-Root auf die aktuelle Hash-Fassung weiter; QR zeigt auf dieselbe Kurz-URL. Mündlich: Kritik zu den drei offenen Punkten ausdrücklich erwünscht – Erzeugungskosten je Wärmequelle, Verlegekosten 677/1'692/3'554 €/m, Trassenmeter je Gebäudetyp (17 m EFH, 12 m Gründerzeitblock).");
  }

  await pres.writeFile({ fileName: "/home/claude/bag-deck-v16.pptx" });
  console.log("written");
})();
