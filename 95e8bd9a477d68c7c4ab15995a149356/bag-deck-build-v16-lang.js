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


  // ============ Methodik (Seite 1) ============
  {
    const s = pres.addSlide();
    frame(s, "METHODIK · ZUR ENTSTEHUNG DIESES PAPIERS", "Ein Fachfremder, eine KI und ein offenes Skript", 1);

    const cardW = 3.75;
    const cards = [
      ["LuBot", C.fw, "KI-Sprachmodell (Claude)", ["recherchiert", "Kostenmodelle konstruiert", "Simulation programmiert und gerechnet", "Texte formuliert"]],
      ["LuUserCheck", C.wp, "Autor — Arzt mit Modellierungserfahrung", ["Fragestellung und Analyserichtung gesetzt", "Einwände und Korrekturen eingebracht", "Logik, Konsistenz, Plausibilität geprüft und nachgefragt", "lektoriert"]],
    ];
    for (let i = 0; i < 2; i++) {
      const [ic, col, h, items] = cards[i]; const y = 1.8 + i * 2.0;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: cardW, h: 1.8, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.rule, width: 0.75 } });
      s.addImage({ data: await icon(ic, col), x: 0.75, y: y + 0.18, w: 0.34, h: 0.34 });
      s.addText(h, { x: 1.13, y: y + 0.14, w: 3.1, h: 0.44, fontFace: FB, fontSize: 11.5, bold: true, color: col, isTextBox: true, margin: 0, valign: "middle" });
      s.addText(items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })), { x: 0.8, y: y + 0.66, w: 3.4, h: 1.05, fontFace: FB, fontSize: 10.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 2 });
    }
    s.addText("Ehrenamtliches Zeitbudget, kein Auftraggeber — eine volkswirtschaftliche Vollkostenrechnung aus Bürgerperspektive, die es so bisher nicht gab.",
      { x: 0.6, y: 5.95, w: cardW, h: 0.9, fontFace: FH, fontSize: 11.5, italic: true, color: C.ink, isTextBox: true, margin: 0, valign: "top" });

    s.addText("Was es nachprüfbar macht", { x: 4.75, y: 1.8, w: 4.2, h: 0.35, fontFace: FB, fontSize: 13, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const checks = [
      ["LuTag", "Jede Modellannahme gekennzeichnet, mit Herkunft und Grundlage"],
      ["LuEye", "Jede Quelle mit ihrer Interessenlage versehen"],
      ["LuFileCode", "Simulation als Python-Skript beigelegt — ohne Abhängigkeiten, MIT-Lizenz, lokal nachrechenbar"],
      ["LuCalculator", "Eine Rechnung im Anhang vollständig von Hand durchgerechnet"],
      ["LuShieldAlert", "Robustheit: 36 Parameterkombinationen, Unsicherheitsmarker an jedem der 30 Urteile"],
    ];
    for (let i = 0; i < checks.length; i++) {
      const [ic, t] = checks[i]; const y = 2.25 + i * 0.66;
      s.addShape(pres.shapes.OVAL, { x: 4.75, y: y + 0.02, w: 0.42, h: 0.42, fill: { color: C.greenBg }, line: { color: C.greenBg } });
      s.addImage({ data: await icon(ic, C.green), x: 4.84, y: y + 0.11, w: 0.24, h: 0.24 });
      s.addText(t, { x: 5.3, y, w: 3.65, h: 0.5, fontFace: FB, fontSize: 11, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
    }
    s.addText("Erstfassung Mai 2026 → V2.11.1 (Sept. 2026) · Versionshistorie im Papier dokumentiert · 11 Anhänge · Quellenverzeichnis mit Interessenlage",
      { x: 4.75, y: 5.95, w: 4.2, h: 0.9, fontFace: FB, fontSize: 10.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.35, y: 1.8, w: 3.35, h: 5.05, rectRadius: 0.08, fill: { color: C.yellowBg }, line: { color: C.yellowBg } });
    s.addText("Die verwundbare Seite", { x: 9.55, y: 1.95, w: 3.0, h: 0.35, fontFace: FB, fontSize: 13, bold: true, color: C.setz, isTextBox: true, margin: 0 });
    s.addText("Der methodische Ansatz trägt. Die Kalibrierung ist der Angriffspunkt — und im Papier durchgehend so markiert. Drei Fragen ans Fachreview:",
      { x: 9.55, y: 2.35, w: 3.0, h: 1.0, fontFace: FB, fontSize: 11, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
    const qs = ["Sind die Erzeugungskosten je Wärmequelle plausibel kalibriert?", "Sind 677 / 1'692 / 3'554 €/m Verlegekosten realistisch?", "Sind 17 Trassenmeter je Einfamilienhaus und 12 je Gründerzeitblock realistisch?"];
    s.addText(qs.map((t, k) => ({ text: t, options: { bullet: { type: "number" }, breakLine: k < qs.length - 1 } })),
      { x: 9.55, y: 3.4, w: 3.0, h: 2.2, fontFace: FB, fontSize: 11, color: C.ink, isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 6 });
    s.addText("Ob eine KI-gestützte Analyse taugt, entscheidet sich nicht an ihrer Entstehung, sondern daran, ob die Zahlen halten.",
      { x: 9.55, y: 5.75, w: 3.0, h: 1.0, fontFace: FH, fontSize: 11.5, italic: true, color: C.setz, isTextBox: true, margin: 0, valign: "top" });

    s.addNotes("Eigene Einordnung als Newcomer und Nicht-Fachmann offen ansprechen — das ist der Grund, warum es hochgezogene Augenbrauen gibt, skeptische wie beeindruckte, und beides ist gewollt. Kernbotschaft: Die Methodik macht das Papier nicht glaubwürdiger, sondern angreifbarer — und genau das ist der Anspruch. Wer die drei Fragen beantworten kann, sitzt bei Agora, Fraunhofer IEE/ISE, ifeu, in der Kalkulation eines Stadtwerks — oder in der BAG Energie. Quelle: Abschnitt „Zur Entstehung dieses Papiers“ und Abschnitt 10.");
  }

  // ============ Folie 3 — Ausgangslage (Seite 1) ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL I · AUSGANGSLAGE", "Deutschland zahlt das Doppelte — für dieselbe Anlage", 2);

    const cats = ["UK", "Frankreich", "Österreich", "Deutschland"];
    s.addChart([
      {
        type: pres.charts.BAR,
        data: [{ name: "obere Grenze", labels: cats, values: [14000, 20000, 20000, 63000] }],
        options: { chartColors: [C.grey], barDir: "bar", barGapWidthPct: 60, dataLabelPosition: "outEnd", dataLabelFormatCode: FMT_TSD },
      },
      {
        type: pres.charts.BAR,
        data: [{ name: "untere Grenze", labels: cats, values: [10000, 12000, 13500, 20000] }],
        options: { chartColors: [C.paper], barDir: "bar", barGapWidthPct: 60, dataLabelPosition: "inEnd", dataLabelFormatCode: FMT_TSD },
      },
    ], {
      ...chartBase(), x: 0.6, y: 1.95, w: 7.7, h: 2.95,
      barDir: "bar", barGrouping: "clustered", barOverlapPct: 100, barGapWidthPct: 60,
      title: "Luft-Wasser-Wärmepumpe, komplett installiert (€)",
      catAxisLabelFontSize: 11,
      dataLabelFormatCode: FMT_TSD, dataLabelFontSize: 10,
      valAxisMinVal: 0, valAxisMaxVal: 72000,
    });
    s.addText("DE: Verbraucherzentrale RLP 2025, Median 35'000 € · UK, FR, AT: HuG und Octopus/RWTH — zwei Erhebungen mit unterschiedlicher Methodik, als Bandbreiten lesen.",
      { x: 0.6, y: 4.95, w: 7.7, h: 0.35, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    s.addText([
      { text: "Referenzfall des Papiers: ", options: { bold: true } },
      { text: "14'000 € netto, die Mitte des UK/AT-Korridors — mit 19 % USt 16'660 €. Abstand zu Deutschland: Faktor 2 bis 2.5." },
    ], { x: 0.6, y: 5.35, w: 7.7, h: 0.4, fontFace: FB, fontSize: 11, color: C.wp, isTextBox: true, margin: 0, valign: "top" });

    s.addText("Nach Förderung wird der Abstand größer", { x: 8.65, y: 1.95, w: 4.05, h: 0.35, fontFace: FB, fontSize: 13, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const eig = [["rund 3'200 €", "Eigenanteil eines britischen Haushalts", C.wp], ["rund 14'900 €", "Eigenanteil in Deutschland bei 50 % BEG", C.gas]];
    eig.forEach(([n, l, col], i) => {
      const y = 2.45 + i * 1.15;
      s.addText(n, { x: 8.65, y, w: 4.05, h: 0.55, fontFace: FH, fontSize: 30, bold: true, color: col, isTextBox: true, margin: 0, valign: "bottom" });
      s.addText(l, { x: 8.65, y: y + 0.58, w: 4.05, h: 0.4, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    });
    s.addText("Fast fünfmal so viel — trotz der höheren deutschen Förderquote.",
      { x: 8.65, y: 4.85, w: 4.05, h: 0.6, fontFace: FH, fontSize: 13, italic: true, color: C.ink, isTextBox: true, margin: 0, valign: "top" });

    fazit(s, "Was das heißt: ", "Gleiche Technik, gleicher Gebäudebestand, doppelter Preis — und nach Förderung fünffacher Eigenanteil. Der Preisunterschied ist kein Naturgesetz, sondern gemacht.", C.wp);

    s.addNotes("Einstieg in Teil I. Kernaussage: Der Preisunterschied ist kein Naturgesetz. Zahlen: DE rund 30'000 € komplett installiert (Verbraucherzentrale RLP 2025, Spanne 20'000–63'000 €, Median 35'000 €), UK 10'000–14'000 €, FR 12'000–20'000 €, AT 13'500–20'000 €. Zwei Erhebungen (HuG, Octopus/RWTH) mit unterschiedlicher Methodik — bewusst als Bandbreiten gezeigt, nicht als Punktwerte. Referenzwert des Papiers: 14'000 € netto, mit 19 % USt 16'660 € (Referenzfall Teil II). Der entscheidende Befund steht rechts: nach Förderung zahlt der britische Haushalt rund 3'200 €, der deutsche bei 50 % BEG rund 14'900 € — fast das Fünffache trotz höherer Förderquote. Fundstelle: Teil I, Ausgangslage und Diagnose.");
  }

  // ============ S2 — Treiber und Hebel ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL I · DIAGNOSE UND REFORM", "Vier Kostentreiber, fünf Hebel, ein Systemfehler", 3);
    s.addText("Mehrkosten gegenüber Großbritannien", { x: 0.6, y: 1.8, w: 5.6, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const drv = [
      ["LuPackage", "Geräte & Markt", "+2'000–4'000 €", "Premium-Bias; günstige Varianten werden in DE nicht angeboten"],
      ["LuHardHat", "Fundament / Aufstellung", "+1'000–2'200 €", "VDI 4645 faktisch verbindlich, technisch oft übertrieben"],
      ["LuPlug", "Elektroanschluss / § 14a", "+1'500–2'500 €", "400 V wegen überdimensionierter Heizstäbe; DSR-Hardwarepflicht"],
      ["LuClock", "Lohn & Arbeitszeit", "+1'000–2'000 €", "110 Monteurstunden (DE) gegen ~40 (UK): Bürokratie-Overhead"],
      ["LuPercent", "Prozentuale Förderung", "× alle obigen", "Moral Hazard: der Endkunde vergleicht nicht, der Betrieb rechnet an die Förderobergrenze"],
      ["LuReceipt", "Mehrwertsteuer", "+5'700 €", "gegenüber UK (0 %); gering gegenüber FR, NL, AT"],
    ];
    for (let i = 0; i < drv.length; i++) {
      const [ic, h, k, b] = drv[i]; const y = 2.2 + i * 0.62; const hot = i === 4;
      if (hot) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: y - 0.05, w: 5.8, h: 0.6, rectRadius: 0.06, fill: { color: C.redBg }, line: { color: C.gas, width: 1 } });
      s.addShape(pres.shapes.OVAL, { x: 0.65, y: y + 0.05, w: 0.42, h: 0.42, fill: { color: hot ? C.gas : C.paperAlt }, line: { color: hot ? C.gas : C.paperAlt } });
      s.addImage({ data: await icon(ic, hot ? C.white : C.ink), x: 0.74, y: y + 0.14, w: 0.24, h: 0.24 });
      s.addText(h, { x: 1.2, y, w: 2.1, h: 0.28, fontFace: FB, fontSize: 11.5, bold: true, color: hot ? C.gas : C.ink, isTextBox: true, margin: 0, valign: "middle" });
      s.addText(k, { x: 3.3, y, w: 1.3, h: 0.28, fontFace: FB, fontSize: 11.5, bold: true, color: hot ? C.gas : C.ink, align: "right", isTextBox: true, margin: 0, valign: "middle" });
      s.addText(b, { x: 1.2, y: y + 0.27, w: 5.0, h: 0.28, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    s.addText("~70 % der Differenz quantifizierbar, ~30 % aus sich verstärkenden Faktoren.", { x: 0.6, y: 5.98, w: 5.7, h: 0.3, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0 });

    // Gruppierung Hebel 1+2
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 2.1, w: 6.0, h: 1.55, rectRadius: 0.06, fill: { color: C.yellowBg }, line: { color: C.yellowBg } });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.68, y: 2.1, w: 0.06, h: 1.55, fill: { color: C.setz }, line: { color: C.setz } });
    s.addText("gleichzeitig drehen: Förderdesign · Haftung · Normenwesen", { x: 4.95, y: 2.72, w: 3.2, h: 0.3, fontFace: FB, fontSize: 8, bold: true, color: C.setz, align: "center", isTextBox: true, margin: 0, valign: "middle", rotate: 270 });
    // rechts: fünf Hebel
    s.addText("Fünf Reformhebel", { x: 6.9, y: 1.8, w: 5.8, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const heb = [
      ["★★★", "BEG: Prozentuale Förderung auf Festbeträge umstellen", "8'000–10'000 €, einkommensgestaffelt (FR); Vorbild UK Boiler Upgrade Scheme · Gesetz, kurzfristig"],
      ["★★★", "Safe Harbour für Komfort-Normen, BEG von VDI-Zertifikaten entkoppeln", "Normen sind formal freiwillig, über Haftung, BAFA und Versicherung faktisch bindend; DIN finanziert ~59 % aus Normenverkauf · mittelfristig"],
      ["★★", "Heizstab ≤ 3 kW als Förderstandard", "Läuft bei korrekter Auslegung in 1.9 % der Heizarbeit (Fraunhofer ISE); unterschreitet die § 14a-Schwelle von 4.2 kW · kurzfristig"],
      ["★★", "§ 14a: Tarifanreize statt Hardware-Pflichten", "Steuerbox/SMGW/EEBus je 1'500–2'500 €; Vorbild PAS 1878/1879, OpenADR · EnWG-Novelle, mittelfristig"],
      ["★★", "EU: Nullsatz auch für dezentrale Heiztechnik", "Anhang III Nr. 22 MwSt-Systemrichtlinie; Einstimmigkeit nach Art. 113 AEUV · langfristig"],
    ];
    for (let i = 0; i < heb.length; i++) {
      const [st, h, b] = heb[i]; const y = 2.2 + i * 0.76;
      s.addText(String(i + 1), { x: 6.9, y, w: 0.3, h: 0.3, fontFace: FH, fontSize: 16, bold: true, color: C.wp, isTextBox: true, margin: 0, valign: "top" });
      s.addText([{ text: h + "  ", options: { bold: true, color: C.ink } }, { text: st, options: { color: C.setz, fontSize: 9 } }], { x: 7.25, y, w: 5.0, h: 0.3, fontFace: FB, fontSize: 11, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x: 7.25, y: y + 0.3, w: 5.0, h: 0.44, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    s.addNotes("Tabelle Teil I, Diagnose. Rund 70 % der Preisdifferenz quantifizierbar (Geräte, Fundament, Elektro, Lohn, USt), 30 % schwer greifbar und sich verstärkend. Die prozentuale Förderung ist der stärkste Einzeltreiber, weil sie alle anderen multipliziert – deshalb ist Hebel 1 (Festbetrag) auch der stärkste Einzelhebel. Systemlogik (Teil I, letzter Abschnitt): Prozentuale Förderung bei gleichzeitiger Normenverbindlichkeit nimmt jede Motivation zur Kostendisziplin; der Kunde zahlt 30–70 % weniger, der Rest kommt vom Staat, und der Handwerker haftet, wenn er die günstigere Lösung wählt. Reform muss an drei Punkten gleichzeitig ansetzen: Haftung (Safe Harbour), Förderdesign (Festbetrag), Normenwesen (Entkoppelung von der Fördervoraussetzung). Einzelmaßnahmen verpuffen. Zu Hebel 2: Koalitionsvertrag 2025 kündigt an, die Bindungswirkung von Normsetzungen auf ein sicherheitsrelevantes Maß zurückzuführen – Umsetzung steht aus. Hebel 5 steht nachrangig, behebt aber als einziger die Ungleichbehandlung PV/WP an der Wurzel.");
  }

  // ============ Gas-Warnung (Seite 3) ============
  {
    const s = pres.addSlide();
    frame(s, "WARNUNG · WARUM TEIL I DRINGEND IST", "Solange die Wärmepumpe 30'000 € kostet, rechnet sich nur Gas*", 4,
      "* Genauer: Vollfaktoriell über 27 Kombinationen (Kesseltausch 6'000/9'000/12'000 € · Gas 9/11/13 ct/kWh · Nutzungsgrad 85/88/92 %) liegt der Kessel in 22 von 27 unter dem heutigen deutschen WP-Preis von 23.73 ct und in 1 von 27 unter der EU-Referenz von 17.75 ct — bei 6'000 €, 9 ct und 92 % mit 17.35 ct/kWh. Grüngaspflicht nach § 43 GModG barwertgewichtet enthalten (2.31 ct); ohne sie 19.20 ct, gegen die Referenz weiterhin teurer. Spannweite 17.35–25.73 ct/kWh; Abschnitt 4.2, Anhang A.10.");

    const cats = ["Gaskessel", "WP heute (DE)", "WP Referenz (EU)"];
    const colors = [C.gas, C.grey, C.wp];
    s.addChart(pres.charts.BAR, [{ name: "Investition", labels: cats, values: [9000, 30000, 16660] }], {
      ...chartBase(), chartColors: colors, barDir: "col",
      x: 0.6, y: 2.0, w: 3.7, h: 3.0, title: "Was der Haushalt sieht: Investition (€)",
      dataLabelFormatCode: FMT_TSD, valAxisMinVal: 0, valAxisMaxVal: 36000,
    });
    s.addChart(pres.charts.BAR, [{ name: "Vollkosten", labels: cats, values: [2151, 2373, 1775] }], {
      ...chartBase(), chartColors: colors, barDir: "col",
      x: 4.5, y: 2.0, w: 3.7, h: 3.0, title: "Was es kostet: Vollkosten (ct/kWh)",
      dataLabelFormatCode: FMT_DEC, valAxisMinVal: 0, valAxisMaxVal: 2800,
    });
    s.addText([
      { text: "Lesart: ", options: { bold: true } },
      { text: "In der Investition gewinnt der Kessel immer. In den Vollkosten verliert er — aber nur gegen die europäische Referenz. Wer heute Gas einbaut, rechnet nicht falsch, sondern gegen die Zahl, die sichtbar und sofort fällig ist." },
    ], { x: 0.6, y: 5.1, w: 7.6, h: 0.75, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });

    const rows = [
      ["LuScale", "Das Ordnungsrecht ist weg.", "GModG 2026: keine 65-%-Regel, kein Betriebsverbot 2045. Gegen den Kessel bleibt nur das ökonomische Argument."],
      ["LuFlame", "Alle rechnen gegen „Gastherme forever“.", "Haushalte wie Planer setzen auf den heutigen Fossilgaspreis — und darauf, dass 2045 verschoben wird. Das ferne Netto-Null-Ziel erzeugt vielleicht Unbehagen, aber keinen Wandel."],
      ["LuGitBranch", "Und im Netz dasselbe Muster.", "Ein Gasvorlauf trägt fast kein Wärmenetz: Er kostet 56 % der Differenz zum Nachfolgesockel und verbraucht dessen Reserve."],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [ic, h, b] = rows[i]; const y = 2.0 + i * 1.28;
      s.addShape(pres.shapes.OVAL, { x: 8.65, y: y + 0.02, w: 0.5, h: 0.5, fill: { color: C.redBg }, line: { color: C.redBg } });
      s.addImage({ data: await icon(ic, C.gas), x: 8.75, y: y + 0.12, w: 0.3, h: 0.3 });
      s.addText(h, { x: 9.35, y, w: 3.4, h: 0.32, fontFace: FB, fontSize: 13, bold: true, color: C.ink, isTextBox: true, margin: 0 });
      s.addText(b, { x: 9.35, y: y + 0.32, w: 3.4, h: 0.9, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    fazit(s, "Was das heißt: ", "Die Reformhebel aus Teil I sind keine Wärmepumpen-Nische. Sie entscheiden, ob sich überhaupt etwas anderes als Gas rechnet — die Wärmepumpe wie die Fernwärme. Nur Wirtschaftlichkeit heute erzeugt Markthochlauf und politisches Momentum.", C.gas);

    s.addNotes("Hintergrund für die Diskussion: Pipeline-Lobby und Fernwärmeseite berichten dasselbe Problem — außer alten Gasheizungen rechnet sich derzeit nichts, weil Haushalte und Planer gegen den aktuellen Fossilgaspreis und die Gastherme rechnen und implizit davon ausgehen, dass 2045 nicht durchgesetzt wird. Den Wandel erzeugt nur die bessere Wirtschaftlichkeit der Erneuerbaren, und die glaubt man nur, wenn die Regulierung nicht wieder kassiert wird. Genau deshalb ist Teil I keine Nebensache: Der Vergleichsmaßstab entscheidet über das Vorzeichen — beim Kessel wie beim Netz. Zahlen: Abschnitt 4.2 und Anhang A.10 (Gasheizung), Abschnitt 7 und Anhang A.7 (Gasvorlauf im Netz). 27 Kombinationen = Investition 6/9/12 T€ × Gaspreis 9/11/13 ct × Nutzungsgrad.");
  }

  // ============ S4 — Systemfrage: Sockel und Budget ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · DIE SYSTEMFRAGE", "Wenn der WP-Preis fällt – was trägt dann noch ein Netz?", 5);
    s.addText("Schritt 1: Wärmegestehungskosten je Quelle – der Sockel vor dem ersten Meter Netz (ct/kWh, Vollkosten)", { x: 0.6, y: 1.8, w: 8.4, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0 });
    const x0 = 3.55, W = 5.2, cmax = 24, ref = 17.75;
    const px = (c) => x0 + W * c / cmax;
    const src = [
      ["Abwärme / Müllverbrennung", 10.02, "+7.72", true, ""],
      ["Gasvorlauf → Abwärme ab 2045*", 15.75, "+2.0", true, ""],
      ["Großwärmepumpe / Geothermie", 16.23, "+1.51", true, ""],
      ["Gasvorlauf → Großwärmepumpe*", 18.55, "−0.8", false, ""],
      ["Gas-KWK, fossil", 20.30, "−2.55", false, ""],
    ];
    src.forEach(([n, v, b, pos], i) => {
      const y = 2.25 + i * 0.5;
      const col = n.startsWith("Gas") ? C.gas : C.fw;
      s.addText(n, { x: 0.6, y, w: 2.9, h: 0.38, fontFace: FB, fontSize: 10.5, color: C.ink, isTextBox: true, margin: 0, valign: "middle", align: "right" });
      // Budget-Pfeilbereich zwischen Sockel und Referenz
      const a = Math.min(px(v), px(ref)), w = Math.abs(px(ref) - px(v));
      s.addShape(pres.shapes.RECTANGLE, { x: a, y: y + 0.12, w, h: 0.14, fill: { color: pos ? C.greenBg : C.redBg }, line: { color: pos ? C.green : C.gas, width: 0.5 } });
      s.addShape(pres.shapes.RECTANGLE, { x: x0, y: y + 0.07, w: px(v) - x0, h: 0.24, fill: { color: col }, line: { color: col } });
      s.addText(v.toFixed(2), { x: px(v) - 0.75, y, w: 0.7, h: 0.38, fontFace: FB, fontSize: 10, bold: true, color: C.white, align: "right", isTextBox: true, margin: 0, valign: "middle" });
      s.addText(b + " ct", { x: 9.05, y, w: 0.9, h: 0.38, fontFace: FB, fontSize: 12, bold: true, color: pos ? C.green : C.gas, isTextBox: true, margin: 0, valign: "middle" });
    });
    // Referenzlinie
    s.addShape(pres.shapes.LINE, { x: px(ref), y: 2.15, w: 0, h: 2.75, line: { color: C.wp, width: 2 } });
    s.addText("Wärmepumpe, Referenz 17.75", { x: px(ref) - 1.0, y: 4.9, w: 2.0, h: 0.25, fontFace: FB, fontSize: 9.5, bold: true, color: C.wp, align: "center", isTextBox: true, margin: 0 });
    s.addText("Budget fürs Netz", { x: 9.05, y: 1.95, w: 1.2, h: 0.3, fontFace: FB, fontSize: 9.5, bold: true, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addText("* Gasvorlauf barwertgewichtet: Gasanlage ~19 Jahre Perspektive (WPG), Netz 50; das Netz wird zu 44 % von der Nachfolgequelle getragen. Sockel hergeleitet aus den Vorlaufkosten des Papiers (5.7 bzw. 2.3 ct über dem Nachfolgesockel). Abschnitt 4.3, 7, A.7.", { x: 0.6, y: 5.2, w: 9.3, h: 0.5, fontFace: FB, fontSize: 8.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });

    // rechts: das Konzept
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.2, y: 1.8, w: 2.5, h: 3.9, rectRadius: 0.08, fill: { color: C.tealBg }, line: { color: C.tealBg } });
    s.addText("Schritt 2: Das Budget", { x: 10.35, y: 1.92, w: 2.25, h: 0.35, fontFace: FB, fontSize: 12, bold: true, color: C.wp, isTextBox: true, margin: 0 });
    s.addText([
      { text: "Budget = WP-Referenz − Sockel", options: { bold: true, breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 5 } },
      { text: "Das ist, was ein Netz je kWh für Trasse, Hausanschluss und Verluste ausgeben darf, bevor die Wärmepumpe billiger ist.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 5 } },
      { text: "Positiv: ", options: { bold: true, color: C.green } }, { text: "das Vorzeichen erlaubt ein Netz – die Dichte entscheidet, ob es reicht.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 5 } },
      { text: "Negativ: ", options: { bold: true, color: C.gas } }, { text: "das Netz verliert vor dem ersten Meter Trasse – bei jeder Dichte." },
    ], { x: 10.35, y: 2.3, w: 2.25, h: 3.3, fontFace: FB, fontSize: 10.5, color: C.ink, isTextBox: true, margin: 0, valign: "top" });

    fazit(s, "Was das heißt: ", "Erst der Sockel, dann die Dichte. Nur wo die Quelle billiger Wärme erzeugt als die Wärmepumpe, bleibt Geld fürs Netz übrig – und das entscheidet das Vorzeichen, nicht die Bebauung.", C.wp);
    s.addNotes("Kern der Methodik, ungewöhnliche Setzung: Statt Dichte als Axiom wird zuerst gerechnet, was die Netzwärme vor dem ersten Meter Trasse kostet (Erzeugung, Speicher, Pumpstrom, Netzverluste, Betrieb – Vollkosten, Annuität), und gegen die preisnormalisierte Wärmepumpe (17.75 ct) gestellt. Die Differenz ist das Budget fürs Verteilnetz (Abschnitt 4.3): Abwärme 10.02 → +7.72; GroßWP/Geothermie 16.23 → +1.51; Gas-KWK fossil 20.30 → −2.55 (negativ: Netz kann bei keiner Dichte gewinnen). Gasvorlauf-Pfade: barwertgewichtet mit der Nachfolgequelle, der Vorlauf kostet stets 56 % der Differenz zwischen Gassockel und Nachfolgesockel – 5.7 ct vor Abwärme, 2.3 ct vor GroßWP; daraus Budget +2.0 bzw. −0.8 (Sockel 15.75 / 18.55 sind aus diesen Papierwerten hergeleitet, nicht als Tabellenzeile im Papier). Nachfolge in grün mit demselben Brennstoff müsste ab 2045 Wärme für 6.4 ct/kWh liefern – weniger als fossiles Gas inkl. CO₂ heute (Befund 2). Nicht auf der Folie: wie gerechnet wird – nächste Folie.");
  }

  // ============ S5 — Wie gerechnet wurde ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · METHODE IN 30 SEKUNDEN", "Wie gerechnet wurde – und was die Wärmeliniendichte ist", 6);
    const meth = [
      ["LuCalculator", "Vollkosten beider Pfade an derselben Systemgrenze", "Annuität nach VDI 2067, 3 % Realzins, Ersatzinvestitionen eingepreist, ohne Förderung"],
      ["LuTarget", "Wärmepumpe auf europäisches Niveau normalisiert", "16'660 € brutto → 17.75 ct/kWh; heutiges DE-Niveau 23.73 ct nur als Vergleich"],
      ["LuRuler", "Trassenkosten aus der AGFW-Praxishilfe", "677 / 1'692 / 3'554 € je Meter (unbefestigt / Bestandsstraße / Innenstadt), Baupreisindex bis Mai 2026"],
      ["LuGauge", "Dichteschwelle berechnet statt gesetzt", "Schwelle = Netzkosten je Trassenmeter und Jahr ÷ Budget; fünf Gebäudetypen, Netzausbau und vorhandenes Netz"],
      ["LuShieldAlert", "Robustheit", "36 Parameterkombinationen, Unsicherheitsmarker an jedem Urteil, Skript offen"],
    ];
    for (let i = 0; i < meth.length; i++) {
      const [ic, h, b] = meth[i]; const y = 1.85 + i * 0.78;
      s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.03, w: 0.44, h: 0.44, fill: { color: C.greenBg }, line: { color: C.greenBg } });
      s.addImage({ data: await icon(ic, C.green), x: 0.7, y: y + 0.13, w: 0.24, h: 0.24 });
      s.addText(h, { x: 1.18, y, w: 5.6, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
      s.addText(b, { x: 1.18, y: y + 0.3, w: 5.6, h: 0.42, fontFace: FB, fontSize: 10, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    // rechts: Wärmeliniendichte
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y: 1.8, w: 5.5, h: 3.95, rectRadius: 0.08, fill: { color: C.blueBg }, line: { color: C.blueBg } });
    s.addText("Wärmeliniendichte", { x: 7.4, y: 1.9, w: 5.1, h: 0.35, fontFace: FB, fontSize: 13, bold: true, color: C.fw, isTextBox: true, margin: 0 });
    s.addText("Wie viel Wärme ein Meter Straße im Jahr verkauft.", { x: 7.4, y: 2.25, w: 5.1, h: 0.4, fontFace: FH, fontSize: 15, italic: true, color: C.ink, isTextBox: true, margin: 0 });
    s.addText("Jahreswärmebedarf aller Anschlüsse ÷ Trassenlänge, in MWh je Meter und Jahr. Kosten hängen an der Länge, Erlöse an der Wärmemenge – dieselbe Trasse ist vor Mehrfamilienhäusern billig je kWh und in der Streusiedlung teuer.", { x: 7.4, y: 2.68, w: 5.1, h: 0.8, fontFace: FB, fontSize: 10.5, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
    const dens = [["EFH-Gebiet, locker", 0.3, 0.8], ["Reihenhaus", 0.8, 1.5], ["Blockrand", 2, 4], ["Innenstadt mit Großverbrauchern", 3, 6]];
    const x0 = 9.55, W = 3.0, dmax = 6.5;
    const px = (d) => x0 + W * d / dmax;
    dens.forEach(([n, a, b], i) => {
      const y = 3.6 + i * 0.36;
      s.addText(n, { x: 7.4, y, w: 2.1, h: 0.3, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "middle" });
      s.addShape(pres.shapes.RECTANGLE, { x: px(a), y: y + 0.07, w: px(b) - px(a), h: 0.16, fill: { color: C.fw }, line: { color: C.fw } });
      s.addText(a + "–" + b, { x: px(b) + 0.06, y, w: 0.7, h: 0.3, fontFace: FB, fontSize: 9, color: C.inkSoft, isTextBox: true, margin: 0, valign: "middle" });
    });
    s.addShape(pres.shapes.LINE, { x: px(1.5), y: 3.55, w: 0, h: 1.5, line: { color: C.setz, width: 1, dashType: "dash" } });
    s.addText("Faustregel Planung: ab ~1.5 (nPro)", { x: px(1.5) - 0.2, y: 5.05, w: 2.5, h: 0.25, fontFace: FB, fontSize: 8.5, color: C.setz, isTextBox: true, margin: 0 });
    s.addText("Typische Bestandsquartiere, MWh/(m·a). Drei Dinge senken sie multiplikativ: lockere Bebauung, geringe Anschlussquote, Sanierung.", { x: 7.4, y: 5.3, w: 5.1, h: 0.4, fontFace: FB, fontSize: 9, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });

    fazit(s, "Die Regel: ", "Ein Netz trägt, wenn die erreichte Wärmeliniendichte über der Schwelle liegt – und die Schwelle ist Netzkosten je Trassenmeter und Jahr ÷ Budget. Kleines Budget, hohe Schwelle.", C.fw);
    s.addNotes("Abschnitt 1, 2, 5.1. Wärmeliniendichte ist die übliche Kenngröße der Netzplanung. Die Formel bewusst groß im Kasten: Sie verbindet die Vorfolie (Budget) mit der nächsten (erreichte Dichte gegen Schwelle). Anschlussquote 100 % in allen Basisrechnungen; bei 60 % steigen sämtliche Schwellen um Faktor 1.67 (Abschnitt 7). Trassenkosten: AGFW-Praxishilfe weist nur Punktwolken aus, Werte sind abgelesene Medianschätzungen für DN 25–100, Preisstand 2021, mit Baupreisindex Ortskanäle ×1.422 auf Mai 2026 hochgerechnet, plus USt. Tiefbau und Oberfläche machen ~60 % der Verlegekosten aus – daher Faktor zwei zwischen unbefestigt und befestigt.");
  }

  // ============ S6 — Wärmequelle entscheidet ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · BEFUND 1 UND 3", "Die Wärmequelle entscheidet, nicht die Dichte", 7);
    const cats = ["EFH-Gebiet", "Reihenhaus", "Kleines MFH", "Gründerzeitblock", "Plattenbau-Zeile"];
    s.addChart(pres.charts.BAR, [
      { name: "erreichte Wärmeliniendichte", labels: cats, values: [86, 214, 391, 744, 423] },
      { name: "Schwelle für ein Abwärmenetz", labels: cats, values: [107, 108, 108, 245, 133] },
    ], {
      x: 0.5, y: 1.8, w: 8.0, h: 3.35, barDir: "col", barGapWidthPct: 45, barGrouping: "clustered",
      chartColors: [C.fw, "D9A441"],
      showLegend: true, legendPos: "t", legendFontFace: FB, legendFontSize: 10, legendColor: C.inkSoft,
      showTitle: true, title: "Netzausbau, MWh je Trassenmeter und Jahr", titleFontFace: FB, titleFontSize: 12, titleColor: C.ink, titleAlign: "left",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: FB, dataLabelFontSize: 9.5, dataLabelColor: C.ink, dataLabelFormatCode: "0\".\"00",
      catAxisLabelFontFace: FB, catAxisLabelFontSize: 10, catAxisLabelColor: C.inkSoft, catAxisLineShow: false, catAxisLabelRotate: 0,
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 850, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      plotArea: { fill: { color: C.paper } }, chartArea: { fill: { color: C.paper } },
    });
    s.addText([
      { text: "Schwelle Großwärmepumpe / Geothermie: ", options: { bold: true, color: C.setz } },
      { text: "5.43 · 5.61 · 5.82 · 20.05 · >50 – kein Quartierstyp erreicht sie beim Netzausbau.", options: { color: C.ink } },
    ], { x: 0.6, y: 5.2, w: 7.9, h: 0.3, fontFace: FB, fontSize: 10.5, isTextBox: true, margin: 0 });
    s.addText("Gründerzeitblock an Innenstadtstraße (3'554 €/m), übrige Bestandsstraße (1'692 €/m) – Setzung. Anschlussquote 100 %. Abschnitt 5.2.", { x: 0.6, y: 5.5, w: 7.9, h: 0.3, fontFace: FB, fontSize: 9, color: C.inkSoft, isTextBox: true, margin: 0 });

    const rows = [
      ["LuFlame", C.fw, "Echte Abwärme trägt fast überall.", "Jeder Quartierstyp außer dem EFH-Gebiet – und dort ist es knapp: bei 13.7 statt 17 Trassenmetern kippt es. In keiner von 36 Kombinationen dreht sich das Ergebnis."],
      ["LuCircleHelp", C.setz, "Großwärmepumpe / Geothermie: unentschieden.", "Beim Ausbau liegen die Pfade wenige Zehntelcent auseinander, das Vorzeichen wechselt in einem Sechstel der Kombinationen. Kein Beleg für fünfzigjährige Infrastruktur."],
      ["LuScale", C.ink, "„Fernwärme rechnet sich bei hoher Dichte“ ist als Pauschale falsch.", "Vor einer Großwärmepumpe verliert der Netzausbau auch im dichtesten Quartier."],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [ic, col, h, b] = rows[i]; const y = 1.85 + i * 1.32;
      s.addShape(pres.shapes.OVAL, { x: 8.85, y: y + 0.02, w: 0.46, h: 0.46, fill: { color: C.paperAlt }, line: { color: C.paperAlt } });
      s.addImage({ data: await icon(ic, col), x: 8.95, y: y + 0.12, w: 0.26, h: 0.26 });
      s.addText(h, { x: 9.45, y, w: 3.3, h: 0.5, fontFace: FB, fontSize: 11.5, bold: true, color: col, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x: 9.45, y: y + 0.5, w: 3.3, h: 0.8, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    fazit(s, "Was das heißt: ", "Die Frage an jedes geplante Netz lautet nicht „Wie dicht ist das Quartier?“, sondern „Welche Quelle speist es – und ist sie vertraglich gesichert?“", C.fw);
    s.addNotes("Befund 1 und 3, Tabelle Abschnitt 5.2. Erreichte Dichte aus der Bebauungsgeometrie hergeleitet (nicht gemessen): EFH 17 m Trasse → 0.86; Reihenhaus 5 m → 2.14; kleines MFH 5 WE, 9 m → 3.91; Gründerzeit 12 WE, 12 m → 7.44; Platte 60 WE, 80 m → 4.23. Abwärmeschwellen 1.07 / 1.08 / 1.08 / 2.45 / 1.33; GroßWP-Schwellen 5.43 / 5.61 / 5.82 / 20.05 / unerreichbar. EFH-Kipppunkt: 13.7 Trassenmeter (Spanne 13–30). Robustheit: 36 Kombinationen (A.3), Abwärme-Ergebnis kippt nie. Befund 3 ausführlich auf Reservefolie R2. Im vorhandenen Netz (Trasse bezahlt) dreht sich die EFH-Zeile: Schwelle 0.63 < 0.86 → Anschluss gewinnt (Befund 6, Abschnitt 5.3) – bei Nachfrage.");
  }

  // ============ S7 — Preisniveau subventioniert Fernwärme ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · BEFUND 4 · POLITISCHE KERNBOTSCHAFT", "Das deutsche WP-Preisniveau subventioniert die Fernwärme", 8);
    const cats = ["gegen WP heute (DE) 23.73 ct", "gegen Gasheizung 21.51 ct", "gegen WP-Referenz 17.75 ct"];
    s.addChart(pres.charts.BAR, [
      { name: "Abwärmenetz", labels: cats, values: [60, 72, 107] },
      { name: "Großwärmepumpen-/Geothermienetz", labels: cats, values: [110, 157, 546] },
    ], {
      x: 0.5, y: 1.8, w: 7.5, h: 3.5, barDir: "col", barGapWidthPct: 50, barGrouping: "clustered",
      chartColors: [C.fw, "D9A441"],
      showLegend: true, legendPos: "t", legendFontFace: FB, legendFontSize: 10, legendColor: C.inkSoft,
      showTitle: true, title: "Dichteschwelle EFH-Referenz, Bestandsstraße – je nach Vergleichsmaßstab (MWh/(m·a))", titleFontFace: FB, titleFontSize: 11.5, titleColor: C.ink, titleAlign: "left",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: FB, dataLabelFontSize: 10, dataLabelColor: C.ink, dataLabelFormatCode: "0\".\"00",
      catAxisLabelFontFace: FB, catAxisLabelFontSize: 10, catAxisLabelColor: C.inkSoft, catAxisLineShow: false, catAxisLabelRotate: 0,
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 620, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      plotArea: { fill: { color: C.paper } }, chartArea: { fill: { color: C.paper } },
    });
    s.addText("So rechnen kommunale Wärmepläne heute: gegen 23.73 ct. Abschnitt 5.1, 5.4.", { x: 0.6, y: 5.35, w: 7.4, h: 0.3, fontFace: FB, fontSize: 9, color: C.inkSoft, isTextBox: true, margin: 0 });
    const rows = [
      [C.fw, "Der Maßstab verschiebt alle Schwellen um Faktor 1.8 bis 5.", "Wärmepläne, die gegen deutsche WP-Vollkosten rechnen, weisen systematisch zu große Netzgebiete aus – ein Planungsrisiko mit fünfzig Jahren Bindung."],
      [C.setz, "Der Hochlauf verschärft das.", "Das Modell rechnet mit heutigen WP-Kosten. Mit dem Markthochlauf sinken sie über die Lernkurve weiter, Tiefbaukosten steigen (Baupreisindex Ortskanäle 2021–2026: +42 %). Nicht im Modell – die Schwellen wandern weiter nach oben."],
      [C.gas, "Und beides bleibt zu teuer.", "Solange Fernwärme gegen einen überhöhten WP-Preis geplant wird, entsteht kein Preisdruck – und der Haushalt entscheidet sich für die dritte Option: fossiles Gas (22 von 27 Kombinationen)."],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [col, h, b] = rows[i]; const y = 1.85 + i * 1.35;
      s.addShape(pres.shapes.RECTANGLE, { x: 8.4, y: y + 0.05, w: 0.07, h: 1.1, fill: { color: col }, line: { color: col } });
      s.addText(h, { x: 8.65, y, w: 4.1, h: 0.45, fontFace: FB, fontSize: 11.5, bold: true, color: col, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x: 8.65, y: y + 0.42, w: 4.1, h: 0.9, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    fazit(s, "Die Kernbotschaft: ", "Das deutsche WP-Preisniveau subventioniert die Fernwärme – und macht beides so teuer, dass viele sich für fossiles Gas entscheiden. Der Preis der Wärmepumpe ist das einzige wirksame Preissignal in einem Sektor, den das Kartellamt für nicht regulierbar hält.", C.gas);
    s.addNotes("Befund 4 (Abschnitt 5.4 Punkt 2) und 9.5. Schwellen EFH-Referenz, Bestandsstraße 1'692 €/m: gegen 23.73 ct Abwärme 0.60, GroßWP 1.10; gegen 21.51 ct 0.72 / 1.57; gegen 17.75 ct 1.07 / 5.46 → Faktor 1.8 (Abwärme) bzw. 5 (GroßWP). Hochlauf-Argument: nicht im Modell gerechnet, aber im Papier benannt (Abschnitt 8, Lernkurve WP vs. Baupreisindex +42 %); als qualitative Verschärfung vortragen, nicht als Zahl. Gas-Verknüpfung: Folie 3 – gegen 23.73 gewinnt der Kessel in 22 von 27 Kombinationen. 9.5: Jeder Euro, den die Reformhebel vom Installationspreis abziehen, wirkt als Obergrenze für den Fernwärmepreis; der Einwand „Gas bleibt billig“ ist ein Einwand gegen den deutschen WP-Preis, nicht gegen die Wärmepumpe.");
  }

  // ============ Heatmap — Befund 4 · Entscheidungsraster (Seite 8) ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · BEFUND 4 · RASTER", "Der Vergleichsmaßstab erzeugt einen ganzen Pfad", 9);

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
    s.addText([
      { text: "Grün: ", options: { bold: true, color: "3B6D11" } }, { text: "das Netz trägt.   " },
      { text: "Blau: ", options: { bold: true, color: "185FA5" } }, { text: "die Wärmepumpe gewinnt.   " },
      { text: "×: ", options: { bold: true, color: "042C53" } }, { text: "negatives Budget – bei keiner Dichte erreichbar.   Je kräftiger der Ton, desto eindeutiger." },
    ], { x: 0.6, y: 4.78, w: 12.1, h: 0.36, fontFace: FB, fontSize: 11, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    const bef = [
      [C.fw, "Links: nur die Abwärme trägt.", "Die Dichte ändert daran fast nichts — die Quelle entscheidet."],
      [C.gas, "Mitte: schon die Gasheizung dreht drei Spalten.", "Der billigste dezentrale Pfad von heute ist ein schwächerer Maßstab als die Wärmepumpe von morgen."],
      [C.setz, "Rechts: alles trägt.", "Gasvorlauf vor Großwärmepumpe geht von „nie“ auf vier von fünf — ein Pfad, den erst der Maßstab erzeugt."],
    ];
    bef.forEach(([col, h, b], i) => {
      const x = 0.6 + i * 4.06;
      s.addText([{ text: h + " ", options: { bold: true, color: col } }, { text: b, options: { color: C.inkSoft } }], { x, y: 5.22, w: 3.86, h: 0.62, fontFace: FB, fontSize: 10, isTextBox: true, margin: 0, valign: "top" });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.95, w: 12.1, h: 0.9, rectRadius: 0.08, fill: { color: C.paperAlt }, line: { color: C.rule, width: 0.75 } });
    s.addText([
      { text: "Die Kernbotschaften:", options: { bold: true, color: C.gas, breakLine: true } },
      { text: "– ", options: { bold: true, color: C.gas } },
      { text: "Das deutsche WP-Preisniveau subventioniert die Fernwärme – und macht beides so teuer, dass viele sich aktuell noch für fossiles Gas entscheiden.", options: { breakLine: true } },
      { text: "– Senken wir WP-Preise, könnten manche Fernwärmenetze zu „sunk costs“ werden oder einen politisch teuren Anschluss- und Benutzungszwang erzwingen." },
    ], { x: 0.8, y: 5.95, w: 11.7, h: 0.9, fontFace: FB, fontSize: 12.5, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
    s.addNotes("Rasterfolie zu Befund 4 — hier wird sichtbar, was die Balken auf der Vorfolie nur als Faktor behaupten. Drei Kacheln, dasselbe Raster, nur der Vergleichsmaßstab wechselt. Links gegen die europäisch normalisierte Referenz von 17.75 ct: Die Abwärmespalte trägt ab dem Reihenhaus, alle drei anderen Spalten verlieren, Gasvorlauf vor Großwärmepumpe ist durchgängig „nie“. Mitte gegen die dezentrale Gasheizung mit 21.51 ct: Anschluss an die Gas-Folie — solange die Wärmepumpe 30'000 € kostet, ist Gas der billigste dezentrale Pfad, und schon gegen diesen schwächeren Maßstab trägt das Netz in drei von vier Spalten. Rechts gegen die heutigen deutschen 23.73 ct: fast vollständig grün, einschließlich des Pfads, der links gar nicht existiert. Kern: Nicht einzelne Schwellen verschieben sich, sondern ein ganzer Technologiepfad entsteht oder verschwindet mit dem Vergleichsmaßstab. Zum EFH-Gebiet: erreichte Dichte aus der Bebauungsgeometrie hergeleitet — 0.86 aus 17 Trassenmetern, plausibel 13 bis 30 m und damit 0.49 bis 1.13; die EFH-Zeile liegt in allen drei Kacheln nah an der Kippkante. Konsequenz trotzdem eindeutig: keine gesicherte Grundlage heißt kein Netz mit fünfzig Jahren Bindung. Alle Zellen hier gegen die Bestandsstraße gerechnet, daher Gründerzeitblock 6.37 statt 3.04 wie auf Folie 6 (dort Innenstadtstraße). Fundstellen: Abschnitt 5.1, Tabelle in 5.2, Anhang A.11.");
  }
  // ============ S9 — Gebäudetyp / Opportunitätskosten ============
  {
    const s = pres.addSlide();
    frame(s, "TEIL II · BEFUND 5 · OPPORTUNITÄTSKOSTEN", "1 Mrd. € = 26'000 oder 413'000 Haushalte, je nach Gebäudetyp", 10);
    const cats = ["WP-Festbetrag 9'000 €", "EFH-Gebiet", "Reihenhausgebiet", "Kleines MFH, 5 WE", "Gründerzeitblock, 12 WE", "Plattenbau-Zeile, 60 WE"];
    s.addChart(pres.charts.BAR, [{ name: "WE", labels: cats, values: [111000, 26000, 54000, 198000, 228000, 413000] }], {
      x: 1.3, y: 1.8, w: 7.2, h: 3.8, barDir: "bar", barGapWidthPct: 40, catAxisOrientation: "maxMin",
      chartColors: [C.wp, C.fw, C.fw, C.fw, C.fw, C.fw],
      showLegend: false, showTitle: true, title: "Wohneinheiten, die 1 Mrd. € umstellt", titleFontFace: FB, titleFontSize: 12, titleColor: C.ink, titleAlign: "left",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFontFace: FB, dataLabelFontSize: 10.5, dataLabelFontBold: true, dataLabelColor: C.ink, dataLabelFormatCode: "#\"'\"##0",
      catAxisLabelFontFace: FB, catAxisLabelFontSize: 10, catAxisLabelColor: C.ink, catAxisLineShow: false,
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 480000, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      plotArea: { fill: { color: C.paper } }, chartArea: { fill: { color: C.paper } },
    });
    // Gruppenklammer "Netzausbau" links der fünf Netz-Balken
    s.addShape(pres.shapes.LEFT_BRACKET, { x: 0.95, y: 2.85, w: 0.15, h: 2.6, line: { color: C.fw, width: 1.5 } });
    s.addText("Netzausbau", { x: 0.05, y: 3.95, w: 1.4, h: 0.4, fontFace: FB, fontSize: 10.5, bold: true, color: C.fw, isTextBox: true, margin: 0, valign: "middle", align: "center", rotate: 270 });
    s.addText("Je Wohneinheit: 9'000 € · 38'770 € · 18'460 € · 5'050 € · 4'390 € · 2'420 €. Trassenmeter aus der Bebauungsgeometrie, Gründerzeitblock in Innenstadtlage. Abschnitt 6, Tabelle 5 des Skripts.", { x: 1.3, y: 5.6, w: 7.2, h: 0.35, fontFace: FB, fontSize: 9, color: C.inkSoft, isTextBox: true, margin: 0 });

    const rows = [
      [C.gas, "Im Ein- und Zweifamilienhausbestand ist der Netzausbau die teuerste Mittelverwendung.", "4.3-mal teurer je Haushalt als der Festbetrag (Spanne 3.6–6.8), im Reihenhaus 2.1-mal. Dort muss sich nicht die Wärmepumpe gegen das Netz rechtfertigen, sondern das Netz gegen die Wärmepumpe."],
      [C.fw, "Im Geschosswohnungsbau ist er die billigste.", "Kleines MFH 0.56, Gründerzeitblock 0.49, Plattenbau 0.27 des Festbetrags – robust über die Trassenmeter-Spanne (Gründerzeit 0.39–0.68)."],
      [C.inkSoft, "Fairness-Hinweis", "Der Festbetrag hebelt privates Kapital; die Rechnung zählt Haushalte in Bewegung je Euro, nicht gebildetes Vermögen."],
    ];
    for (let i = 0; i < rows.length; i++) {
      const [col, h, b] = rows[i]; const y = 1.85 + i * 1.35;
      s.addShape(pres.shapes.RECTANGLE, { x: 8.85, y: y + 0.05, w: 0.07, h: 1.15, fill: { color: col }, line: { color: col } });
      s.addText(h, { x: 9.1, y, w: 3.65, h: 0.5, fontFace: FB, fontSize: 11, bold: true, color: col, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x: 9.1, y: y + 0.5, w: 3.65, h: 0.85, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    fazit(s, "Was das heißt: ", "Eine Wärmeplanung, die eine Stadt mit einem Schwellenwert überzieht, entscheidet in einem Teil des Gebiets systematisch falsch – in beide Richtungen.", C.fw);
    s.addNotes("Abschnitt 6, Tabelle 5 des Skripts. Alle fünf Netzbalken: Szenario „1 Mrd. € geht als Netzausbau in genau diesen Quartierstyp“, nur die Trasse in der Straße, Anschlussquote 100 %. Je WE: Festbetrag 9'000 (britisches Vorbild) → 111'000 WE; EFH 38'770 → 26'000 (4.3×, 3.6–6.8); Reihenhaus 18'460 → 54'000 (2.1×); kleines MFH 5'050 → 198'000 (0.56×); Gründerzeit 4'390 → 228'000 (0.49×, 0.39–0.68); Platte 2'420 → 413'000 (0.27×). Einordnung: Prognos/VKU 2024 beziffert 74 Mrd. € Investitionsbedarf bis 2045 und 3.4–3.5 Mrd. €/a Förderbedarf – Verbandsgutachten, unabhängige Gegenrechnungen fehlen. Befund 6 bei Nachfrage: rund 6 Mio. angeschlossene WE – dort ist die Trasse bezahlt (58 % des Neubauwerts bei 25 a Restlebensdauer), Frage ist Quellenwechsel, nicht Netzbau; mit Abwärme gewinnt der Anschluss in jedem Gebäudetyp.");
  }

  // ============ Politische Kernbotschaften ============
  {
    const s = pres.addSlide();
    frame(s, "POLITISCHE KERNBOTSCHAFTEN", "Was daraus politisch folgt", 11);
    const msgs = [
      [C.gas, "Anschluss- und Benutzungszwang ist ein politisches Risiko.", "Je besser die WP-Reform wirkt, desto öfter zwingt er Bürger ans teurere Netz – oder das Netz wird zum stranded asset."],
      [C.gas, "Der WP-Preis ist das einzige Preissignal im Wärmesektor.", "Das deutsche Preisniveau subventioniert die Fernwärme – und treibt Haushalte zu fossilem Gas."],
      [C.setz, "2045 ist nichts wert, solange nicht gebaut wird.", "Alle wetten auf ein Verschieben des Ziels. Nur Wirtschaftlichkeit heute erzeugt Hochlauf."],
    ];
    for (let i = 0; i < msgs.length; i++) {
      const [col, h, b] = msgs[i]; const y = 2.2 + i * 1.45;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.04, w: 0.09, h: 1.15, fill: { color: col }, line: { color: col } });
      s.addText(h, { x: 0.9, y, w: 11.7, h: 0.5, fontFace: FB, fontSize: 17, bold: true, color: col, isTextBox: true, margin: 0, valign: "top" });
      s.addText(b, { x: 0.9, y: y + 0.5, w: 11.7, h: 0.75, fontFace: FB, fontSize: 12.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    s.addNotes("Anschlusszwang: 9.7 – Zielkonflikt stranded asset vs. Zwang gegen die erkennbare Präferenz der Bürger; beides wird wahrscheinlicher, je besser die Reformhebel aus Teil I wirken. Preissignal: 9.5, Kartellamt hält Fernwärme für faktisch nicht regulierbar. 2045: Gas-Folie und Hintergrundgespräche mit Pipelinelobby und Fernwärmeseite. Weitere, hier gestrichen: falscher Maßstab in Wärmeplänen (Befund 4), Gasvorlauf als Wette auf Politikversagen (6.4 ct/kWh ab 2045, Befund 2), ein Schwellenwert für die ganze Stadt entscheidet in beide Richtungen falsch (Befund 5).");
  }

  // ============ S10 — Instrumente ============
  {
    const s = pres.addSlide();
    frame(s, "SCHLUSSFOLGERUNGEN", "Fünf Instrumente für Bund und Land", 12);
    const inst = [
      ["LuEuro", "BEG-Festbetrag und Normen-Safe-Harbour – das ist zugleich Fernwärme-Preisdisziplin", "Die glaubwürdige dezentrale Alternative ist das einzige wirksame Preissignal. Jeder Euro weniger Installationspreis wirkt als Obergrenze für den Fernwärmepreis. Kapital ist dabei nicht der Engpass: Der Bund bürgt schon für Wärmenetze – niemand ruft es ab.", "Teil I, Hebel 1 + 2 · Abschnitt 9.5"],
      ["LuCircleHelp", "Braucht die Wärmeplanung einen feineren Benchmark?", "Nach Gebäudetyp statt Gebietsmittelwert, gegen die europäisch normalisierte WP-Referenz, mit offengelegten Annahmen und Sanierungs-Stresstest – damit nicht an der falschen Stelle ausgebaut wird.", "WPG § 18 · Abschnitt 9.1"],
      ["LuDatabase", "Ein Wärmenetzregister", "Trassenlängen, Anschlusswerte, Wärmequellen, Netzverluste. Der Branchenverband erfasst heute nur etwa ein Fünftel der geschätzten Trassenlänge.", "Abschnitt 9.2, A.4"],
      ["LuTarget", "Förderung quellenscharf steuern – und Quellenwechsel fördern", "Abwärmenetze haben Vorrang. Gasvorlauf-Netze werden wie ihre Nachfolgequelle behandelt: Förderentscheidung am Nachweis für die Quelle ab 2045. GroßWP/Geothermie nur mit projektspezifischer Vollkostenrechnung. Im Bestand: den Quellenwechsel fördern, nicht neue Trassen.", "BEW · Abschnitt 9.3"],
      ["LuRuler", "Eine öffentliche Kostenbenchmark für den Trassenbau", "Dieses Papier rechnet mit aus AGFW-Punktwolken abgelesenen Medianschätzungen, Preisstand 2021, mit Baupreisindex fortgeschrieben – die Datenlage trägt keine Wirtschaftlichkeitsrechnung. Regelmäßige, regional aufgelöste Erhebung durch Destatis oder BNetzA.", "Abschnitt 9.4, 2.1"],
    ];
    for (let i = 0; i < inst.length; i++) {
      const [ic, h, b, src] = inst[i]; const y = 1.8 + i * 1.0;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 12.1, h: 0.9, rectRadius: 0.06, fill: { color: i === 0 ? C.tealBg : C.white }, line: { color: i === 0 ? C.wp : C.rule, width: 0.75 } });
      s.addText(String(i + 1), { x: 0.75, y: y + 0.1, w: 0.4, h: 0.7, fontFace: FH, fontSize: 22, bold: true, color: C.wp, isTextBox: true, margin: 0, valign: "middle", align: "center" });
      s.addImage({ data: await icon(ic, C.wp), x: 1.25, y: y + 0.3, w: 0.3, h: 0.3 });
      s.addText(h, { x: 1.7, y: y + 0.08, w: 8.6, h: 0.3, fontFace: FB, fontSize: 12, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "middle" });
      s.addText(b, { x: 1.7, y: y + 0.38, w: 8.6, h: 0.5, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
      s.addText(src, { x: 10.4, y: y + 0.08, w: 2.2, h: 0.74, fontFace: FB, fontSize: 9, color: C.setz, align: "right", isTextBox: true, margin: 0, valign: "middle" });
    }
    s.addNotes("Abschnitt 9, umsortiert: Preisdisziplin zuerst, weil sie Teil I und Teil II verbindet. Zu 1 (9.5): vierte Wirkung – je weniger Gaskessel neu eingebaut werden, desto kleiner die Gruppe, der ab 2040 60 % Grüngas zuzumuten wäre, desto eher bleibt die Pflicht durchsetzbar. Zu 2 (9.1): bewusst als Frage – die BAG entscheidet, was sie daraus macht; im Papier steht die Forderung hart (Eignungsgebiete nur gegen normalisierte Referenz, nach Gebäudetyp aufgeschlüsselt). Zu 3 (9.2): AVBFernwärmeV-Novelle zur Kostentransparenz steht ebenfalls im Papier, hier nur das Register. Zu 4 (9.3): rund 6 Mio. bereits angeschlossene Wohnungen – Förderfrage dort ist Quellenwechsel. Zu 5 (9.4): Datenlage aus Abschnitt 2.1 und A.4 – AGFW-Praxishilfe weist nur Punktwolken aus, keine Mittelwerte oder Quantile; Werte sind abgelesene Medianschätzungen DN 25–100, Preisstand 2021, Baupreisindex Ortskanäle ×1.422 auf Mai 2026. Was die BAG damit tun könnte: nicht vorgeben – Reservefolie R1 nur bei direkter Frage.");
  }

  // ============ Linkfolie (Seite 11) ============
  {
    const s = pres.addSlide();
    frame(s, "ZUM NACHLESEN", "Das vollständige Papier", 13);
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
  // ============ R1 — Reserve: Was die BAG damit tun könnte ============
  {
    const s = pres.addSlide();
    frame(s, "RESERVE · NUR BEI NACHFRAGE", "Was die BAG damit tun könnte – Fragen an jede Wärmeplanung", "R1");
    const qs = [
      ["Wurde gegen die Wärmepumpe gerechnet – und gegen welchen Preis?", "23.73 ct oder 17.75 ct entscheiden über das Vorzeichen (Befund 4)."],
      ["Welche Quelle trägt das Netz, und ist sie vertraglich gesichert?", "Abwärme trägt, Großwärmepumpe unentschieden, Gasvorlauf verschiebt die Entscheidung auf 2045 (Befund 1–3)."],
      ["Welche Anschlussquote unterstellt die Rechnung – und was passiert bei Unterschreitung?", "Bei 60 % statt 100 % steigen alle Schwellen um Faktor 1.67."],
      ["Ist Kühlung mitgedacht, und wird ein Leerrohr mitverlegt?", "Wird die Straße ohnehin geöffnet, kostet ein Leerrohr für Kälte- oder Glasfaser heute wenig und morgen ein zweites Aufreißen."],
      ["Wer trägt das Risiko, wenn der WP-Preis fällt und das Netz schon im Boden liegt?", "Bleiben Anschlüsse aus: stranded asset. Kommt der Anschlusszwang: Entscheidung gegen die erkennbare Präferenz der Bürger. Beides wird wahrscheinlicher, je besser die Reformhebel wirken (9.7)."],
    ];
    for (let i = 0; i < qs.length; i++) {
      const [q, a] = qs[i]; const y = 1.85 + i * 0.95;
      s.addText(String(i + 1), { x: 0.6, y, w: 0.4, h: 0.4, fontFace: FH, fontSize: 18, bold: true, color: C.setz, isTextBox: true, margin: 0, valign: "top" });
      s.addText(q, { x: 1.1, y, w: 11.5, h: 0.4, fontFace: FB, fontSize: 12.5, bold: true, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
      s.addText(a, { x: 1.1, y: y + 0.4, w: 11.5, h: 0.5, fontFace: FB, fontSize: 10.5, color: C.inkSoft, isTextBox: true, margin: 0, valign: "top" });
    }
    s.addText("Abschnitt 9.6 und 9.7. Die Antworten entscheiden mehr über die Tragfähigkeit eines Plans als die ausgewiesene Bebauungsdichte.", { x: 0.6, y: 6.6, w: 12.1, h: 0.3, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addNotes("Bewusst als Fragen, nicht als Forderungen an die BAG – sie weiß selbst, was sie tun kann. Nur zeigen, wenn direkt danach gefragt wird.");
  }

  // ============ R2 — Reserve: Großwärmepumpe / Geothermie ============
  {
    const s = pres.addSlide();
    frame(s, "RESERVE · BEFUND 3 · WARUM DAS EIN ERGEBNIS IST", "Großwärmepumpe und Geothermie: unentschieden", "R2");
    const blocks = [
      ["Netzausbau", "Die Vollkosten beider Pfade liegen innerhalb weniger Zehntelcent beieinander; das Vorzeichen wechselt in einem Sechstel der 36 Parameterkombinationen. Budget nur +1.51 ct – die Schwelle liegt bei 5.46 MWh/(m·a) in der Bestandsstraße, über 500 in der Plattenbau-Zeile.", C.fw],
      ["Vorhandenes Netz", "Das Vorzeichen hängt am Gebäudetyp und an einer ungemessenen Größe: der Kostendegression großer Wärmepumpen. Kleines MFH kippt zum Netz (Schwelle 3.40 gegen 3.91 erreicht), Gründerzeitblock nicht (11.72 gegen 7.44).", C.wp],
      ["Ausgerechnet im Geschosswohnungsbau", "Dort, wo sie in der Planungspraxis vorgesehen sind, ist ihr Budget am kleinsten: 0.87 ct im Gründerzeitblock, 0.02 ct in der Plattenbau-Zeile – weil die WP-Vollkosten mit der Gebäudegröße fallen und der Netzsockel kaum.", C.setz],
    ];
    for (let i = 0; i < blocks.length; i++) {
      const [h, b, col] = blocks[i]; const x = 0.6 + i * 4.1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.85, w: 3.9, h: 2.5, rectRadius: 0.08, fill: { color: C.white }, line: { color: C.rule, width: 0.75 } });
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: 2.05, w: 0.6, h: 0.07, fill: { color: col }, line: { color: col } });
      s.addText(h, { x: x + 0.2, y: 2.2, w: 3.5, h: 0.4, fontFace: FB, fontSize: 13, bold: true, color: col, isTextBox: true, margin: 0 });
      s.addText(b, { x: x + 0.2, y: 2.65, w: 3.5, h: 1.6, fontFace: FB, fontSize: 11, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
    }
    fazit(s, "Was das heißt: ", "Für die am häufigsten vorgeschlagene klimaneutrale Netzquelle ist der ökonomische Vorteil nicht nachweisbar. Wer darauf fünfzigjährige Infrastruktur gründet, tut das ohne belastbare Grundlage – Förderung nur mit projektspezifischer Vollkostenrechnung.", C.setz);
    s.addNotes("Kurzfassung Befund 3, Abschnitt 4.4, 5.2, 5.3, 7, 9.3. Das Unentschieden ist ein Ergebnis: Es widerlegt die Planungsvermutung, nicht die Technologie. Die Kostendegression großer WP ist die eine Größe, die niemand gemessen hat – Frage 1 der drei offenen Fragen ans Fachreview.");
  }

  // ============ R3 — Reserve: Bundesbürgschaft für Wärmenetze ============
  {
    const s = pres.addSlide();
    frame(s, "RESERVE · NUR BEI NACHFRAGE", "Der Bund bürgt schon für Wärmenetze – niemand ruft es ab", "R3");
    s.addText("Bundeshaushalt 2026 · Einzelplan 32 · Kap. 3208 · Vermerk 5.19", { x: 0.6, y: 2.0, w: 12.1, h: 0.4, fontFace: FB, fontSize: 14, bold: true, color: C.fw, isTextBox: true, margin: 0 });
    s.addText("KfW-Gewährleistungen des Bundes für Investitionen in Energieinfrastruktur – ausdrücklich Wärmenetze –\ndurch private und öffentliche Investoren. Begründung: über 100 Mrd. € Wärmenetze bis 2045.", { x: 0.6, y: 2.45, w: 12.1, h: 1.0, fontFace: FB, fontSize: 15, color: C.ink, isTextBox: true, margin: 0, valign: "top" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.65, w: 12.1, h: 0.02, fill: { color: C.rule }, line: { color: C.rule } });
    const pts = [
      [C.gas, "Kapital ist nicht der Engpass.", "Eine Bürgschaft senkt den Zins, dreht aber kein negatives Budget."],
      [C.wp, "Sie wirkt, wo das Netz ohnehin trägt:", "Abwärme im Geschosswohnungsbau."],
      [C.setz, "Es fehlt ein Land oder eine Kommune,", "die damit ein Referenzprojekt finanziert."],
    ];
    for (let i = 0; i < pts.length; i++) {
      const [col, h, b] = pts[i]; const y = 3.9 + i * 0.65;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.08, w: 0.08, h: 0.4, fill: { color: col }, line: { color: col } });
      s.addText([{ text: h + " ", options: { bold: true, color: col } }, { text: b, options: { color: C.ink } }], { x: 0.9, y, w: 11.8, h: 0.55, fontFace: FB, fontSize: 15, isTextBox: true, margin: 0, valign: "middle" });
    }
    s.addText("Laut Deckblatt BMF vom 6.11.2025, Hintergrundinformation. Nicht gegen den Haushaltsbeschluss geprüft.", { x: 0.6, y: 6.3, w: 12.1, h: 0.3, fontFace: FB, fontSize: 9.5, color: C.inkSoft, isTextBox: true, margin: 0 });
    s.addNotes("Quelle: Deckblatt zum Einzelplan 32, BMF 6.11.2025, Kap. 3208 Bürgschaften, neuer Vermerk 5.19: Ermächtigung zur Übernahme von Gewährleistungen aus KfW-Förderprodukten im Rahmen des Deutschlandfonds; Begünstigter im Dokument geschwärzt. Begründung: Investitionsbedarf dreistelliger Milliardenbereich bis 2045, Stromverteilnetze 190 Mrd. bis 2035, Wärmenetze über 100 Mrd. bis 2045. Zweck: Kapitalaufnahme erleichtern, auch für Förderungen ohne bestehenden Tatbestand. Verbindung zum Papier: Modell rechnet 3 % real – Bürgschaft verschiebt Kapitalkosten, nicht das Vorzeichen des Budgets. Nur zeigen, wenn nach Finanzierung, KfW, Deutschlandfonds oder Ländern gefragt wird. Mündlicher Anker auf der Schlussfolie bei Instrument 1: „Kapital ist übrigens nicht das Problem – der Bund bürgt schon dafür, und niemand ruft es ab.“");
  }

  await pres.writeFile({ fileName: "/home/claude/bag-deck-v16-lang.pptx" });
  console.log("written");
})();
