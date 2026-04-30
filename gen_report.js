/**
 * OCRA Checklist – Word Report Generator
 * Replicates exactly the style of OCRA___20260430.docx
 * Input: JSON data via stdin or process.argv[2] (file path)
 * Output: report.docx
 */
"use strict";
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, PageBreak, PositionalTabAlignment,
  PositionalTabRelativeTo, PositionalTabLeader
} = require("docx");
const fs = require("fs");

// ─── input ────────────────────────────────────────────────────────────────────
const inPath  = process.argv[2];
const outPath = process.argv[3] || "report.docx";
const raw = inPath ? fs.readFileSync(inPath, "utf8") : fs.readFileSync("/dev/stdin", "utf8");
const D = JSON.parse(raw);

// ─── helpers ──────────────────────────────────────────────────────────────────
const AR = Arial => ({ font: "Arial" });
const arial = (text, opts = {}) => new TextRun({
  text: String(text ?? "–"),
  font: "Arial",
  size: opts.size ?? 15,        // half-points
  bold: opts.bold ?? false,
  color: opts.color ?? "1A1A2E",
  italics: opts.italic ?? false,
});
const bdr = { style: BorderStyle.SINGLE, size: 4, color: "C8D6E5", space: 0 };
const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const tcMar = (v=40) => ({ top: v, bottom: v, left: v, right: v });

function cell(text, opts = {}) {
  const {
    fill = "FFFFFF", bold = false, color = "1A1A2E", size = 15,
    align = AlignmentType.LEFT, italic = false, w, colspan
  } = opts;
  const children_p = [];
  if (Array.isArray(text)) {
    // multiple TextRun passed directly
    children_p.push(new Paragraph({
      children: text,
      alignment: align,
      spacing: { before: 0, after: 0 },
    }));
  } else {
    children_p.push(new Paragraph({
      children: [arial(text, { bold, color, size, italic })],
      alignment: align,
      spacing: { before: 0, after: 0 },
    }));
  }
  const tcProps = {
    borders,
    shading: { fill, type: ShadingType.CLEAR },
    margins: tcMar(opts.margin ?? 40),
    verticalAlign: VerticalAlign.CENTER,
  };
  if (w !== undefined) tcProps.width = { size: w, type: WidthType.DXA };
  if (colspan) tcProps.columnSpan = colspan;
  return new TableCell({ ...tcProps, children: children_p });
}

function row(cells, height = 180) {
  return new TableRow({
    children: cells,
    height: { value: height, rule: "atLeast" },
  });
}

function h1(text) {
  return new Paragraph({
    children: [arial(text, { bold: true, color: "1E3A5F", size: 20 })],
    spacing: { before: 100, after: 40 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E5F8A", space: 1 } },
  });
}
function h2(text, opts = {}) {
  return new Paragraph({
    children: [arial(text, { bold: true, color: opts.color ?? "2E5F8A", size: 17 })],
    spacing: { before: 80, after: 20 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "C8D6E5", space: 1 } },
  });
}
function spacer(after = 40, before = 0) {
  return new Paragraph({ spacing: { before, after } });
}

// color map for livello → hex
function lvc(livello) {
  if (!livello) return "888888";
  if (livello.includes("OTTIMALE") || livello.includes("ACCETTABILE")) return "27AE60";
  if (livello.includes("BORDERLINE") || livello.includes("MOLTO LIEVE")) return "E6A817";
  if (livello.includes("MEDIO")) return "D4681A";
  if (livello.includes("ELEVATO") && !livello.includes("MOLTO")) return "C0392B";
  if (livello.includes("MOLTO ELEVATO")) return "7B241C";
  return "888888";
}

// ─── document width ───────────────────────────────────────────────────────────
// A4, margins 1cm each side → content ≈ 9638 DXA  (from sample XML)
const TW = 9638;
const BORDER_C = "C8D6E5";

// ─── data shortcuts ───────────────────────────────────────────────────────────
const an  = D.anagrafica || {};
const res = D.results    || {};
const disponibili = ["DX","SX"].filter(s => res[s] && !res[s].errors);
const r0  = disponibili.length ? res[disponibili[0]] : null;

function fmtVal(v) { return v !== undefined && v !== null ? String(v) : "–"; }

// ─────────────────────────────────────────────────────────────────────────────
//  SECTIONS
// ─────────────────────────────────────────────────────────────────────────────

// 1. BANNER HEADER TABLE
const bannerTable = new Table({
  width: { size: TW, type: WidthType.DXA },
  columnWidths: [TW],
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
             left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
             insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE } },
  rows: [row([new TableCell({
    width: { size: TW, type: WidthType.DXA },
    shading: { fill: "1E3A5F", type: ShadingType.CLEAR },
    margins: tcMar(40),
    children: [
      new Paragraph({
        children: [arial("OCRA CHECKLIST  –  VALUTAZIONE RISCHIO MOVIMENTI RIPETITIVI ARTI SUPERIORI",
          { bold: true, color: "FFFFFF", size: 22 })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 20 },
      }),
      new Paragraph({
        children: [arial("ISO 11228-3  ·  D.Lgs. 81/08  ·  Studio Essepi S.r.l.",
          { color: "8BB4D4", size: 15 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
      }),
    ],
  })], 0)],
});

// 2. DATI GENERALI
const arti_str = disponibili.length
  ? disponibili.map(s => `Arto ${s}`).join(", ")
  : "–";
const anagRows = [
  ["Campo", "Valore"],
  ["Azienda", an.azienda || "–"],
  ["Reparto", an.reparto || "–"],
  ["Mansione / Postazione", an.mansione || "–"],
  ["Valutatore", an.valutatore || "–"],
  ["Data valutazione", an.data || "–"],
  ["N° revisione", an.nr_rev || "01"],
  ["Arti esaminati", arti_str],
];
const anagTable = new Table({
  width: { size: TW, type: WidthType.DXA },
  columnWidths: [2835, 6803],
  rows: anagRows.map((r, i) => {
    const isH = i === 0;
    return new TableRow({
      children: [
        cell(r[0], { w: 2835, fill: isH ? "2E5F8A" : "EEF4FB",
          bold: isH, color: isH ? "FFFFFF" : "1A1A2E", size: 15 }),
        cell(r[1], { w: 6803, fill: isH ? "2E5F8A" : "FFFFFF",
          bold: isH, color: isH ? "FFFFFF" : "1A1A2E", size: 15 }),
      ],
      height: { value: 180, rule: "atLeast" },
    });
  }),
});

// 3. TABELLA LIVELLI RISCHIO
const livRows = [
  ["Punteggio", "Livello di Rischio", "Azione raccomandata"],
  ["≤ 7.5", "OTTIMALE / ACCETTABILE", "Nessun intervento necessario"],
  ["7.6–11.0", "BORDERLINE / MOLTO LIEVE", "Consigliati miglioramenti ergonomici"],
  ["11.1–14.0", "MEDIO", "Interventi di miglioramento necessari"],
  ["14.1–22.5", "ELEVATO", "Interventi urgenti necessari"],
  ["> 22.5", "MOLTO ELEVATO", "Intervento immediato obbligatorio"],
];
const livColors = ["","27AE60","E6A817","D4681A","C0392B","7B241C"];
const livTable = new Table({
  width: { size: TW, type: WidthType.DXA },
  columnWidths: [1700, 3500, 4438],
  rows: livRows.map((r, i) => {
    const isH = i === 0;
    const cc = livColors[i] || "FFFFFF";
    const textColor = (i > 0) ? "FFFFFF" : "FFFFFF";
    const cellFill = isH ? "2E5F8A" : cc;
    return new TableRow({
      children: r.map((txt, ci) => cell(txt, {
        fill: cellFill,
        bold: true,
        color: "FFFFFF",
        size: 15,
        align: ci === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
        w: [1700, 3500, 4438][ci],
      })),
      height: { value: 180, rule: "atLeast" },
    });
  }),
});

// ─── FATTORE TABLE HELPERS ────────────────────────────────────────────────────

function simpleTable(colWidths, rows_data) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows_data.map((rowArr, ri) => {
      const isH = ri === 0;
      return new TableRow({
        children: rowArr.map((txt, ci) => cell(txt, {
          w: colWidths[ci],
          fill: isH ? "2E5F8A" : (ri % 2 === 0 ? "FFFFFF" : "EBF3FB"),
          bold: isH,
          color: isH ? "FFFFFF" : "1A1A2E",
          size: 15,
          align: ci === 0 && !isH ? AlignmentType.LEFT : AlignmentType.LEFT,
        })),
        height: { value: 180, rule: "atLeast" },
      });
    }),
  });
}

// ─────────────────────────────────────────────────────────────────────────────
//  DETTAGLIO FATTORI
// ─────────────────────────────────────────────────────────────────────────────
const detail = [];

// Fattore A
detail.push(h2("Fattore A  –  Tempo di Recupero  (comune a entrambi gli arti)"));
if (r0) {
  detail.push(simpleTable([7638, 2000],
    [["Condizione rilevata", "Punteggio"],
     [r0.desc_recupero || "–", String(r0.pA ?? "–")]]));
} else {
  detail.push(new Paragraph({ children: [arial("Dati non disponibili.", { italic: true, color: "888888" })], spacing: { before: 0, after: 40 } }));
}
detail.push(spacer(40));

// Fattore B
detail.push(h2("Fattore B  –  Frequenza di Azione"));
{
  const bRows = [["Arto", "Condizione rilevata", "Punteggio"]];
  disponibili.forEach(s => {
    const r = res[s];
    bRows.push([`Arto ${s}`, r.desc_frequenza || "–", String(r.pB ?? "–")]);
  });
  detail.push(simpleTable([1400, 6838, 1400], bRows));
}
detail.push(spacer(40));

// Fattore C
detail.push(h2("Fattore C  –  Forza"));
{
  const cRows = [["Arto", "Condizione rilevata", "Punteggio"]];
  disponibili.forEach(s => {
    const r = res[s];
    cRows.push([`Arto ${s}`, r.desc_forza || "–", String(r.pC ?? "–")]);
  });
  detail.push(simpleTable([1400, 6838, 1400], cRows));
}
detail.push(spacer(40));

// Fattore D
detail.push(h2("Fattore D  –  Postura e Movimenti Incongrui"));
{
  const dColW = [1400, 1800, 5238, 1200];
  const dRows = [["Arto", "Sottofattore", "Condizione rilevata", "Punt."]];
  disponibili.forEach(s => {
    const r = res[s];
    const sc_st_s = String(r.sc_st !== undefined ? r.sc_st : 0);
    dRows.push([`Arto ${s}`, "Spalla",      r.desc_spalla || "–", String(r.sc_sp ?? 0)]);
    dRows.push(["",          "Gomito",      r.desc_gomito || "–", String(r.sc_go ?? 0)]);
    dRows.push(["",          "Polso",       r.desc_polso  || "–", String(r.sc_po ?? 0)]);
    dRows.push(["",          "Mano/Dita",   r.desc_mano   || "–", String(r.sc_ma ?? 0)]);
    dRows.push(["",          "Stereotipia", r.desc_ster   || "–", sc_st_s]);
    dRows.push(["", `MAX=${r.pD ?? 0} + Ster=${r.sc_st ?? 0}`, "➡ TOTALE D", String(r.pD_tot ?? 0)]);
  });
  detail.push(simpleTable(dColW, dRows));
}
detail.push(spacer(40));

// Fattore E
detail.push(h2("Fattore E  –  Fattori Complementari"));
disponibili.forEach(s => {
  const r = res[s];
  detail.push(new Paragraph({
    children: [arial(`   Arto ${s}  [punteggio: ${r.pE ?? 0} / max 12]`,
      { bold: true, color: "555555", size: 15 })],
    spacing: { before: 20, after: 8 },
  }));
  if (r.complementari_attivi && r.complementari_attivi.length > 0) {
    const eRows = [["Fattore complementare", "Punti"]];
    r.complementari_attivi.forEach(([desc, sc]) => eRows.push([desc, String(sc)]));
    detail.push(simpleTable([7638, 2000], eRows));
  } else {
    detail.push(new Paragraph({
      children: [arial("Nessun fattore complementare rilevato.", { italic: true, color: "888888", size: 14 })],
      spacing: { before: 0, after: 10 },
    }));
  }
});
detail.push(spacer(40));

// MD
detail.push(h2("Moltiplicatore Durata (MD)  (comune a entrambi gli arti)"));
if (r0) {
  detail.push(simpleTable([7638, 2000],
    [["Durata compito ripetitivo", "Moltiplicatore"],
     [r0.desc_durata || "–", String(r0.md ?? "–")]]));
}
detail.push(spacer(40));

// ─────────────────────────────────────────────────────────────────────────────
//  CALCOLO RIEPILOGATIVO
// ─────────────────────────────────────────────────────────────────────────────
const calcoloParas = [];
disponibili.forEach(s => {
  const r = res[s];
  const col = lvc(r.livello);
  calcoloParas.push(new Paragraph({
    children: [
      arial(`Arto ${s}:  `, { bold: true, size: 17 }),
      arial(`Grezzo = ${r.pA} + ${r.pB} + ${r.pC} + ${r.pD_tot} + ${r.pE} = ${r.grezzo}  →  OCRA = ${r.grezzo} × ${r.md} = `,
        { size: 17 }),
      arial(String(r.finale), { bold: true, size: 22, color: col }),
      arial(`  →  ${r.livello}`, { bold: true, size: 17, color: col }),
    ],
    spacing: { before: 20, after: 20 },
  }));
});

// ─────────────────────────────────────────────────────────────────────────────
//  RISULTATO SINTETICO
// ─────────────────────────────────────────────────────────────────────────────
const nArti = disponibili.length;
const colW_sint = nArti === 2
  ? [2835, 3402, 3401]
  : nArti === 1 ? [2835, 6803] : [9638];
const hdrFills_sint = ["2E5F8A", "1A5276", "1A6B3C"];

const sintRows = [];
// header
sintRows.push(new TableRow({
  children: (["Fattore"].concat(disponibili.map(s => `Arto ${s}`))).map((txt, ci) =>
    cell(txt, { w: colW_sint[ci], fill: hdrFills_sint[ci], bold: true, color: "FFFFFF",
                size: 16, align: AlignmentType.CENTER, margin: 50 })),
  height: { value: 200, rule: "atLeast" },
}));
// data rows
const sintFactors = [
  ["pA", "A – Recupero"], ["pB", "B – Frequenza"], ["pC", "C – Forza"],
  ["pD_tot", "D – Postura"], ["pE", "E – Complementari"], ["md", "MD – Durata"],
];
sintFactors.forEach(([fkey, flabel], ri) => {
  sintRows.push(new TableRow({
    children: ([flabel].concat(disponibili.map(s => fmtVal(res[s][fkey])))).map((txt, ci) =>
      cell(txt, { w: colW_sint[ci], fill: ri % 2 === 0 ? "FFFFFF" : "EEF4FB",
                  color: "1A1A2E", size: 15,
                  align: ci === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })),
    height: { value: 180, rule: "atLeast" },
  }));
});
// PUNTEGGIO FINALE row
sintRows.push(new TableRow({
  children: (["PUNTEGGIO FINALE"].concat(disponibili.map(s => fmtVal(res[s].finale)))).map((txt, ci) =>
    cell(txt, {
      w: colW_sint[ci],
      fill: ci === 0 ? "E8F2FB" : "1E3A5F",
      bold: true,
      color: ci === 0 ? "1A1A2E" : "FFFFFF",
      size: ci === 0 ? 15 : 32,
      align: AlignmentType.CENTER,
      margin: 60,
    })),
  height: { value: 420, rule: "atLeast" },
}));
// LIVELLO row
sintRows.push(new TableRow({
  children: (["LIVELLO DI RISCHIO"].concat(disponibili.map(s => fmtVal(res[s].livello)))).map((txt, ci) =>
    cell(txt, {
      w: colW_sint[ci],
      fill: ci === 0 ? "E8F2FB" : lvc(ci > 0 ? res[disponibili[ci-1]].livello : ""),
      bold: ci > 0,
      color: ci === 0 ? "1A1A2E" : "FFFFFF",
      size: ci === 0 ? 15 : 16,
      align: AlignmentType.CENTER,
      margin: 50,
    })),
  height: { value: 380, rule: "atLeast" },
}));

const sintTable = new Table({
  width: { size: TW, type: WidthType.DXA },
  columnWidths: colW_sint,
  rows: sintRows,
});

// note finali
const noteParas = disponibili.map(s => {
  const r = res[s];
  return new Paragraph({
    children: [
      arial(`Arto ${s}: `, { bold: true, size: 15 }),
      arial(r.messaggio || "", { italic: true, size: 15 }),
    ],
    spacing: { before: 10, after: 8 },
  });
});

// ─────────────────────────────────────────────────────────────────────────────
//  FOOTER
// ─────────────────────────────────────────────────────────────────────────────
const az_footer = (an.azienda || "–").trim() || "–";
const dat_yyyymmdd = (an.data || "").replace(/\//g, "").split("").reverse().join("").replace(/(\d{4})(\d{2})(\d{2})/,"$3$2$1") ||
  new Date().toISOString().slice(0,10).replace(/-/g,"");
// simple: "Valutazione rischio movimenti ripetitivi OCRA  –  {az}  –  {date}"
const footerPara = new Paragraph({
  children: [arial(
    `Valutazione rischio movimenti ripetitivi OCRA  –  ${az_footer}  –  ${dat_yyyymmdd}`,
    { size: 14, color: "888888" })],
});

// ─────────────────────────────────────────────────────────────────────────────
//  ASSEMBLE DOCUMENT
// ─────────────────────────────────────────────────────────────────────────────
const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 },
      },
    },
    footers: {
      default: new Footer({ children: [footerPara] }),
    },
    children: [
      bannerTable,
      spacer(40),

      h1("1.  DATI GENERALI"),
      anagTable,
      spacer(40),

      h1("2.  TABELLA DI RIFERIMENTO – LIVELLI DI RISCHIO OCRA CHECKLIST"),
      livTable,
      spacer(40),

      h1("3.  DETTAGLIO DEI FATTORI"),
      ...detail,

      h1("4.  CALCOLO RIEPILOGATIVO"),
      ...calcoloParas,
      spacer(40),

      h1("5.  RISULTATO SINTETICO"),
      sintTable,
      spacer(20),
      ...noteParas,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log("OK:" + outPath);
}).catch(err => {
  console.error("ERR:" + err.message);
  process.exit(1);
});
