// Bilancio di genere 2025 – ODCEC Forlì – VERSIONE 2: schema delle Linee guida CN, grafica "report a colori".
// Legge dati_2025.json e i grafici in grafici/. Uso: python3 grafici.py && node genera_bilancio_v2.js
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, Footer, Header, PageNumber, ImageRun, TableOfContents, BorderStyle, VerticalAlign,
} = require('docx');

const D = JSON.parse(fs.readFileSync(path.join(__dirname, 'dati_2025.json'), 'utf8'));
const FONT = 'Verdana';
const INK = '1F2A37', GREY = '4B5563', WHITE = 'FFFFFF';
const BLU = '1F4E79', VIOLA = '4A3AA7', AZZ = '2A78D6', ARANCIO = 'C9541F';
const FILL_BLU = 'E8EFF8', FILL_VIOLA = 'EFECFA', FILL_SABBIA = 'F6F3EC', FILL_ARANCIO = 'FCEEE6';
const W = 9306; // larghezza utile (A4, margini 1300)

const run = (t, o = {}) => new TextRun({ text: t, font: FONT, color: INK, size: 21, ...o });
const daFare = (t) => new TextRun({ text: `[DA INTEGRARE: ${t}]`, font: FONT, color: INK, size: 21, shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'FFE873' } });
const kids = (c) => (Array.isArray(c) ? c : [c]).map(x => typeof x === 'string' ? run(x) : x);
const p = (c, o = {}) => new Paragraph({ children: kids(c), spacing: { after: 140, line: 340 }, ...o });
const b = (c) => new Paragraph({ numbering: { reference: 'punti', level: 0 }, spacing: { after: 80, line: 320 }, children: kids(c) });
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NOB = { top: NONE, bottom: NONE, left: NONE, right: NONE };

// titolo di capitolo: fascia colorata a tutta larghezza
const h1 = (num, t, fill = BLU) => new Paragraph({
  heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { before: 0, after: 260 },
  shading: { type: ShadingType.CLEAR, color: 'auto', fill },
  border: { top: { style: BorderStyle.SINGLE, size: 24, color: fill, space: 8 }, bottom: { style: BorderStyle.SINGLE, size: 24, color: fill, space: 8 } },
  indent: { left: 0, right: 0 },
  children: [new TextRun({ text: num ? `${num}  ` : '', font: FONT, bold: true, size: 34, color: 'C9D8EC' }), new TextRun({ text: t, font: FONT, bold: true, size: 34, color: WHITE })],
});
const h2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 320, after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: AZZ, space: 4 } },
  children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: BLU })],
});
const fonte = (t) => p(run('Fonte: ' + t, { size: 16, color: GREY, italics: true }), { spacing: { after: 160 } });
const it = (n, dec = 0) => n.toLocaleString('it-IT', { minimumFractionDigits: dec, maximumFractionDigits: dec });
const pct = (a, tot) => it(a / tot * 100, 1) + '%';
const img = (file, w = 600, h = 315) => {
  const f = path.join(__dirname, 'grafici', file);
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 80 }, children: [new ImageRun({ type: 'png', data: fs.readFileSync(f), transformation: { width: w, height: h }, altText: { title: file, description: 'Grafico: i valori sono riportati nella tabella vicina', name: file } })] });
};
// dimensioni proporzionali dei PNG
const pngSize = (file, w) => { const buf = fs.readFileSync(path.join(__dirname, 'grafici', file)); const pw = buf.readUInt32BE(16), ph = buf.readUInt32BE(20); return [w, Math.round(w * ph / pw)]; };
const graf = (file, w = 600) => img(file, ...pngSize(file, w));

function table(head, rows, widths, opt = {}) {
  const total = widths.reduce((a, c) => a + c, 0);
  const cell = (t, w, hd, zebra, bold) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: hd ? { type: ShadingType.CLEAR, color: 'auto', fill: opt.headFill || BLU } : zebra ? { type: ShadingType.CLEAR, color: 'auto', fill: FILL_SABBIA } : undefined,
    borders: { top: NONE, left: NONE, right: NONE, bottom: { style: BorderStyle.SINGLE, size: 4, color: 'D9DEE6' } },
    margins: { top: 80, bottom: 80, left: 110, right: 110 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [t instanceof TextRun ? t : run(String(t), { size: 18, bold: hd || bold, color: hd ? WHITE : INK })] })],
  });
  const tr = [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], true)) })];
  rows.forEach((r, k) => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false, k % 2 === 1, opt.boldLast && k === rows.length - 1)) })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: tr });
}

// riquadro "in sintesi" con barra laterale colorata
function box(titolo, testi, fill = FILL_VIOLA, bar = VIOLA) {
  const ch = [new Paragraph({ spacing: { after: 100 }, children: [run(titolo.toUpperCase(), { bold: true, size: 17, color: bar })] })];
  [].concat(testi).forEach(t => ch.push(new Paragraph({ spacing: { after: 80, line: 320 }, children: kids(t) })));
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: 'auto', fill },
      borders: { top: NONE, bottom: NONE, right: NONE, left: { style: BorderStyle.SINGLE, size: 36, color: bar } },
      margins: { top: 160, bottom: 140, left: 240, right: 200 }, children: ch,
    })] })],
  });
}
const spazio = () => new Paragraph({ spacing: { after: 120 }, children: [] });

// riquadri con i numeri chiave (2 o 3 per riga)
function tiles(items, perRiga = 3) {
  const cw = Math.floor(W / perRiga);
  const fills = [FILL_BLU, FILL_VIOLA, FILL_ARANCIO];
  const colori = [BLU, VIOLA, ARANCIO];
  const rows = [];
  for (let i = 0; i < items.length; i += perRiga) {
    rows.push(new TableRow({ cantSplit: true, children: items.slice(i, i + perRiga).map(([num, testo], j) => new TableCell({
      width: { size: cw, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: 'auto', fill: fills[(i + j) % 3] },
      borders: { top: { style: BorderStyle.SINGLE, size: 24, color: WHITE }, bottom: { style: BorderStyle.SINGLE, size: 24, color: WHITE }, left: { style: BorderStyle.SINGLE, size: 24, color: WHITE }, right: { style: BorderStyle.SINGLE, size: 24, color: WHITE } },
      margins: { top: 200, bottom: 200, left: 180, right: 180 },
      children: [
        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: num, font: FONT, bold: true, size: 44, color: colori[(i + j) % 3] })] }),
        new Paragraph({ spacing: { line: 280 }, children: [run(testo, { size: 17 })] }),
      ],
    })) }));
  }
  return new Table({ width: { size: cw * perRiga, type: WidthType.DXA }, columnWidths: Array(perRiga).fill(cw), rows });
}

const C = [];
const I = D.iscritti, CA = D.cassa_cnpadc;
const [f25, m25] = I.per_genere['2025'], tot25 = f25 + m25;
const pr = D.praticanti;

// ================= COPERTINA
C.push(new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: [W],
  rows: [new TableRow({ height: { value: 6200, rule: 'atLeast' }, children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: 'auto', fill: BLU }, borders: NOB,
    margins: { top: 600, bottom: 500, left: 500, right: 500 }, verticalAlign: VerticalAlign.BOTTOM,
    children: [
      new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: 'ORDINE DEI DOTTORI COMMERCIALISTI E DEGLI ESPERTI CONTABILI DI FORLÌ', font: FONT, size: 18, bold: true, color: 'C9D8EC' })] }),
      new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 120 }, children: [new TextRun({ text: 'Bilancio di genere', font: FONT, bold: true, size: 64, color: WHITE })] }),
      new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: '2025', font: FONT, bold: true, size: 96, color: 'F2B38F' })] }),
      new Paragraph({ children: [new TextRun({ text: 'A cura del Comitato Pari Opportunità', font: FONT, size: 24, color: WHITE })] }),
    ],
  })] })],
}));
C.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [run('Assemblea degli iscritti – Cesena, 26 novembre 2026', { size: 20, color: GREY })] }));
C.push(p([run('BOZZA DI LAVORO del 17 novembre 2026 – versione 2. ', { bold: true, size: 18 }), run('Struttura secondo le Linee guida del CPO nazionale (14 marzo 2024).', { size: 18, color: GREY })]));
C.push(spazio());
C.push(tiles([
  [pct(f25, tot25), `donne tra gli iscritti (${f25} su ${tot25}). In Italia sono il 33,9%.`],
  ['61%', `donne tra i praticanti (${pr['2025'][0]} su ${pr['2025'][0] + pr['2025'][1]}).`],
  ['54,3%', 'divario di reddito donne-uomini (Cassa Dottori, redditi 2024). Emilia-Romagna 48,2%, Italia 46,8%.'],
  ['5 su 11', 'donne nel Consiglio 2026-2030, guidato da una Presidente. Tre cariche su quattro a donne.'],
  ['27,7%', 'relatrici negli eventi formativi del 2025: poco più di una su quattro.'],
  ['1.500 €', 'primo budget del CPO, nel bilancio di previsione 2026.'],
]));
C.push(p(run('Iscritti e praticanti al 31/12/2025; divario sui redditi 2024 comunicati alla Cassa nel 2025; Consiglio in carica dal 23/2/2026.', { size: 16, color: GREY }), { spacing: { before: 100 } }));

// ================= SOMMARIO
C.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 200 }, children: [new TextRun({ text: 'Sommario', font: FONT, bold: true, size: 32, color: BLU })] }));
C.push(new TableOfContents('Sommario', { hyperlink: true, headingStyleRange: '1-2' }));
C.push(p(run('In Word il sommario si aggiorna con tasto destro → Aggiorna campo.', { size: 16, color: GREY })));
C.push(spazio());
C.push(box('Come leggere questo documento', [
  'Abbiamo scritto il bilancio per essere letto da tutte e tutti, anche da chi ha un disturbo specifico dell’apprendimento (DSA): carattere senza grazie, frasi brevi, ogni grafico accompagnato dalla tabella con i numeri.',
  'Usiamo il femminile e il maschile; quando per brevità c’è solo il maschile, si intende riferito a tutte le persone. Il documento è pubblicato solo in formato digitale sul sito dell’Ordine.',
], FILL_BLU, AZZ));

// ================= 1 PRESENTAZIONE
C.push(h1('1', 'Presentazione'));
C.push(p('Il bilancio di genere fotografa la presenza di donne e uomini nella professione e nell’Ordine di Forlì e misura le differenze di reddito, di carriera e di rappresentanza. Serve a decidere dove intervenire e, l’anno dopo, a verificare se gli interventi hanno funzionato.'));
C.push(p('È il terzo bilancio di genere dell’Ordine e il primo del Comitato eletto per il mandato 2026-2030. Rispetto agli anni scorsi guarda anche fuori dall’Albo: gli incarichi dell’OCC, le cariche nelle imprese della provincia, la formazione, le spese.'));
if (D.confronto_italia) {
  C.push(h2('Forlì e l’Italia in una pagina'));
  C.push(table(['Indicatore', 'Forlì', 'Italia'], D.confronto_italia, [3000, 2900, 3406]));
  fonte('Forlì: elaborazioni CPO su dati della Segreteria e della Cassa Dottori. Italia: Bilancio di genere 2024 del Comitato Nazionale Pari Opportunità (25 novembre 2025) e Cassa Dottori.');
  C.push(box('In sintesi', D.confronto_testo));
}

// ================= 2 QUADRO NORMATIVO
C.push(h1('2', 'Il quadro di riferimento'));
C.push(h2('Internazionale ed europeo'));
C.push(b('Conferenza ONU di Pechino (1995): integrare la prospettiva di genere nelle politiche e nei bilanci (gender budgeting).'));
C.push(b('Strategia europea per la parità di genere 2020-2025 e risoluzioni del Parlamento europeo sul gender budgeting (2003, 2013, 2019).'));
C.push(b([run('Direttiva (UE) 2023/970 del 10 maggio 2023 sulla trasparenza retributiva, recepita in Italia con il d.lgs. 7 maggio 2026, n. 96, in vigore dal 7 giugno 2026. Tra l’altro vieta di chiedere ai candidati la retribuzione precedente e obbliga i datori di lavoro con almeno 100 dipendenti a comunicare il divario retributivo di genere; se supera il 5% senza giustificazione oggettiva, va avviata una valutazione congiunta con le rappresentanze dei lavoratori.')]));
C.push(h2('Nazionale e della professione'));
C.push(b('Costituzione, artt. 3, 37, 51 e 117: uguaglianza, parità nel lavoro e nell’accesso alle cariche.'));
C.push(b('Codice delle pari opportunità (d.lgs. 198/2006) e legge 162/2021, che ha introdotto la certificazione della parità di genere (UNI/PdR 125:2022).'));
C.push(b('D.lgs. 139/2005 sull’ordinamento della professione; Regolamento per la costituzione e l’elezione dei CPO (Consiglio Nazionale, 22 giugno 2023); Linee guida al bilancio di genere (CPO nazionale, 14 marzo 2024).'));
C.push(box('Perché ci riguarda', 'La direttiva sulla trasparenza retributiva conta due volte per la nostra professione: negli studi come datori di lavoro e verso le imprese clienti, che dovranno misurare e comunicare il proprio divario retributivo.', FILL_ARANCIO, ARANCIO));

// ================= 3 CPO
C.push(h1('3', 'Il Comitato Pari Opportunità'));
C.push(h2('Composizione'));
C.push(p(`Il Comitato in carica è stato nominato il ${D.cpo.nomina} per il quadriennio 2026-2030 ed è composto da 7 persone.`));
C.push(table(['Componente', 'Ruolo'], D.cpo.nomi, [5000, 4306]));
C.push(p('Il Comitato precedente (2022-2026) era presieduto da Monica Orecchioni ed era composto da 5 donne e 2 uomini. Ha curato i bilanci di genere 2023 e 2024.'));
C.push(h2('Funzioni'));
C.push(p('Il CPO promuove la parità tra iscritte, iscritti e tirocinanti: analizza i dati, propone iniziative al Consiglio, organizza eventi formativi, vigila contro le discriminazioni e redige il bilancio di genere (Regolamento CPO, art. 3).'));
C.push(h2('Risorse e personale'));
C.push(tiles([
  [`${it(D.cpo.budget)} €`, 'budget del CPO nel bilancio di previsione 2026, per le attività formative. È il primo stanziamento proprio.'],
  ['2', 'dipendenti dell’Ordine, entrambe donne, nell’Ufficio Segreteria: seguono anche le attività del CPO.'],
], 2));
C.push(p(`Personale di supporto: ${D.cpo.personale}. Fino al 2025 il Comitato ha operato senza budget, a titolo gratuito o con sponsor. Lo stanziamento è in linea con altri Ordini: anche Ancona, nel bilancio di genere 2025, indica 1.500 euro per il proprio CPO.`, { spacing: { before: 140, after: 140, line: 340 } }));

// ================= 4 ATTIVITÀ
C.push(h1('4', 'Rendicontazione dell’attività'));
C.push(h2('2025: il Comitato 2022-2026'));
C.push(p(D.attivita_2025));
C.push(h2('2026: il primo anno del nuovo Comitato'));
C.push(box('Il convegno del 19 giugno 2026', [
  [run('“Comunicazione e leadership di genere: differenze tra universo maschile e femminile”', { bold: true })],
  'Cesenatico, Museo della Marineria, ore 15-18. Relatrice Roberta Fabbri. Organizzato con l’Ordine dei Consulenti del Lavoro di Forlì-Cesena (delibera del 27 marzo 2026) e con la Fondazione; accreditato dal Consiglio Nazionale anche per la deontologia; 100 posti (66 per commercialisti di Forlì e Ravenna, 34 per consulenti del lavoro); contributo di uno sponsor privato di 500 euro.',
]));
C.push(spazio());
C.push(b('Nomina del Comitato (23 febbraio 2026) e avvio della collaborazione con il CPO dell’Ordine di Ravenna e con gli altri Ordini professionali della provincia.'));
C.push(b('Partecipazione al convegno “Zero disparità, più produttività – La certificazione della parità di genere” (Cesena, 4 giugno 2026), con Camera di Commercio della Romagna e Consigliera di Parità regionale.'));
C.push(b('Rete regionale: rubrica dei Presidenti dei CPO dell’Emilia-Romagna e coinvolgimento dei Consiglieri delegati (marzo 2026); assemblea dei Presidenti a Bologna (23 settembre 2026), che ha eletto la rappresentante regionale nel Comitato Nazionale; riunione del coordinamento regionale del 24 novembre 2026.'));
C.push(b('Check-up dei bilanci di genere 2023 e 2024 e confronto con i bilanci di oltre venti Ordini.'));
C.push(b('Raccolta dei dati per questo bilancio: Ordine e Fondazione (28 settembre), OCC Romagna (5 ottobre), Cassa Ragionieri, Tribunale e Camera di Commercio (12 ottobre).'));
C.push(b('Preparato il questionario breve agli iscritti, che sarà inviato nel 2026-2027 (capitolo 8).'));
C.push(h2('Il programma del mandato'));
C.push(p('Dai verbali del Comitato: un punto d’ascolto per iscritte e iscritti, progetti di educazione fiscale nelle scuole, un desk per i tirocinanti, formazione sulle pari opportunità e la rete con gli altri Ordini della provincia.'));

// ================= 5 METODO
C.push(h1('5', 'Metodo di redazione'));
C.push(p('Seguiamo le Linee guida del CPO nazionale del 14 marzo 2024 (Informativa CNDCEC n. 31/2024). Il bilancio è redatto dal Comitato Pari Opportunità, è presentato all’assemblea degli iscritti e pubblicato sul sito dell’Ordine.'));
C.push(h2('Fonti'));
C.push(table(['Dato', 'Fonte', 'Riferimento'], [
  ['Iscritti, Elenco speciale, praticanti, flussi', 'Albo e Registro del tirocinio (Segreteria)', '2021-2025'],
  ['Consiglio, organi, commissioni, eventi, esoneri', 'Segreteria dell’Ordine', '2022-2026'],
  ['Redditi e volume d’affari', 'Cassa Dottori (CNPADC) e Cassa Ragionieri (CNPR)', 'Redditi 2023 e 2024'],
  ['Incarichi ai gestori della crisi', 'OCC Romagna', '2021-2025'],
  ['Cariche nelle imprese', 'Camera di commercio della Romagna', '2021-2025'],
  ['Spese', 'Bilancio della Fondazione; previsione 2026 dell’Ordine', '2025-2026'],
  ['Confronto nazionale', 'Bilancio di genere CN 2024 e questionario CNPO 2025', '31/12/2024'],
], [3300, 3906, 2100]));
C.push(h2('Regole di calcolo'));
C.push(b('Divario di reddito = (reddito medio uomini − reddito medio donne) / reddito medio uomini, sulle medie ponderate fornite dalla Cassa.'));
C.push(b('Tutela dei dati: solo dati aggregati; non mostriamo gruppi con meno di 5 persone quando potrebbero renderle riconoscibili.'));
C.push(b('Raccordo con il bilancio 2024: aggiungiamo i praticanti per genere, il divario di reddito di Forlì (prima solo regionale e nazionale) e le sezioni dell’Albo per genere.'));

// ================= 6 CONTESTO
C.push(h1('6', 'Analisi del contesto'));
C.push(h2('6.1 Iscritte e iscritti'));
C.push(tiles([
  [String(tot25), 'iscritti al 31/12/2025, tra Albo ed Elenco speciale'],
  [pct(f25, tot25), `donne (${f25}). Italia 33,9%, Emilia-Romagna 42%`],
  [`${it(I.eta_media['2025'][0], 0)} / ${it(I.eta_media['2025'][1], 0)}`, 'età media donne / uomini'],
]));
const sezRow = (k) => { const a = I.sezioni_2024[k], c = I.sezioni_2025[k]; return [k, `${a[0]} / ${a[1]}`, `${c[0]} / ${c[1]}`, pct(c[0], c[0] + c[1])]; };
C.push(spazio());
C.push(table(['Sezione', '2024 donne / uomini', '2025 donne / uomini', '% donne 2025'], [
  ...Object.keys(I.sezioni_2025).map(sezRow),
  ['Totale', '299 / 348', `${f25} / ${m25}`, pct(f25, tot25)],
], [2600, 2300, 2300, 2106], { boldLast: true }));
fonte('Albo dell’Ordine, elaborazione CPO. Nella sezione B e nell’Elenco speciale i numeri sono piccoli.');
const st = I.serie_totali;
C.push(table(['Anno', 'Albo', 'Elenco speciale', 'Totale', 'Donne', '% donne'], [
  ['2020', st['2020'][0], st['2020'][1], st['2020'][0] + st['2020'][1], '—', '—'],
  ...['2021', '2022', '2023', '2024', '2025'].map(y => { const g = I.per_genere[y]; const t = g[0] + g[1]; return [y, st[y][0], st[y][1], t, g[0], pct(g[0], t)]; }),
], [1200, 1400, 1800, 1400, 1600, 1906]));
fonte('bilanci di genere 2023 e 2024 (2020); Albo dell’Ordine 2021-2025.');
C.push(p('In cinque anni la quota di donne è passata dal 44,9% al 45,7%. Gli iscritti sono stabili intorno ai 650: le donne sono la stessa cifra di cinque anni fa (292), gli uomini 12 in meno.'));
C.push(h2('6.2 L’età: la parità arriva dai più giovani'));
C.push(graf('eta_genere.png'));
const ce = I.classi_eta_2025;
C.push(table(['Classe d’età', 'Donne', 'Uomini', '% donne'], Object.entries(ce).map(([k, v]) => [k, v[0], v[1], pct(v[0], v[0] + v[1])]), [3300, 2000, 2000, 2006]));
fonte('Albo dell’Ordine al 31/12/2025, classi del questionario nazionale CNPO.');
C.push(box('In sintesi', `Sotto i 56 anni le donne sono la maggioranza in ogni classe; tra gli over 65 sono poco più di una su quattro. Con il ricambio generazionale l’Albo arriverà alla parità numerica nei prossimi anni.`));
C.push(h2('6.3 Tirocinanti'));
C.push(graf('praticanti.png'));
C.push(table(['Anno', 'Totale', 'Donne', 'Uomini', '% donne'], [
  ['2020', pr['2020'], '—', '—', '—'],
  ...['2021', '2022', '2023', '2024', '2025'].map(y => [y, pr[y][0] + pr[y][1], pr[y][0], pr[y][1], pct(pr[y][0], pr[y][0] + pr[y][1])]),
], [1600, 1700, 2000, 2000, 2006]));
fonte('Registro del tirocinio. Per il 2023 il bilancio 2024 riportava 28 praticanti, l’estrazione ne conta 29.');
C.push(p('Negli ultimi due anni sei praticanti su dieci sono donne. È la base della professione di domani: il punto è che restino e crescano.'));
C.push(h2('6.4 Iscrizioni e cancellazioni'));
C.push(table(['Anno', 'Iscrizioni donne / uomini', 'Cancellazioni donne / uomini', 'Motivi'], D.flussi_ufficiali, [1200, 2300, 2400, 3406]));
fonte('Segreteria dell’Ordine. Il motivo della cancellazione manca per 8 casi del 2023.');
const u = D.flussi_u50.totale;
C.push(box('In sintesi', `In cinque anni sono entrate 29 donne e 23 uomini, ma sono uscite 29 donne e 34 uomini. Chi lascia presto è soprattutto donna: tra le cancellazioni sotto i 50 anni le donne sono ${u[0]} e gli uomini ${u[1]}. Anche a livello nazionale il tasso di cancellazione delle donne è più alto (1,92% contro 1,23% nel 2024).`, FILL_ARANCIO, ARANCIO));
C.push(h2('6.5 Cariche istituzionali'));
C.push(graf('organi.png', 560));
C.push(table(['Organo', 'Donne', 'Uomini', '% donne'], [
  ['Consiglio 2022-2026', 4, 7, pct(4, 11)],
  ['Consiglio 2026-2030', ...D.consiglio['2026-2030'], pct(D.consiglio['2026-2030'][0], 11)],
  ['CPO 2026-2030', D.cpo.componenti[0], D.cpo.componenti[1], pct(D.cpo.componenti[0], 7)],
  ...D.organi.map(o => [o[0], o[1], o[2], pct(o[1], o[1] + o[2])]),
], [4300, 1600, 1600, 1806]));
fonte('Segreteria dell’Ordine. Consiglio 2022-2026 in carica dal 25/2/2022, Consiglio 2026-2030 dal 23/2/2026.');
C.push(p(`Nel Consiglio 2022-2026 le donne erano il 36%, dieci punti sotto la loro presenza tra gli iscritti. Nel Consiglio 2026-2030 sono il 45%, in linea con l’Albo, e tre cariche su quattro sono di donne: la Presidente, dott.ssa Sara Pennacchi, la Segretaria e la Tesoriera. Cariche 2022-2026: ${D.consiglio.cariche_2022}.`));
C.push(p('Negli organi di controllo e di disciplina le donne sono la maggioranza o quasi. Nelle commissioni le componenti sono quasi la metà, ma nel 2026-2030 tre presidenze su quattro sono andate a uomini.'));
C.push(h2('6.6 Genitorialità e cura'));
C.push(table(['Esoneri dalla formazione 2025', 'Donne', 'Uomini'], D.esoneri, [4900, 2200, 2206]));
fonte('Segreteria dell’Ordine.');
C.push(p('Nel 2024, in Italia, le indennità di maternità richieste alle Casse sono state 778 e quelle di paternità 1.452 (Bilancio di genere CN 2024). Il confronto con i dati della Cassa per Forlì completerà il quadro nel bilancio 2026.'));
C.push(h2('6.7 Formazione: chi parla agli eventi'));
C.push(tiles([
  [D.eventi_kpi.pct_relatrici, 'relatrici sul totale degli interventi (74 su 267)'],
  [D.eventi_kpi.solo_uomini, 'eventi in cui hanno parlato solo uomini; mai il contrario'],
  ['46%', 'donne tra gli iscritti: il termine di confronto'],
]));
C.push(spazio());
C.push(table(['', 'Eventi', 'Relatrici', 'Relatori', '% relatrici', 'Panel di soli uomini'], D.formazione_eventi, [1700, 1300, 1400, 1400, 1600, 1906]));
fonte('Segreteria dell’Ordine: 51 eventi del 2025. Un evento in aula e in webinar conta due volte.');

// ================= 7 DIVARIO
C.push(h1('7', 'Il divario retributivo: focus sul gender pay gap', VIOLA));
C.push(p('I dati sono quelli della Cassa Dottori Commercialisti per gli iscritti attivi del territorio dell’Ordine; i ragionieri iscritti alla Cassa Ragionieri sono trattati a parte.'));
C.push(graf('divario.png'));
const r = CA.reddito_medio, ia = CA.iscritti_attivi, dv = CA.divario_reddito;
const terr = ['Forlì', 'Emilia-Romagna', 'Italia'];
C.push(table(['', ...terr], [
  ['Iscritti attivi (% donne)', ...terr.map(t => `${it(ia[t][0])} (${it(ia[t][1], 1)}%)`)],
  ['Reddito medio donne', ...terr.map(t => `${it(r[t][0])} €`)],
  ['Reddito medio uomini', ...terr.map(t => `${it(r[t][1])} €`)],
  ['Divario di reddito, redditi 2024', ...terr.map(t => `${it(dv[t][1], 1)}%`)],
  ['Divario di reddito, redditi 2023', ...terr.map(t => `${it(dv[t][0], 1)}%`)],
  ['Divario di volume d’affari, 2024', ...terr.map(t => `${it(CA.divario_volume_affari[t], 1)}%`)],
], [3300, 2000, 2000, 2006], { headFill: VIOLA }));
fonte('CNPADC, dati comunicati nel 2024 e nel 2025.');
C.push(box('In sintesi', 'A Forlì una professionista guadagna in media meno della metà di un collega. In un anno il reddito medio degli uomini è cresciuto dell’11,9%, quello delle donne dello 0,5%: il divario è salito di 5 punti e oggi supera la media regionale e nazionale.'));
C.push(h2('Il divario per età'));
const re = CA.rapporto_eta;
C.push(table(['Reddito delle donne in % di quello degli uomini', ...terr], Object.entries(re).map(([k, v]) => [k, ...v.map(x => x + '%')]), [3800, 1800, 1800, 1906], { headFill: VIOLA }));
fonte('CNPADC, redditi 2024. Fino a 30 anni a Forlì gli iscritti attivi sono 19; oltre 65 anni le donne sono 8: valori da leggere con prudenza.');
C.push(p('Fino a 50 anni le professioniste di Forlì stanno meglio che in regione e in Italia. Il divario così alto nasce soprattutto tra i 51 e i 65 anni, dove si concentrano gli uomini con i redditi più alti. Il problema non è l’ingresso nella professione, ma la crescita a metà carriera.'));
C.push(h2('Cassa Ragionieri'));
C.push(table(['Iscritti attivi CNPR, Ordine di Forlì', 'Donne', 'Uomini', 'Divario'], D.cassa_cnpr, [4000, 1700, 1700, 1906], { headFill: VIOLA }));
fonte('Cassa Ragionieri (CNPR), dati comunicati nel 2025.');
if (D.cnpr_testo) [].concat(D.cnpr_testo).forEach(t => C.push(p(t)));

// ================= 8 QUESTIONARIO
C.push(h1('8', 'Questionario'));
C.push(h2('Un’azione per il 2026-2027'));
C.push(p('Quest’anno il bilancio non contiene ancora la voce diretta degli iscritti: il CPO, insediato a febbraio 2026, ha dato priorità alla raccolta dei dati ufficiali. Ha preparato un questionario breve e anonimo su conciliazione tra lavoro e famiglia, carriera e incarichi, esperienze di discriminazione e, per la prima volta tra i bilanci di genere esaminati, DSA e inclusione.'));
C.push(b('Invio tramite la Segreteria a iscritte, iscritti e tirocinanti, con un promemoria dopo due settimane.'));
C.push(b('Obiettivo: almeno un quarto degli iscritti (circa 160 risposte).'));
C.push(b('Risultati solo aggregati, senza gruppi sotto le 5 risposte, nel bilancio di genere 2026; alcune domande uguali a quelle nazionali per il confronto.'));
const Q = D.cnpo_2025;
C.push(h2('Il quadro nazionale: questionario CNPO 2025'));
C.push(p(Q.intro));
C.push(table(['Risposta', 'Donne', 'Uomini'], Q.righe, [6100, 1600, 1606]));
fonte('CPO Nazionale e Fondazione Nazionale di Ricerca dei Commercialisti, “Conciliazione vita-lavoro e gender pay gap. Questionario CNPO 2025”, 25 novembre 2025. Percentuali sulle risposte di ciascun genere.');
Q.testo.slice(0, 2).forEach(t => C.push(p(t)));

// ================= 9 APPROFONDIMENTI
C.push(h1('9', 'Analisi facoltative: incarichi, cariche, spese'));
C.push(p('Le Linee guida suggeriscono di guardare oltre l’Albo: chi riceve gli incarichi e chi siede nei collegi sindacali delle imprese. Pochi Ordini lo fanno. Abbiamo chiesto i dati al Tribunale, all’OCC Romagna e alla Camera di Commercio.'));
C.push(h2('9.1 Tribunale di Forlì'));
C.push(p(D.tribunale_nota));
C.push(h2('9.2 OCC Romagna: i gestori della crisi'));
C.push(graf('occ.png'));
C.push(table(['', 'Donne', 'Uomini', '% donne'], D.occ, [4300, 1600, 1600, 1806]));
fonte('OCC Romagna, prospetto compilato il 30/10/2026. Gestori iscritti all’Ordine di Forlì. Il dato 2022 è in corso di verifica con l’OCC.');
if (D.occ_testo) [].concat(D.occ_testo).forEach(t => C.push(p(t)));
C.push(h2('9.3 Collegi sindacali e revisione nelle imprese'));
C.push(table(['Carica, imprese registrate in provincia', 'Donne 2025', 'Uomini 2025', '% donne 2025', '% donne 2021'], D.cciaa, [3300, 1450, 1450, 1500, 1606]));
fonte('Elaborazione a cura dell’Osservatorio economico e sociale della Camera di commercio della Romagna - Forlì-Cesena e Rimini. Comprende tutte le persone con la carica, non solo i commercialisti.');
[].concat(D.cciaa_testo).forEach(t => C.push(p(t)));
C.push(graf('cciaa.png', 560));
C.push(table(['Carica, imprese attive della provincia', 'Donne', 'Uomini', '% donne'], D.cciaa_attive, [4700, 1500, 1500, 1606]));
fonte('Infocamere (StockView), elaborazione Camera di commercio della Romagna, 4° trimestre 2025. Cariche di persone fisiche in imprese attive con sede in provincia.');
C.push(p(D.cciaa_attive_testo));
C.push(h2('9.4 Ruoli nelle procedure concorsuali'));
C.push(table(['Ruolo registrato', 'Donne 2025', 'Uomini 2025', '% donne 2025', '% donne 2021'], D.cciaa_procedure, [3300, 1450, 1450, 1500, 1606]));
fonte('Camera di commercio della Romagna. Indicatore indiretto: persone indicate nel Registro delle imprese, non incarichi conferiti nell’anno.');
C.push(p(D.cciaa_proc_testo));
C.push(h2('9.5 Le spese con la lente di genere'));
C.push(p('È il cuore del gender budgeting, ripreso dalle Linee guida, ma quasi nessun Ordine lo applica. Il metodo divide le spese in tre gruppi:'));
C.push(b('neutre: non toccano in modo diverso donne e uomini (affitto, utenze, software, quote al Consiglio Nazionale);'));
C.push(b('dirette a ridurre le diseguaglianze di genere: attività del CPO, progetti nelle scuole;'));
C.push(b('attività socio-culturali: da leggere guardando chi vi partecipa.'));
C.push(p(D.spese_note));
C.push(p('Per il 2025 la riclassificazione completa delle spese dell’Ordine non è stata possibile, perché il consuntivo per capitolo non era disponibile in tempo. Il metodo sarà esteso all’Ordine dal bilancio di genere 2026, quando il CPO rendiconterà anche il suo primo stanziamento.'));

// ================= 10 OBIETTIVI
C.push(h1('10', 'Obiettivi per il 2027'));
C.push(p('Ogni obiettivo ha un indicatore, un valore di partenza e un traguardo: nel bilancio 2026 diremo se l’abbiamo raggiunto. Gli obiettivi sono una proposta del CPO da condividere con il Consiglio.'));
C.push(table(['Obiettivo', 'Indicatore', 'Partenza 2025', 'Traguardo 2027'], [
  ['Bilancio completo', 'Contenuti minimi delle Linee guida coperti', 'circa 50% nel 2024', '100%'],
  ['Più relatrici', '% relatrici negli eventi dell’Ordine', D.eventi_kpi.pct_relatrici, 'almeno 40%'],
  ['“No women no panel”', 'Eventi con soli relatori uomini', D.eventi_kpi.solo_uomini, '0 tra gli eventi con 3 o più relatori, come regola proposta al Consiglio'],
  ['Presidenze delle commissioni', 'Commissioni presiedute da donne', '1 su 4', 'almeno 2 su 4 al prossimo rinnovo'],
  ['Il CPO ha risorse', 'Budget assegnato', 'nessuno (1.500 € nel 2026)', 'stanziamento confermato nel 2027'],
  ['Spese con la lente di genere', 'Bilancio riclassificato in tre gruppi', 'solo Fondazione', 'Ordine e Fondazione'],
  ['Ascoltare gli iscritti', 'Risposte al questionario', 'non ancora rilevato', 'almeno 25% degli iscritti'],
  ['Divario di reddito', 'Divario Forlì rispetto alla media regionale', '+6,1 punti', 'sotto la media regionale entro il 2028'],
  ['Bilancio accessibile', 'Sintesi facile e PDF per sintesi vocale', 'sì', 'sì, anche versione audio'],
], [2300, 2700, 2000, 2306], { headFill: VIOLA }));
C.push(p('Il divario di reddito dipende da molti fattori fuori dal controllo dell’Ordine. Il CPO lo segue con azioni mirate alla fascia di metà carriera: formazione su specializzazioni e incarichi, mentoring, rete con la Consigliera di Parità.', { spacing: { before: 160, after: 140, line: 340 } }));

// ================= 11 CONCLUSIONI
C.push(h1('11', 'Conclusioni'));
C.push(box('Il messaggio di questo bilancio', [
  'Forlì è un Ordine con molte donne, più giovani dei colleghi, e con un’ampia base di praticanti. La parità nei numeri è vicina.',
  'Non lo è nei redditi: il divario è alto e cresce, soprattutto a metà carriera. Non lo è ancora del tutto nella rappresentanza, anche se il nuovo Consiglio è guidato da una Presidente.',
]));
C.push(spazio());
C.push(p('Con questo bilancio abbiamo fatto due passi in più: misurare dove si decidono le carriere (incarichi, cariche, formazione) e iniziare a leggere con la lente di genere le spese, partendo dalla Fondazione. Il prossimo anno sapremo se gli obiettivi fissati sono stati raggiunti.'));
C.push(p('Forlì, ____________________ 2026', { spacing: { before: 400, after: 300 } }));
C.push(p(run('Il Comitato Pari Opportunità', { bold: true, color: BLU })));
D.cpo.nomi.forEach(([n, rr]) => C.push(p(`${n} – ${rr}    ________________________________`, { spacing: { after: 260 } })));

// ================= APPENDICE
C.push(h1('', 'Appendice – Stato dei dati', '6B7280'));
C.push(table(['N.', 'Dato', 'Fonte', 'Stato'], [
  ['1', 'Iscrizioni e cancellazioni 2021-2025', 'Segreteria', 'inserito'],
  ['2', 'Iscritti e praticanti per genere 2021-2023', 'Segreteria', 'inserito'],
  ['3', 'Consiglio, cariche, organi, commissioni', 'Segreteria', 'inserito'],
  ['4', 'Eventi formativi 2025, relatrici e relatori', 'Segreteria', 'inserito'],
  ['5', 'Esoneri per maternità, paternità e cura', 'Segreteria', 'inserito'],
  ['6', 'Consuntivo 2025 per capitolo', 'Tesoriera', 'non disponibile in tempo: dal bilancio 2026'],
  ['7', 'Spese socio-culturali della Fondazione', 'Fondazione', 'inserito'],
  ['8', 'Redditi Cassa Ragionieri', 'CNPR', 'inserito'],
  ['9', 'Incarichi del Tribunale', 'Tribunale di Forlì', 'non concessi'],
  ['10', 'Gestori e incarichi', 'OCC Romagna', 'inserito (2022 da verificare)'],
  ['11', 'Sindaci, revisori e cariche nelle imprese', 'Camera di Commercio', 'inserito'],
  ['–', 'Questionario agli iscritti', 'CPO', 'azione 2026-2027'],
], [700, 4000, 2300, 2306], { headFill: '6B7280' }));

const doc = new Document({
  creator: 'CPO ODCEC Forlì', title: 'Bilancio di genere 2025 – ODCEC Forlì', description: 'Bozza di lavoro, versione 2',
  features: { updateFields: true },
  styles: { default: { document: { run: { font: FONT, size: 21, color: INK } } } },
  numbering: { config: [{ reference: 'punti', levels: [{ level: 0, format: LevelFormat.BULLET, text: '■', alignment: AlignmentType.LEFT, style: { run: { color: AZZ, size: 14 }, paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AZZ, space: 4 } }, children: [run('Bilancio di genere 2025  ·  ODCEC Forlì  ·  Comitato Pari Opportunità', { size: 15, color: GREY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, bold: true, color: BLU })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[2] || path.join(__dirname, 'Bilancio_di_genere_2025_ODCEC_Forli_v2.docx'), buf); console.log('ok'); });
