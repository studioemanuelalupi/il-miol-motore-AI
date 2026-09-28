const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType,
  ShadingType, LevelFormat, Footer, PageNumber, HeadingLevel,
} = require('docx');

const FONT = 'Verdana', INK = '1F2A37', ACCENT = '2E5E8C', GREY = '4B5563';
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, color: INK, size: 22, ...o });
const p = (c, o = {}) => new Paragraph({ children: Array.isArray(c) ? c : [typeof c === 'string' ? run(c) : c], spacing: { after: 140, line: 340 }, ...o });
const right = (t, o = {}) => p(run(t, o), { alignment: AlignmentType.LEFT, indent: { left: 5000 }, spacing: { after: 0, line: 300 } });
const num = (t, ref = 'num') => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 320 }, children: typeof t === 'string' ? [run(t)] : t });
const bul = (t) => num(t, 'punti');
const h = (t, brk = false) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: brk, keepNext: true, spacing: { before: 240, after: 140 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 26, color: ACCENT })] });
const h3 = (t) => new Paragraph({ keepNext: true, spacing: { before: 200, after: 100 }, children: [run(t, { bold: true, color: ACCENT })] });

function table(head, rows, widths) {
  const cell = (t, w, hd) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: hd ? { type: ShadingType.CLEAR, color: 'auto', fill: 'DCE7F2' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: String(t).split('\n').map(l => new Paragraph({ children: [run(l, { size: 18, bold: hd })] })),
  });
  return new Table({
    width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: head.map((x, i) => cell(x, widths[i], true)) }),
      ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((x, i) => cell(x, widths[i], false)) }))],
  });
}

const C = [];
// ---------- intestazione
C.push(p(run('COMITATO PARI OPPORTUNITÀ', { bold: true, color: ACCENT }), { spacing: { after: 0 } }));
C.push(p(run('Ordine dei Dottori Commercialisti e degli Esperti Contabili di Forlì', { color: GREY }), { spacing: { after: 360 } }));
C.push(right('Alla Presidente'));
C.push(right('Dott.ssa Sara Pennacchi'));
C.push(right('e p.c. alle Consigliere e ai Consiglieri'));
C.push(right('Consiglio dell’Ordine dei Dottori Commercialisti'));
C.push(right('e degli Esperti Contabili di Forlì'));
C.push(right('Via Silver Sirotti, 31 – 47122 Forlì'));
C.push(p('', { spacing: { after: 240 } }));
C.push(p('Forlì, [data]'));
C.push(p([run('Oggetto: ', { bold: true }), run('Bilancio di genere 2025 – richiesta di dati e di supporto al Comitato Pari Opportunità')]));

C.push(p('Cara Presidente, care Consigliere e cari Consiglieri,'));
C.push(p('il Comitato Pari Opportunità, nominato dal Consiglio il 23 febbraio 2026, sta preparando il Bilancio di genere dell’anno 2025, che il Regolamento nazionale dei CPO affida alla sua competenza (art. 3, lett. i) e che le Linee guida del Comitato Nazionale Pari Opportunità (Informativa CNDCEC n. 31 del 14 marzo 2024) chiedono di presentare ogni anno all’assemblea degli iscritti e di pubblicare sul sito dell’Ordine.'));
C.push(p('Nel lavoro preparatorio abbiamo confrontato i bilanci di Forlì del 2023 e del 2024 con le Linee guida e con i bilanci di altri Ordini. Ne è emerso che alcuni contenuti minimi non sono ancora presenti, soprattutto perché mancavano i dati: la serie storica per genere, le cariche del Consiglio e delle commissioni, i praticanti per genere, i movimenti di iscrizione e cancellazione e un divario di reddito calcolato sui dati di Forlì.'));
C.push(p('Per colmare queste lacune, e perché diverse informazioni sono in possesso di enti esterni che rispondono più facilmente a una richiesta istituzionale, Vi chiediamo di:'));
[
  'mettere a disposizione del Comitato i dati interni dell’Ordine indicati nella Parte A dell’allegato, tramite la Segreteria e la Tesoreria;',
  'inoltrare a nome dell’Ordine le richieste agli enti esterni indicati nella Parte B (Casse di previdenza, Tribunale di Forlì, Organismo di composizione della crisi, Camera di Commercio della Romagna). In allegato trovate un fac-simile di lettera per ciascun ente, da adattare e firmare;',
  'indicare, come previsto dalle Linee guida, il budget che il Consiglio intende assegnare al Comitato per il 2026 e il nominativo della persona della Segreteria che potrà fare da riferimento per il Comitato.',
].forEach(t => C.push(num(t)));

C.push(p('Per tutela della riservatezza chiediamo solo dati aggregati o anonimi (senza nomi né codici fiscali). Il Comitato pubblicherà unicamente tabelle di sintesi e non mostrerà dati riferiti a meno di cinque persone. Dove i dati sono disponibili, sarebbe utile riceverli in formato Excel.'));
C.push(p('Per consentire la presentazione del bilancio all’assemblea di approvazione del conto consuntivo, proponiamo queste scadenze:'));
C.push(bul('dati interni (Parte A): entro il 30 novembre 2026;'));
C.push(bul('invio delle richieste agli enti esterni (Parte B): entro il 31 ottobre 2026, con riscontro entro il 31 gennaio 2027.'));
C.push(p('Il Comitato resta a disposizione per illustrare la richiesta in una prossima seduta del Consiglio e per ogni chiarimento. Ringraziamo fin d’ora il Consiglio e la Segreteria per la collaborazione.'));
C.push(p('Con i migliori saluti,', { spacing: { before: 200, after: 480 } }));
C.push(right('Dott.ssa Emanuela Lupi', { bold: true }));
C.push(right('Presidente del Comitato Pari Opportunità'));
C.push(right('ODCEC di Forlì'));
C.push(p('', { spacing: { after: 200 } }));
C.push(p(run('Allegati: elenco dei dati richiesti (Parti A e B); fac-simili di richiesta agli enti esterni.', { size: 18, color: GREY })));

// ---------- allegato dati
C.push(h('Allegato – Elenco dei dati richiesti', true));
C.push(p(run('Periodo: dove non diversamente indicato, dati al 31 dicembre di ciascun anno dal 2021 al 2025. Tutti i dati sono richiesti suddivisi per genere.', { size: 20 })));
C.push(h3('Parte A – Dati interni dell’Ordine'));
C.push(table(['N.', 'Dato richiesto', 'Dettaglio', 'Ufficio'], [
  ['A1', 'Iscritti all’Albo (sez. A e B) e all’Elenco speciale', 'Per genere, anno di nascita e sezione, 2021-2025. Per 2024 e 2025 l’estrazione è già disponibile', 'Segreteria'],
  ['A2', 'Nuove iscrizioni e cancellazioni', 'Per genere, anno di nascita, sezione e motivo (trasferimento, dimissioni, pensione, decesso, morosità, altro), 2021-2025', 'Segreteria'],
  ['A3', 'Praticanti', 'Iscritti al Registro al 31/12 per genere, anno di nascita e sezione; tirocini conclusi e interrotti, 2021-2025', 'Segreteria'],
  ['A4', 'Consiglio dell’Ordine', 'Composizione per genere e cariche (Presidente, Vicepresidente, Segretario, Tesoriere) per le consigliature 2017-2021, 2022-2025 e 2026-2030', 'Segreteria'],
  ['A5', 'Altri organi', 'Consiglio di disciplina, Revisore o Collegio dei revisori, delegati: composizione per genere e cariche, stesse consigliature', 'Segreteria'],
  ['A6', 'Commissioni di studio e gruppi di lavoro', 'Componenti e presidenti per genere, 2025 e 2026', 'Segreteria'],
  ['A7', 'Candidature alle elezioni 2026', 'Numero di candidate e candidati ed elette ed eletti, per genere', 'Segreteria'],
  ['A8', 'Formazione professionale continua 2025', 'Elenco eventi organizzati con: numero di relatrici e relatori, moderatrici e moderatori, orario, modalità (presenza/online), partecipanti per genere se disponibili', 'Segreteria / Fondazione'],
  ['A9', 'Esoneri dalla formazione continua', 'Numero di esoneri per maternità, paternità, malattia e assistenza a familiari, per genere, 2021-2025', 'Segreteria'],
  ['A10', 'Designazioni e nomine fatte dall’Ordine', 'Incarichi su designazione dell’Ordine (commissioni, enti, terne), per genere, 2021-2025', 'Segreteria'],
  ['A11', 'Personale dipendente dell’Ordine', 'Numero per genere, livello, tipo di contratto, part-time, lavoro agile', 'Segreteria'],
  ['A12', 'Conto consuntivo 2025 e preventivo 2026', 'Dettaglio delle spese per capitolo, per la riclassificazione di genere (spese neutre, sensibili, dirette a ridurre le diseguaglianze)', 'Tesoreria'],
  ['A13', 'Budget e referente del CPO', 'Importo assegnato per il 2026 (o motivazioni se non assegnato) e nominativo del referente in Segreteria', 'Consiglio'],
], [600, 2500, 4526, 1400]));

C.push(h3('Parte B – Dati da richiedere a enti esterni'));
C.push(table(['N.', 'Ente', 'Dato richiesto', 'Note'], [
  ['B1', 'Cassa Nazionale di Previdenza e Assistenza dei Dottori Commercialisti', 'Iscritti attivi, reddito medio (Irpef) e volume d’affari medio (IVA) per genere e classe d’età, Ordine di Forlì, comunicazioni 2021-2023. Pensioni medie per genere. Indennità di maternità e contributi di paternità erogati a iscritti di Forlì, 2021-2025', 'I dati 2024 e 2025 sono già disponibili'],
  ['B2', 'Cassa Nazionale di Previdenza e Assistenza dei Ragionieri e Periti Commerciali', 'Stessi dati del punto B1 per gli iscritti all’Ordine di Forlì', 'Necessari per un divario di reddito completo'],
  ['B3', 'Tribunale di Forlì', 'Incarichi conferiti a iscritti all’Ordine di Forlì nel 2021-2025, per genere e tipo (curatore/liquidatore giudiziale, commissario giudiziale, CTU, custode, delegato alle vendite, altri), con numero e, se possibile, compensi liquidati', 'Analisi indicata dalle Linee guida tra quelle facoltative'],
  ['B4', 'Organismo di composizione della crisi (OCC)', 'Gestori della crisi iscritti all’Ordine: numero per genere e incarichi assegnati, 2021-2025', ''],
  ['B5', 'Camera di Commercio della Romagna', 'Cariche di sindaco, revisore legale e componente di organi di controllo nelle imprese della provincia di Forlì-Cesena, per genere, al 31/12/2025', 'Analisi indicata dalle Linee guida tra quelle facoltative'],
], [600, 2300, 4426, 1700]));

// ---------- fac-simili
C.push(h('Fac-simili di richiesta agli enti esterni', true));
C.push(p(run('Testi da riportare su carta intestata dell’Ordine e da firmare a cura della Presidente.', { size: 20, color: GREY })));

const fac = (dest, oggetto, corpo) => {
  C.push(h3(dest));
  C.push(p([run('Oggetto: ', { bold: true }), run(oggetto)]));
  corpo.forEach(t => C.push(p(t)));
  C.push(p('Ringraziando per la collaborazione, si porgono cordiali saluti.'));
  C.push(p(run('La Presidente – Dott.ssa Sara Pennacchi', { italics: false }), { spacing: { after: 360 } }));
};
const premessa = 'L’Ordine dei Dottori Commercialisti e degli Esperti Contabili di Forlì, tramite il proprio Comitato Pari Opportunità, redige ogni anno il Bilancio di genere secondo le Linee guida del Consiglio Nazionale (Informativa n. 31/2024). Il documento viene presentato all’assemblea degli iscritti e pubblicato sul sito dell’Ordine.';
const privacy = 'I dati saranno utilizzati esclusivamente in forma aggregata e anonima, senza pubblicare informazioni riferite a meno di cinque persone. Si chiede cortesemente di trasmetterli, ove possibile, in formato Excel entro il 31 gennaio 2027 all’indirizzo [PEC/e-mail dell’Ordine].';

fac('1. Cassa Nazionale di Previdenza e Assistenza dei Dottori Commercialisti', 'Richiesta di dati aggregati per il Bilancio di genere dell’Ordine di Forlì', [
  premessa,
  'Si chiede di ricevere, per gli iscritti alla Cassa appartenenti all’Ordine di Forlì: numero di iscritti attivi, reddito netto professionale medio e volume d’affari medio, per genere e classe d’età, relativi alle comunicazioni degli anni 2021, 2022 e 2023; importo medio delle pensioni erogate per genere; numero e importo medio delle indennità di maternità e dei contributi di paternità erogati negli anni 2021-2025.',
  privacy,
]);
fac('2. Cassa Nazionale di Previdenza e Assistenza dei Ragionieri e Periti Commerciali', 'Richiesta di dati aggregati per il Bilancio di genere dell’Ordine di Forlì', [
  premessa,
  'Si chiede di ricevere, per gli iscritti alla Cassa appartenenti all’Ordine di Forlì: numero di iscritti attivi, reddito professionale medio e volume d’affari medio, per genere e classe d’età (fino a 30, 31-40, 41-50, 51-65, oltre 65 anni), relativi alle comunicazioni degli anni 2021-2025; importo medio delle pensioni erogate per genere; numero delle indennità di maternità erogate.',
  privacy,
]);
fac('3. Presidente del Tribunale di Forlì', 'Richiesta di dati aggregati sugli incarichi conferiti, per il Bilancio di genere dell’Ordine', [
  premessa,
  'Le Linee guida suggeriscono di analizzare per genere gli incarichi conferiti dal Tribunale. Si chiede pertanto di ricevere, per gli anni 2021-2025, il numero di incarichi conferiti a iscritte e iscritti all’Ordine di Forlì, distinti per genere e per tipologia (curatore o liquidatore giudiziale, commissario giudiziale, consulente tecnico d’ufficio, custode, delegato alle vendite, altri incarichi), con l’indicazione, ove disponibile, dei compensi complessivamente liquidati.',
  'L’Ordine è disponibile a concordare con le Cancellerie le modalità di estrazione più semplici, anche fornendo l’elenco degli iscritti con l’indicazione del solo genere.',
  privacy,
]);
fac('4. Organismo di composizione della crisi (OCC)', 'Richiesta di dati aggregati per il Bilancio di genere dell’Ordine di Forlì', [
  premessa,
  'Si chiede di ricevere, per gli anni 2021-2025, il numero di gestori della crisi iscritti all’Ordine di Forlì e il numero di incarichi loro assegnati, distinti per genere.',
  privacy,
]);
fac('5. Camera di Commercio della Romagna – Forlì-Cesena e Rimini', 'Richiesta di elaborazione statistica dal Registro delle imprese per il Bilancio di genere', [
  premessa,
  'Le Linee guida suggeriscono di analizzare per genere le cariche di revisore e sindaco. Si chiede pertanto un’elaborazione statistica, al 31 dicembre 2025, del numero di cariche di sindaco effettivo, sindaco supplente, revisore legale e componente di organi di controllo nelle imprese con sede nella provincia di Forlì-Cesena, distinte per genere del titolare. Se tecnicamente possibile, si chiede di indicare separatamente le cariche ricoperte da persone residenti nella provincia.',
  privacy,
]);

const doc = new Document({
  creator: 'CPO ODCEC Forlì', title: 'Richiesta dati Bilancio di genere – CPO ODCEC Forlì',
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } } },
  numbering: { config: [
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] },
    { reference: 'punti', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] },
  ] },
  sections: [{
    properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run('Pagina ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('ok'); });
