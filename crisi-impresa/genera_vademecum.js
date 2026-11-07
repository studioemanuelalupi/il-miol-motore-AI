// Vademecum per il collaboratore: come si redigono ricorso e relazione del Gestore
// nelle 4 procedure da sovraindebitamento (D.Lgs. 14/2019 - CCII).
// Uso: node genera_vademecum.js [file.docx]
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, Footer, PageNumber, AlignmentType, BorderStyle,
  TableOfContents,
} = require('docx');

const FONT = 'Verdana', INK = '1F2A37', ACCENT = '2E5E8C', GREY = '4B5563';
const AZZURRO = 'DCE7F2', GIALLO = 'FFF4CC', VERDE = 'E3F1E5', GRIGIO = 'F1F3F5', ROSA = 'FBE4E4';
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, color: INK, size: 22, ...o });
const p = (c, o = {}) => new Paragraph({ children: Array.isArray(c) ? c : [typeof c === 'string' ? run(c) : c], spacing: { after: 120, line: 340 }, ...o });
// "Titolo: testo" mette in grassetto la parte prima dei due punti
const kids = (t, size = 22) => {
  const i = t.indexOf(': ');
  return i > 0 && i < 70 ? [run(t.slice(0, i + 1), { bold: true, size }), run(t.slice(i + 1), { size })] : [run(t, { size })];
};
const b = (t) => new Paragraph({ numbering: { reference: 'punti', level: 0 }, spacing: { after: 60, line: 320 }, children: kids(t) });
const b2 = (t) => new Paragraph({ numbering: { reference: 'punti', level: 1 }, spacing: { after: 40, line: 300 }, children: kids(t, 20) });
const h1 = (t, brk = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: brk, spacing: { after: 160 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 30, color: ACCENT })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 280, after: 120 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: ACCENT })] });
const h3 = (t) => new Paragraph({ keepNext: true, spacing: { before: 200, after: 80 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 22, color: INK })] });
const nota = (t) => p(run(t, { size: 18, color: GREY }));
const C = [];
const list = (arr) => arr.forEach(t => C.push(Array.isArray(t) ? b2(t[0]) : b(t)));
const par = (...arr) => arr.forEach(t => C.push(p(kids(t))));
const spazio = () => C.push(new Paragraph({ spacing: { after: 60 }, children: [] }));

const NO = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NOB = { top: NO, bottom: NO, left: NO, right: NO };
const LINE = (c = ACCENT) => ({ style: BorderStyle.SINGLE, size: 8, color: c });
const BOX = (c) => ({ top: LINE(c), bottom: LINE(c), left: LINE(c), right: LINE(c) });

function table(head, rows, widths) {
  const total = widths.reduce((a, c) => a + c, 0);
  const cell = (t, w, hd) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: hd ? { type: ShadingType.CLEAR, color: 'auto', fill: AZZURRO } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: String(t).split('\n').map(r => new Paragraph({ children: [run(r, { size: 18, bold: hd })] })),
  });
  const tr = [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], true)) })];
  rows.forEach(r => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false)) })));
  C.push(new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: tr }));
  spazio();
}

// riquadro colorato: Attenzione / Esempio / Ricorda
function box(titolo, righe, fill = GIALLO) {
  const children = [new Paragraph({ spacing: { after: 80 }, children: [run(titolo, { bold: true, color: ACCENT })] })];
  righe.forEach(r => children.push(new Paragraph({ spacing: { after: 60, line: 300 }, children: kids(r, 20) })));
  C.push(new Table({
    width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026],
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: 9026, type: WidthType.DXA }, borders: BOX(ACCENT), shading: { type: ShadingType.CLEAR, color: 'auto', fill }, margins: { top: 120, bottom: 120, left: 160, right: 160 }, children })] })],
  }));
  spazio();
}

// schema a blocchi verticale: [titolo, testo, colore?]; tra un blocco e l'altro una freccia
function flow(titolo, passi, w = 8000) {
  if (titolo) C.push(h3(titolo));
  const rows = [];
  passi.forEach((s, i) => {
    const [tit, txt, fill] = s;
    const ch = [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(tit, { bold: true, size: 20 })] })];
    if (txt) txt.split('\n').forEach(r => ch.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [run(r, { size: 18 })] })));
    rows.push(new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: w, type: WidthType.DXA }, borders: BOX(ACCENT), shading: { type: ShadingType.CLEAR, color: 'auto', fill: fill || AZZURRO }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: ch })] }));
    if (i < passi.length - 1) rows.push(new TableRow({ children: [new TableCell({ width: { size: w, type: WidthType.DXA }, borders: NOB, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('▼', { color: ACCENT, size: 22 })] })] })] }));
  });
  C.push(new Table({ alignment: AlignmentType.CENTER, width: { size: w, type: WidthType.DXA }, columnWidths: [w], rows }));
  spazio();
}

// albero decisionale: domanda a sinistra, risposta "SÌ" a destra, "NO" scende alla domanda successiva
function albero(domande, finale) {
  const W1 = 4600, W2 = 600, W3 = 3826, tot = W1 + W2 + W3;
  const cel = (w, ch, fill, borders) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: borders || NOB, shading: fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: ch });
  const para = (t, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER, children: [run(t, { size: 18, ...o })] });
  const rows = [];
  domande.forEach(([dom, si], i) => {
    rows.push(new TableRow({ cantSplit: true, children: [
      cel(W1, [para(dom, { bold: true })], GIALLO, BOX(ACCENT)),
      cel(W2, [para('SÌ ►', { bold: true, color: ACCENT })]),
      cel(W3, si.split('\n').map((r, j) => para(r, { bold: j === 0 })), VERDE, BOX(ACCENT)),
    ] }));
    rows.push(new TableRow({ children: [cel(W1, [para(i < domande.length - 1 || finale ? 'NO ▼' : '', { bold: true, color: ACCENT })]), cel(W2, [para('')]), cel(W3, [para('')])] }));
  });
  if (finale) rows.push(new TableRow({ cantSplit: true, children: [cel(W1, finale.split('\n').map((r, j) => para(r, { bold: j === 0 })), ROSA, BOX(ACCENT)), cel(W2, [para('')]), cel(W3, [para('')])] }));
  C.push(new Table({ width: { size: tot, type: WidthType.DXA }, columnWidths: [W1, W2, W3], rows }));
  spazio();
}

// fac-simile: testo in carattere monospazio su fondo grigio
function facsimile(righe) {
  const ch = righe.map(r => new Paragraph({ spacing: { after: 40, line: 280 }, children: [new TextRun({ text: r, font: 'Courier New', size: 18, color: INK, bold: /^[A-ZÀ-Ù0-9 .–—:()'’\/-]{6,}$/.test(r.trim()) })] }));
  C.push(new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: [new TableRow({ children: [new TableCell({ width: { size: 9026, type: WidthType.DXA }, borders: BOX(GREY), shading: { type: ShadingType.CLEAR, color: 'auto', fill: GRIGIO }, margins: { top: 140, bottom: 140, left: 180, right: 180 }, children: ch })] })] }));
  spazio();
}

// =====================================================================
// COPERTINA
// =====================================================================
C.push(p(run('Materiale di formazione interno – Studio / OCC', { size: 20, color: GREY })));
C.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 120 }, children: [new TextRun({ text: 'Vademecum del sovraindebitamento', font: FONT, bold: true, size: 40, color: ACCENT })] }));
C.push(p(run('Come si scrivono il ricorso e la relazione del Gestore della crisi nelle 4 procedure', { size: 24, color: GREY })));
C.push(p(run('Per Lorenzo – prima edizione, novembre 2026', { size: 20, color: GREY })));
spazio();
box('Come usare questo vademecum', [
  'Prima lettura: leggi tutto una volta, con calma, dall’inizio alla fine. Non serve capire tutto subito.',
  'Capitoli 1-4: sono le basi (parole, norme, scelta della procedura, raccolta dei documenti). Valgono per tutte le pratiche.',
  'Capitolo 5: insegna a fare i conti che servono in ogni ricorso e in ogni relazione.',
  'Capitoli 6-9: una scheda per ogni procedura, con schema dei passaggi, indice del ricorso e indice della relazione.',
  'Capitoli 10-13: confronto tra le procedure, errori da evitare, checklist prima del deposito e come lavorare con Claude.',
  'Appendice: fac-simile da cui partire. Sono tracce, non moduli da copiare: ogni pratica va scritta sul caso concreto.',
], AZZURRO);
box('Regola d’oro', [
  'Ogni bozza che prepari (con o senza Claude) la rivedo io prima che esca dallo studio. Nessun atto si firma o si deposita senza il mio controllo.',
  'Ogni numero deve avere un documento dietro: se non c’è il documento, il numero non si scrive (oppure si scrive “da verificare”).',
  'Le norme cambiano: il Codice della crisi è stato corretto più volte (ultimo grande correttivo: D.Lgs. 136/2024, in vigore dal 28 settembre 2024). Prima di citare un comma controlla il testo vigente su Normattiva. Dove in questo vademecum trovi la scritta [VERIFICARE] il punto è discusso o è cambiato di recente.',
], GIALLO);

C.push(h2('Indice'));
C.push(new TableOfContents('Indice', { hyperlink: true, headingStyleRange: '1-2' }));
nota('Se l’indice è vuoto: in Word clic destro sull’indice → “Aggiorna campo”.');

// =====================================================================
// 1. LE PAROLE
// =====================================================================
C.push(h1('1. Le parole che devi conoscere'));
par('Il sovraindebitamento ha un suo linguaggio. Queste sono le parole che troverai in ogni pratica. Se una parola non ti è chiara mentre lavori, torna qui.');
table(['Parola', 'Cosa significa, in parole semplici'], [
  ['CCII', 'Codice della crisi d’impresa e dell’insolvenza, D.Lgs. 12 gennaio 2019 n. 14. È la legge che regola tutto. Ha sostituito la vecchia L. 3/2012 (“legge salva-suicidi”) dal 15 luglio 2022.'],
  ['Sovraindebitamento', 'Art. 2 c.1 lett. c) CCII: il debitore non riesce più a pagare regolarmente i debiti con il suo patrimonio e il suo reddito. Riguarda chi NON può fallire (oggi: “liquidazione giudiziale”): consumatori, professionisti, piccoli imprenditori, imprenditori agricoli, start-up innovative.'],
  ['Consumatore', 'Art. 2 c.1 lett. e): persona fisica che ha fatto debiti per scopi estranei alla sua attività di impresa o professionale. Conta la NATURA dei debiti, non il lavoro che fa oggi. Chi ha debiti da fideiussione per la società di famiglia di regola NON è consumatore per quei debiti (tema discusso: vedi cap. 6).'],
  ['Impresa minore', 'Art. 2 c.1 lett. d): impresa che nei 3 esercizi precedenti (o da quando esiste) ha avuto: attivo patrimoniale annuo non superiore a 300.000 €, ricavi lordi annui non superiori a 200.000 €, debiti anche non scaduti non superiori a 500.000 €. Tutti e tre i requisiti insieme.'],
  ['OCC', 'Organismo di composizione della crisi da sovraindebitamento (art. 2 c.1 lett. t). Ente (spesso presso l’Ordine dei commercialisti o la Camera di commercio) che riceve la domanda del debitore e nomina il Gestore. Iscritto nel registro del Ministero della giustizia (DM 202/2014).'],
  ['Gestore della crisi', 'La persona fisica che l’OCC nomina per la singola pratica. Studia il caso, verifica i documenti, scrive la RELAZIONE. Deve essere indipendente dal debitore. In liquidazione controllata di solito diventa il LIQUIDATORE.'],
  ['Referente OCC', 'Chi dirige l’OCC e assegna gli incarichi ai Gestori.'],
  ['Ricorso', 'L’atto con cui si chiede al Tribunale di aprire la procedura. Lo presenta il debitore (con l’OCC e, quando serve, con un avvocato).'],
  ['Relazione particolareggiata', 'Il documento del Gestore che accompagna il ricorso. Il giudice si fida della relazione per decidere: è la parte più delicata.'],
  ['Proposta e piano', 'La proposta è “cosa offro ai creditori” (es. il 30%). Il piano è “come e quando lo pago” (es. 400 € al mese per 72 mesi).'],
  ['Prededuzione', 'Crediti pagati prima di tutti gli altri perché nati per la procedura (es. compenso OCC, art. 6 CCII).'],
  ['Privilegio, pegno, ipoteca', 'Cause di prelazione: danno al creditore il diritto di essere pagato prima dei chirografari. Il privilegio può essere generale (es. dipendenti, Erario, INPS) o speciale su un bene.'],
  ['Chirografario', 'Creditore senza prelazione: viene pagato per ultimo, in proporzione (banche per prestiti personali, finanziarie, fornitori...).'],
  ['Falcidia', 'Riduzione del debito: pagare meno del dovuto (es. il 25% invece del 100%).'],
  ['Classi', 'Gruppi di creditori con posizione giuridica e interessi economici omogenei, a cui si possono fare offerte diverse.'],
  ['Alternativa liquidatoria', 'Quanto prenderebbero i creditori se si vendesse tutto in una liquidazione controllata. È il “termine di paragone” per dire se un piano è conveniente.'],
  ['Merito creditizio', 'Il controllo che banca o finanziaria dovevano fare prima di prestare soldi (art. 124-bis TUB): “questa persona potrà restituirli?”. Se non l’hanno fatto bene, perdono alcune difese (art. 69 c.2 CCII).'],
  ['Meritevolezza', 'Il debitore non deve aver causato i debiti con colpa grave, malafede o frode. È il requisito “morale” delle procedure.'],
  ['Omologa / omologazione', 'La sentenza con cui il Tribunale approva il piano e lo rende obbligatorio per tutti i creditori.'],
  ['Esdebitazione', 'La liberazione dai debiti rimasti non pagati. È il vero obiettivo del debitore: ripartire da zero (“fresh start”).'],
  ['Scala di equivalenza ISEE', 'Coefficienti del DPCM 159/2013 che servono a “pesare” il numero di persone in famiglia: 1 persona = 1; 2 = 1,57; 3 = 2,04; 4 = 2,46; 5 = 2,85; +0,35 per ogni persona in più.'],
  ['Assegno sociale', 'Importo fissato ogni anno dall’INPS. Il CCII lo usa come misura del “minimo per vivere”. Per il 2026 risulta di 546,24 € al mese per 13 mensilità (7.101,12 € annui) [VERIFICARE sulla circolare INPS dell’anno della pratica].'],
  ['PCT / PEC', 'Processo civile telematico: gli atti si depositano solo online, firmati digitalmente. Le comunicazioni con i creditori passano via PEC.'],
], [2300, 6726]);

// =====================================================================
// 2. LE NORME
// =====================================================================
C.push(h1('2. Le norme: dove si trova cosa'));
par('Il sovraindebitamento sta in più punti del CCII. Questa mappa ti dice quale articolo aprire per ogni problema. Tienila sempre a portata di mano.');
table(['Articoli CCII', 'Argomento', 'Quando ti serve'], [
  ['Art. 2', 'Definizioni (sovraindebitamento, consumatore, impresa minore, OCC)', 'Sempre: per dire chi è il debitore e perché può accedere'],
  ['Art. 6', 'Crediti prededucibili', 'Per collocare il compenso OCC nel piano'],
  ['Artt. 27-28', 'Competenza del Tribunale', 'Intestazione del ricorso (consumatore: luogo di residenza)'],
  ['Art. 65', 'Disposizioni generali sulle procedure di composizione', 'Introduzione: chi può accedere, ruolo dell’OCC'],
  ['Art. 66', 'Procedura familiare', 'Quando ci sono coniugi o parenti conviventi con debiti in comune'],
  ['Artt. 67-73', 'Ristrutturazione dei debiti del consumatore', 'Procedura n. 1'],
  ['Artt. 74-83', 'Concordato minore', 'Procedura n. 2'],
  ['Artt. 268-277', 'Liquidazione controllata', 'Procedura n. 3'],
  ['Artt. 278-282', 'Esdebitazione (in generale e dopo la liquidazione)', 'Dopo la liquidazione controllata'],
  ['Art. 283', 'Esdebitazione del sovraindebitato incapiente', 'Procedura n. 4'],
  ['Artt. 356-358', 'Requisiti dei gestori e dei liquidatori', 'Premessa della relazione (requisiti e indipendenza)'],
], [1700, 4000, 3326]);
h2c('Altre norme che userai spesso');
function h2c(t) { C.push(h2(t)); }
list([
  'Art. 124-bis D.Lgs. 385/1993 (TUB): obbligo della banca di verificare il merito creditizio. Si cita nella parte della relazione sul merito creditizio.',
  'Art. 545 c.p.c.: limiti di pignorabilità di stipendi e pensioni (di regola un quinto). Serve per capire quanta parte di reddito andrebbe ai creditori in liquidazione.',
  'Artt. 2740-2741 c.c.: il debitore risponde con tutti i suoi beni; i creditori sono uguali salvo le cause di prelazione.',
  'Artt. 2745 ss. c.c. (privilegi), 2751-bis c.c. (privilegio di professionisti e dipendenti), 2752 c.c. (privilegio dell’Erario), 2753-2754 c.c. (INPS): servono per mettere in ordine i creditori.',
  'DM 202/2014: regolamento degli OCC e tariffe dei compensi (art. 16 e seguenti).',
  'DPCM 159/2013: scala di equivalenza ISEE.',
]);
box('Come si cita una norma', [
  'Forma corretta: “art. 67, comma 4, D.Lgs. 14/2019” oppure, dopo la prima volta, “art. 67 c.4 CCII”.',
  'Lettere: “art. 2, comma 1, lett. c), CCII”.',
  'Mai citare un comma “a memoria”: aprilo e rileggilo. Se una sentenza è citata, indica Tribunale, data e numero (e verifica che esista davvero).',
], AZZURRO);

// =====================================================================
// 3. QUALE PROCEDURA
// =====================================================================
C.push(h1('3. Le 4 procedure e come si sceglie'));
par('Le procedure da sovraindebitamento sono quattro. Si sceglie in base a CHI è il debitore e a COSA può offrire ai creditori.');
table(['Procedura', 'Per chi', 'Idea di fondo', 'Serve il voto dei creditori?'], [
  ['1. Ristrutturazione dei debiti del consumatore (artt. 67-73)', 'Solo consumatore', 'Pago una parte dei debiti con il mio reddito futuro e tengo i miei beni (se il piano lo prevede)', 'NO: decide il giudice'],
  ['2. Concordato minore (artt. 74-83)', 'Professionisti, piccoli imprenditori, agricoli, start-up; anche consumatori con debiti “misti” (es. ex soci)', 'Proseguo l’attività e pago i creditori con i suoi frutti, oppure porto finanza esterna', 'SÌ: maggioranza dei crediti'],
  ['3. Liquidazione controllata (artt. 268-277)', 'Tutti i sovraindebitati', 'Metto a disposizione tutti i beni e una parte del reddito per circa 3 anni; poi esdebitazione', 'NO'],
  ['4. Esdebitazione dell’incapiente (art. 283)', 'Solo persona fisica che non può offrire NULLA', 'Non ho niente da dare: chiedo di essere liberato dai debiti, una volta sola', 'NO (i creditori possono opporsi)'],
], [2300, 2100, 3026, 1600]);

C.push(h2('Schema 1 – Albero per scegliere la procedura'));
par('Parti dall’alto e rispondi alle domande. Prima però verifica che il debitore sia davvero sovraindebitato e non possa fallire (se è un’impresa sopra le soglie dell’impresa minore va alla liquidazione giudiziale, non qui).');
par('Prima domanda di tutte: il debitore può offrire ai creditori qualcosa, adesso o in futuro (reddito sopra il minimo vitale, beni, aiuto di familiari)? Se la risposta è NO, vai direttamente al riquadro rosa in fondo allo schema.');
albero([
  ['È un consumatore (i debiti sono tutti personali/familiari)?', 'RISTRUTTURAZIONE DEI DEBITI DEL CONSUMATORE\nartt. 67-73 CCII\n(se non c’è nessuna utilità neanche futura: art. 283)'],
  ['È professionista o piccolo imprenditore che vuole continuare l’attività? Oppure ha qualcuno che porta finanza esterna che aumenta in modo apprezzabile quanto prendono i creditori?', 'CONCORDATO MINORE\nartt. 74-83 CCII'],
  ['Ha beni e/o reddito da mettere a disposizione ma non può (o non conviene) fare un piano?', 'LIQUIDAZIONE CONTROLLATA\nartt. 268-277 CCII\npoi esdebitazione (art. 282)'],
], 'NESSUNA UTILITÀ, NEMMENO FUTURA\nEsdebitazione dell’incapiente – art. 283 CCII\n(solo persona fisica, solo una volta nella vita)');
box('Attenzione: prima di scegliere controlla gli “ostacoli”', [
  'Esdebitazione già ottenuta nei 5 anni precedenti o già ottenuta 2 volte: blocca RDC e concordato minore (artt. 69 e 77).',
  'Colpa grave, malafede o frode nel creare i debiti: blocca la RDC (art. 69 c.1). Nel concordato minore rilevano gli atti in frode (art. 77). Nell’esdebitazione rileva l’assenza di dolo o colpa grave (art. 283) e le condizioni dell’art. 280.',
  'Procedura precedente revocata o annullata per colpa del debitore: verificare, perché può impedire l’accesso.',
  'Esecuzioni in corso (pignoramenti, aste): non bloccano la domanda, ma vanno indicate subito perché si possono chiedere misure di sospensione.',
], GIALLO);

// =====================================================================
// 4. IL LAVORO PASSO PER PASSO
// =====================================================================
C.push(h1('4. Il lavoro in studio, passo per passo'));
par('Ogni pratica segue lo stesso percorso, qualunque sia la procedura. Questo è lo schema generale; nelle schede delle singole procedure trovi i dettagli.');
flow('Schema 2 – Dal primo incontro al deposito', [
  ['1. Primo colloquio con il debitore', 'Storia dei debiti, famiglia, lavoro, beni.\nConsegna della lista documenti (tabella sotto).'],
  ['2. Domanda all’OCC e nomina del Gestore', 'Il debitore presenta istanza all’OCC.\nL’OCC nomina il Gestore, che accetta, dichiara la propria indipendenza e fa il preventivo del compenso.'],
  ['3. Comunicazione al Fisco entro 7 giorni', 'Il Gestore avvisa Agenzia Entrate-Riscossione e uffici fiscali (anche Comune/Regione).\nEssi hanno 15 giorni per comunicare il debito (artt. 68 c.4, 76 c.4, 269 CCII).'],
  ['4. Raccolta documenti e visure', 'Centrale Rischi Banca d’Italia, CRIF / banche dati private, estratto di ruolo AdER, cassetto fiscale, visure catastali e PRA, casellario e carichi pendenti.'],
  ['5. Analisi dei numeri', 'Tabella dei debiti, reddito, spese di famiglia, beni.\nCalcolo della quota disponibile per i creditori (cap. 5).'],
  ['6. Scelta della procedura', 'Con l’albero del cap. 3. Ne parliamo sempre insieme prima di scrivere.'],
  ['7. Bozza della proposta / del ricorso', 'Usa l’indice della scheda della procedura e il fac-simile in appendice.'],
  ['8. Bozza della relazione del Gestore', 'Verifica INDIPENDENTE di quanto scritto nel ricorso. Il Gestore la rivede, la completa e la firma.'],
  ['9. Controllo finale (cap. 12) e firme', 'Firma digitale del ricorso (debitore/avvocato) e della relazione (Gestore).'],
  ['10. Deposito telematico al Tribunale', 'Con gli allegati numerati e il contributo unificato.', VERDE],
]);

C.push(h2('La lista dei documenti da chiedere al debitore'));
par('Consegna questa lista al primo incontro. Spunta ogni voce quando arriva il documento. I documenti mancanti vanno sollecitati per iscritto (mail), così resta traccia.');
table(['Area', 'Documenti', 'Perché servono'], [
  ['Identità e famiglia', 'Carta d’identità, codice fiscale, certificato di residenza e stato di famiglia, eventuale sentenza di separazione/divorzio', 'Competenza, nucleo familiare, mantenimento ex coniuge e figli'],
  ['Redditi', 'Dichiarazioni dei redditi (730/Unico) e CU degli ultimi 3 anni; ultime 12 buste paga o cedolini pensione; per professionisti e imprese: bilanci o scritture contabili, dichiarazioni IVA e IRAP degli ultimi 3 anni', 'Obbligatori per legge (artt. 67 c.2, 75 c.1, 283 CCII)'],
  ['Debiti', 'Contratti di mutuo, prestiti, cessioni del quinto, carte revolving; ultimi estratti conto dei finanziamenti; cartelle e avvisi; decreti ingiuntivi, atti di pignoramento, precetti', 'Elenco dei creditori con importi e cause di prelazione'],
  ['Banche dati', 'Centrale Rischi Banca d’Italia; CRIF/CTC/Experian; estratto di ruolo AdER; cassetto fiscale; Anagrafe dei rapporti finanziari (se disponibile)', 'Verificare che non manchi nessun debito'],
  ['Patrimonio', 'Visure catastali, atti di provenienza degli immobili, perizia o stima; libretto auto e visura PRA; estratti conto di tutti i conti degli ultimi 12-24 mesi; polizze, titoli, quote societarie, TFR maturato', 'Consistenza e composizione del patrimonio'],
  ['Atti di straordinaria amministrazione', 'Atti di vendita, donazione, costituzione di fondo patrimoniale, trust, cessioni di quote negli ULTIMI 5 ANNI', 'Obbligatorio: il giudice verifica che non si siano “svuotati” i beni'],
  ['Spese di famiglia', 'Contratto di affitto, bollette, spese mediche, scolastiche, assicurazioni, trasporti (con ricevute)', 'Calcolo di quanto serve al mantenimento: valgono solo le spese documentate'],
  ['Storia e cause', 'Certificati medici, lettere di licenziamento, documenti su crisi della società, eventuale certificazione SerD (ludopatia)', 'Cause dell’indebitamento e meritevolezza'],
  ['Penale', 'Certificato del casellario giudiziale e dei carichi pendenti', 'Condizioni ostative (art. 280 per l’esdebitazione)'],
], [1800, 4426, 2800]);

// =====================================================================
// 5. I CONTI
// =====================================================================
C.push(h1('5. I conti che servono in ogni pratica'));
par('Ricorso e relazione si reggono su cinque tabelle. Se queste sono giuste, il resto si scrive facilmente. Lavorale in Excel, con una colonna “documento di riferimento” accanto a ogni numero.');
C.push(h2('Tabella A – I debiti'));
par('Una riga per ogni creditore. Il debito si prende dal documento più recente (estratto conto, precisazione del creditore, estratto di ruolo), confrontandolo con Centrale Rischi e CRIF.');
table(['N.', 'Creditore', 'Titolo del debito', 'Importo (€)', 'Prelazione', 'Fonte'], [
  ['1', 'Banca Alfa', 'Mutuo prima casa', '62.000', 'Ipoteca 1° grado', 'Piano di ammortamento al ../..'],
  ['2', 'Finanziaria Beta', 'Cessione del quinto', '14.500', 'Chirografo', 'Conteggio estintivo'],
  ['3', 'AdER', 'Cartelle IRPEF', '8.200', 'Privilegio art. 2752 c.c. (in parte chirografo: sanzioni)', 'Estratto di ruolo'],
  ['4', 'Banca Gamma', 'Carta revolving', '6.300', 'Chirografo', 'Estratto conto'],
  ['', 'TOTALE', '', '91.000', '', ''],
], [500, 1800, 1900, 1200, 2026, 1600]);
box('Ricorda', [
  'Dividi sempre il debito tributario: imposte e interessi (privilegio) da sanzioni (chirografo). L’estratto di ruolo lo distingue.',
  'Se la Centrale Rischi mostra un debito che il debitore non ti ha detto: chiedi spiegazioni e mettilo comunque in elenco.',
  'Fideiussioni prestate per altri: vanno indicate anche se non ancora escusse.',
], AZZURRO);

C.push(h2('Tabella B – Il reddito della famiglia'));
par('Reddito netto mensile di ogni componente. Attenzione alle mensilità: se c’è tredicesima, il reddito annuo si divide per 12 per avere la media mensile reale. Il reddito del coniuge non va ai creditori, ma si indica sempre perché conta per valutare le spese di famiglia.');
C.push(h2('Tabella C – Le spese di mantenimento'));
par('Elenco delle spese mensili necessarie (affitto, utenze, cibo, trasporti, salute, scuola). Valgono solo quelle documentate o ragionevoli rispetto ai dati ISTAT sulla spesa delle famiglie. Spese “di lusso” (abbonamenti superflui, auto costose) non vanno messe.');

C.push(h2('Tabella D – La quota disponibile per i creditori'));
flow('Schema 3 – Come si arriva alla rata del piano', [
  ['Reddito netto mensile del debitore', '(media su 12 mesi, comprese 13ª/14ª)'],
  ['meno: spese di mantenimento documentate', '(sua parte delle spese di famiglia)'],
  ['= quota disponibile', 'È il massimo che può dare ogni mese', VERDE],
  ['meno: margine di prudenza', 'Per imprevisti (es. 5-10%): un piano troppo “tirato” non è fattibile'],
  ['= RATA DEL PIANO', 'moltiplicata per i mesi di durata = somma totale per i creditori\n(prima si pagano prededuzione e privilegiati, poi i chirografari)', GIALLO],
], 7200);

C.push(h2('Tabella E – Il merito creditizio (per ogni finanziamento)'));
par('Il Gestore deve dire se chi ha prestato i soldi ha valutato bene il merito creditizio (artt. 68 c.3, 76 c.3, 283 CCII). La legge considera idonea, come “minimo per vivere”, la somma non inferiore all’assegno sociale moltiplicato per il parametro della scala ISEE del nucleo familiare. Si ricostruisce la situazione AL MOMENTO di ogni finanziamento:');
table(['Finanziamento', 'Data', 'Rata (€/mese)', 'Reddito netto allora (€/mese)', 'Rate già in corso (€/mese)', 'Reddito residuo dopo tutte le rate', 'Minimo vitale (assegno sociale × scala ISEE)', 'Esito'], [
  ['Prestito Beta', '03/2019', '280', '1.650', '420', '950', '1.050', 'Sotto il minimo: merito non verificato'],
  ['Revolving Gamma', '06/2021', '150', '1.700', '700', '850', '1.080', 'Sotto il minimo'],
], [1200, 800, 900, 1200, 1100, 1300, 1426, 1100]);
nota('Esempio con numeri inventati. Il minimo vitale si calcola con l’assegno sociale dell’anno del finanziamento (assegno annuo / 12 × parametro ISEE).');
box('Perché è importante', [
  'Se il finanziatore non ha valutato correttamente il merito creditizio, nella RDC non può fare opposizione né reclamo contro l’omologa, né sollevare cause di inammissibilità che non derivino da dolo del debitore (art. 69 c.2 CCII).',
  'La giurisprudenza ritiene che la colpa del finanziatore “pesi” anche sulla valutazione della colpa del debitore.',
], VERDE);

C.push(h2('Esempio completo: il signor Rossi (dati inventati)'));
box('Il caso', [
  'Chi è: Mario Rossi, 45 anni, operaio, sposato con un figlio. Moglie con reddito part-time. Vivono in affitto.',
  'Debiti: 85.000 € tutti chirografari (prestiti personali, cessione del quinto, revolving).',
  'Reddito: 1.800 € netti × 13 mensilità = 23.400 €/anno = 1.950 € medi al mese.',
  'Spese documentate (sua parte): 1.450 €/mese.',
  'Beni: auto del 2012, valore di realizzo circa 2.000 €.',
], GRIGIO);
table(['Voce', 'Piano (RDC)', 'Alternativa: liquidazione controllata'], [
  ['Durata', '72 mesi (6 anni)', '36 mesi (circa 3 anni)'],
  ['Quota mensile ai creditori', '450 € (1.950 – 1.450 = 500, meno 10% di prudenza)', 'Quota lasciata dal giudice: ipotesi 500 €/mese di reddito acquisibile → in prudenza 450'],
  ['Totale da reddito', '450 × 72 = 32.400 €', '450 × 36 = 16.200 €'],
  ['Beni', 'L’auto resta al debitore (serve per lavoro)', '+ 2.000 € (vendita auto)'],
  ['Costi della procedura', '– 3.000 € (compenso OCC accantonato)', '– 3.500 € (compenso liquidatore e spese di vendita)'],
  ['Netto per i creditori', '29.400 € → 34,6%', '14.700 € → 17,3%'],
  ['Conclusione', 'Il piano dà ai creditori PIÙ della liquidazione: è conveniente', ''],
], [2200, 3400, 3426]);
nota('Il confronto con la liquidazione è richiesto espressamente nel concordato minore (art. 76 c.2 lett. d) e in ogni caso il giudice lo usa quando un creditore contesta la convenienza (artt. 70 e 80). Farlo sempre è buona pratica.');

// =====================================================================
// 6. RDC
// =====================================================================
C.push(h1('6. Procedura 1 – Ristrutturazione dei debiti del consumatore (artt. 67-73 CCII)'));
C.push(h2('In una frase'));
par('Il consumatore propone di pagare una parte dei debiti, con tempi e modi indicati in un piano; il giudice, se il piano è ammissibile e fattibile, lo omologa anche contro la volontà dei creditori. Non si vota.');
C.push(h2('Requisiti'));
list([
  'Soggettivo: essere CONSUMATORE (art. 2 c.1 lett. e) in stato di sovraindebitamento (art. 2 c.1 lett. c). Possibile la procedura familiare (art. 66).',
  'Nessuna condizione ostativa (art. 69 c.1): non esdebitato nei 5 anni precedenti la domanda, non esdebitato già 2 volte, sovraindebitamento NON determinato con colpa grave, malafede o frode.',
  'Domanda tramite OCC (art. 68 c.1): la domanda arriva al giudice “tramite” l’OCC del circondario del Tribunale competente. Per questa procedura l’assistenza di un avvocato non è obbligatoria (orientamento prevalente; verificare la prassi del Tribunale).',
]);
C.push(h2('Cosa si può proporre (art. 67)'));
list([
  'Contenuto libero: il piano indica in modo specifico tempi e modalità per superare la crisi e può prevedere il pagamento anche parziale e differenziato dei crediti, in qualsiasi forma (art. 67 c.1).',
  'Privilegiati: si possono pagare non integralmente, purché in misura non inferiore a quanto otterrebbero dalla liquidazione dei beni su cui hanno la prelazione, avuto riguardo al valore di mercato attestato dall’OCC (art. 67 c.3).',
  'Cessione del quinto e delegazione di pagamento: si possono falcidiare come gli altri crediti (art. 67 c.4). Con l’apertura il giudice può sospendere le trattenute.',
  'Mutuo della casa di abitazione: si può continuare a pagare il mutuo garantito da ipoteca sulla casa principale alle scadenze originarie se il debitore, alla data del deposito, ha adempiuto le proprie obbligazioni o se il giudice lo autorizza a pagare il debito scaduto per capitale e interessi (art. 67 c.5). È il modo per salvare la casa.',
  'Durata: la legge non fissa un massimo; i Tribunali chiedono una durata “ragionevole” (spesso 5-7 anni). Piani più lunghi vanno motivati bene.',
  'La proposta deve dare comunque qualcosa ai creditori: un piano solo “dilatorio” (che rinvia senza pagare) non è ammissibile.',
]);
C.push(h2('Documenti da allegare (art. 67 c.2)'));
list([
  'a) Elenco di tutti i creditori, con le somme dovute e le cause di prelazione.',
  'b) Elenco della consistenza e composizione del patrimonio.',
  'c) Elenco degli atti di straordinaria amministrazione compiuti negli ultimi 5 anni.',
  'd) Dichiarazioni dei redditi degli ultimi 3 anni.',
  'e) Indicazione di stipendi, pensioni, salari e altre entrate del debitore e del nucleo familiare, con quanto occorre al mantenimento della famiglia.',
]);
C.push(h2('La relazione del Gestore (art. 68 c.2-4)'));
par('La relazione particolareggiata deve contenere OBBLIGATORIAMENTE:');
list([
  'a) le cause dell’indebitamento e la diligenza impiegata dal debitore nell’assumere le obbligazioni;',
  'b) le ragioni dell’incapacità del debitore di adempiere le obbligazioni assunte;',
  'c) la valutazione sulla completezza e attendibilità della documentazione depositata;',
  'd) l’indicazione presunta dei costi della procedura;',
  'c.3: se il finanziatore ha tenuto conto del merito creditizio del debitore (Tabella E del cap. 5);',
  'c.4: che l’OCC, entro 7 giorni dall’incarico, ha avvisato agente della riscossione e uffici fiscali (anche enti locali), che hanno 15 giorni per comunicare il debito tributario accertato e gli accertamenti pendenti.',
]);
nota('Anche se non elencata alla lettera nell’art. 68, conviene inserire sempre la valutazione di fattibilità e il confronto con la liquidazione: il giudice li verifica (art. 70) e servono se un creditore contesta la convenienza.');

flow('Schema 4 – L’iter della ristrutturazione dei debiti del consumatore', [
  ['Deposito del ricorso + relazione OCC', 'Tribunale del luogo di residenza (art. 27 c.2)'],
  ['Decreto del giudice (art. 70 c.1)', 'Se ammissibile: pubblicazione sul sito e comunicazione ai creditori.\nSu istanza del debitore: misure protettive (sospensione esecuzioni, divieto di azioni esecutive e cautelari, divieto di atti di straordinaria amministrazione) – art. 70 c.4'],
  ['Entro 30 giorni: l’OCC comunica a tutti i creditori', 'Proposta, piano e decreto, via PEC (art. 70 c.1)'],
  ['20 giorni: osservazioni dei creditori', 'Inviate alla PEC dell’OCC (art. 70 c.3)'],
  ['10 giorni dopo: l’OCC riferisce al giudice', 'Sentito il debitore, propone le eventuali modifiche al piano (art. 70 c.6)'],
  ['Sentenza di omologa (art. 70 c.7)', 'Il giudice verifica ammissibilità giuridica e fattibilità economica e risolve le contestazioni.\nSe un creditore contesta la convenienza: omologa se il creditore riceve almeno quanto avrebbe dalla liquidazione.', VERDE],
  ['Esecuzione del piano (art. 71)', 'Il debitore paga; l’OCC vigila e riferisce al giudice periodicamente (di regola ogni 6 mesi).'],
  ['Relazione finale e compenso OCC (art. 71)', 'Il giudice, se il piano è stato eseguito, liquida il compenso dell’OCC.\nEsdebitazione per i debiti residui.', VERDE],
], 8200);
box('Se le cose vanno male', [
  'Mancata omologa: il giudice dichiara inammissibile/rigetta con decreto; in caso di frode, su istanza, può aprire la liquidazione controllata (art. 70).',
  'Revoca dell’omologa (art. 72): se il debitore ha aumentato o diminuito dolosamente il passivo, sottratto o nascosto attivo, non esegue il piano, ecc.',
  'Conversione in liquidazione controllata (art. 73): in caso di revoca o inadempimento, su istanza del debitore, dei creditori o del PM.',
], ROSA);

C.push(h2('Indice tipo del RICORSO – RDC'));
table(['Parte', 'Cosa scrivere', 'Norma'], [
  ['Intestazione', 'TRIBUNALE DI ... – Sezione/Ufficio procedure concorsuali. “Ricorso per l’omologazione del piano di ristrutturazione dei debiti del consumatore”', 'artt. 27, 67 CCII'],
  ['Ricorrente', 'Nome, data e luogo di nascita, C.F., residenza, PEC per le comunicazioni. “Con l’ausilio dell’OCC ..., Gestore dott. ...”', 'art. 68 c.1'],
  ['1. Premessa', 'Chi è il debitore, famiglia, lavoro. Due-tre paragrafi.', ''],
  ['2. Requisiti di accesso', 'Competenza; qualità di consumatore (perché i debiti sono estranei ad attività d’impresa); stato di sovraindebitamento (numeri: debiti vs reddito e patrimonio); assenza di condizioni ostative.', 'artt. 2, 27, 69 c.1'],
  ['3. Storia e cause dell’indebitamento', 'Racconto cronologico: quando e perché è nato ogni debito, cosa è successo (perdita lavoro, malattia, separazione...). Collegare ogni evento a un documento.', 'art. 68 c.2 lett. a-b'],
  ['4. Situazione debitoria', 'Tabella A completa, totale per categoria (prededuzione, privilegiati, chirografari).', 'art. 67 c.2 lett. a'],
  ['5. Patrimonio', 'Immobili, mobili registrati, conti, altri beni, con valore.', 'art. 67 c.2 lett. b'],
  ['6. Atti di straordinaria amministrazione', 'Elenco degli ultimi 5 anni, oppure dichiarazione che non ce ne sono.', 'art. 67 c.2 lett. c'],
  ['7. Redditi e spese della famiglia', 'Tabelle B e C. Dichiarazioni dei redditi dei 3 anni.', 'art. 67 c.2 lett. d-e'],
  ['8. Proposta e piano', 'Durata, rata, decorrenza, ordine di pagamento (prededuzione → privilegiati → chirografari), percentuali, trattamento mutuo casa e cessione del quinto, chi paga e come (conto dedicato).', 'art. 67 c.1, 3, 4, 5'],
  ['9. Fattibilità e convenienza', 'Perché il piano è sostenibile; confronto con la liquidazione (esempio cap. 5).', 'art. 70'],
  ['10. Istanze', 'Misure protettive: sospensione esecuzioni, divieto di azioni, sospensione delle trattenute del quinto. Autorizzazione a pagare il mutuo.', 'art. 70 c.4; art. 67 c.5'],
  ['Conclusioni', '“Voglia l’Ill.mo Tribunale: dichiarare ammissibile ..., disporre ..., omologare ...”', 'art. 70'],
  ['Allegati', 'Elenco numerato (doc. 1, 2, 3 ...). Relazione OCC sempre come primo allegato.', ''],
], [2000, 5326, 1700]);

C.push(h2('Indice tipo della RELAZIONE DEL GESTORE – RDC'));
table(['Capitolo', 'Cosa scrivere', 'Norma'], [
  ['1. Incarico', 'Data della domanda all’OCC, nomina, accettazione, dichiarazione di indipendenza e requisiti del Gestore.', 'artt. 356-358 CCII; DM 202/2014'],
  ['2. Attività svolta', 'Incontri con il debitore, documenti ricevuti, banche dati consultate, comunicazioni a fisco e AdER con le risposte ricevute.', 'art. 68 c.4'],
  ['3. Il debitore e la famiglia', 'Dati anagrafici, nucleo familiare, lavoro, residenza. Qualità di consumatore.', 'art. 2 c.1 lett. e'],
  ['4. Situazione debitoria verificata', 'Tabella A confrontata con Centrale Rischi, CRIF, estratto di ruolo: segnalare le differenze.', 'art. 67 c.2 lett. a'],
  ['5. Cause dell’indebitamento e diligenza', 'Giudizio motivato: le cause sono esterne o colpevoli? Il debitore è stato diligente? (assenza di colpa grave, malafede, frode).', 'art. 68 c.2 lett. a; art. 69 c.1'],
  ['6. Ragioni dell’incapacità di adempiere', 'Perché oggi non riesce a pagare: confronto rate dovute / reddito disponibile.', 'art. 68 c.2 lett. b'],
  ['7. Merito creditizio', 'Tabella E per ogni finanziamento e conclusione.', 'art. 68 c.3; art. 124-bis TUB'],
  ['8. Completezza e attendibilità della documentazione', 'Giudizio esplicito: “la documentazione è completa e attendibile” (o cosa manca e perché).', 'art. 68 c.2 lett. c'],
  ['9. Patrimonio, redditi, spese', 'Verifica delle tabelle B, C e del patrimonio; valori di mercato dei beni.', 'art. 67 c.3'],
  ['10. Il piano: fattibilità', 'Sostenibilità della rata, margine di sicurezza, rischi (età, salute, stabilità del lavoro).', 'art. 70 c.7'],
  ['11. Confronto con la liquidazione', 'Tabella come nell’esempio del cap. 5.', 'art. 70'],
  ['12. Costi presunti della procedura', 'Compenso OCC (secondo DM 202/2014 e preventivo), spese di pubblicità e trascrizione.', 'art. 68 c.2 lett. d'],
  ['13. Conclusioni', 'Sintesi dei giudizi e attestazioni. Luogo, data, firma digitale del Gestore.', ''],
], [2300, 4926, 1800]);

// =====================================================================
// 7. CONCORDATO MINORE
// =====================================================================
C.push(h1('7. Procedura 2 – Concordato minore (artt. 74-83 CCII)'));
C.push(h2('In una frase'));
par('Il debitore NON consumatore (professionista, piccolo imprenditore, agricoltore, start-up, ma anche persona con debiti “misti”) propone ai creditori un accordo per pagare in parte i debiti; i creditori VOTANO e, se si raggiunge la maggioranza, il giudice omologa.');
C.push(h2('Requisiti'));
list([
  'Soggettivo: debitore sovraindebitato diverso dal consumatore (art. 74 c.1): impresa minore, imprenditore agricolo, start-up innovativa, professionista, e chi ha debiti anche di natura imprenditoriale.',
  'Due strade: (1) proposta che consente di PROSEGUIRE l’attività imprenditoriale o professionale (art. 74 c.1); (2) fuori da questo caso, solo se c’è apporto di RISORSE ESTERNE che aumentano in misura apprezzabile la soddisfazione dei creditori (art. 74 c.2). Senza continuità e senza finanza esterna il concordato minore non si può fare: si va in liquidazione controllata.',
  'Inammissibilità (art. 77): mancano i documenti degli artt. 75-76; il debitore supera i limiti dell’impresa minore; è stato esdebitato nei 5 anni precedenti o già 2 volte; risultano atti diretti a frodare i creditori.',
  'Avvocato: per il concordato minore di regola serve il difensore (art. 9 CCII) [VERIFICARE la prassi del Tribunale].',
  'Rinvio: per quanto non previsto si applicano le norme sul concordato preventivo, in quanto compatibili (art. 74 c.4).',
]);
C.push(h2('Cosa si può proporre'));
list([
  'Contenuto: la proposta indica in modo specifico tempi e modalità per superare la crisi, può prevedere il pagamento anche parziale dei crediti in qualsiasi forma e la suddivisione dei creditori in CLASSI (art. 74 c.3).',
  'Privilegiati: pagamento non integrale possibile, purché non inferiore a quanto realizzabile in caso di liquidazione sul bene oggetto della prelazione, come attestato dall’OCC (art. 75 c.2).',
  'Mutuo su beni strumentali: in caso di continuità si può continuare a pagare alle scadenze il mutuo garantito da beni strumentali all’attività, alle condizioni dell’art. 75 c.3.',
  'Finanza esterna: impegno scritto del terzo (familiare, socio), con prova che ha i soldi e preferibilmente con rinuncia al regresso.',
]);
C.push(h2('Documenti da allegare (art. 75 c.1)'));
list([
  'a) Piano con bilanci, scritture contabili e fiscali obbligatorie, dichiarazioni dei redditi, IRAP e IVA dei 3 anni precedenti (o dall’inizio dell’attività).',
  'b) Relazione aggiornata sulla situazione economica, patrimoniale e finanziaria.',
  'c) Elenco di tutti i creditori, con crediti e cause di prelazione.',
  'd) Atti di straordinaria amministrazione degli ultimi 5 anni.',
  'e) Documentazione di stipendi, pensioni, salari e altre entrate del debitore e della famiglia, con quanto occorre al mantenimento.',
]);
C.push(h2('La relazione del Gestore (art. 76 c.2-4)'));
par('È più ampia di quella della RDC. Deve contenere:');
list([
  'a) cause dell’indebitamento e diligenza del debitore nell’assumere le obbligazioni;',
  'b) ragioni dell’incapacità di adempiere;',
  'c) eventuali atti del debitore impugnati dai creditori;',
  'd) valutazione sulla completezza e attendibilità della documentazione E sulla CONVENIENZA del piano rispetto all’alternativa liquidatoria;',
  'e) costi presumibili della procedura;',
  'f) percentuale, modalità e tempi di soddisfacimento dei creditori;',
  'g) criteri adottati nella formazione delle classi, se previste;',
  'inoltre: merito creditizio dei finanziatori (art. 76 c.3) e comunicazione entro 7 giorni a fisco e AdER (art. 76 c.4);',
  'se si chiede l’omologa senza il voto favorevole di Fisco/INPS (cram down, vedi sotto): specifica valutazione della convenienza per l’amministrazione finanziaria.',
]);
flow('Schema 5 – L’iter del concordato minore', [
  ['Deposito del ricorso + relazione OCC', 'Con avvocato, al Tribunale competente'],
  ['Decreto di apertura (art. 78)', 'Se la domanda è ammissibile: decreto non reclamabile.\nPubblicazione e comunicazione ai creditori a cura dell’OCC.\nTermine (max 30 giorni) per far pervenire all’OCC, via PEC, adesione o dissenso e contestazioni.\nSu istanza: divieto di azioni esecutive e cautelari fino all’omologa definitiva.'],
  ['Voto dei creditori (art. 79)', 'Approvazione: maggioranza dei crediti ammessi al voto.\nSe un solo creditore ha più della metà dei crediti: serve anche la maggioranza per teste dei votanti.\nCon le classi: maggioranza anche nel maggior numero di classi.\nIl SILENZIO vale come consenso.'],
  ['Relazione dell’OCC sul voto', 'L’OCC riferisce al giudice l’esito del voto e le contestazioni.'],
  ['Sentenza di omologa (art. 80)', 'Il giudice verifica ammissibilità, fattibilità e maggioranze.\nSe un creditore contesta la convenienza: omologa se riceve almeno quanto dalla liquidazione.\nCram down fiscale: omologa anche senza voto di Fisco/INPS se il loro voto era determinante e la proposta è per loro conveniente.', VERDE],
  ['Esecuzione (art. 81)', 'Il debitore esegue; l’OCC vigila, riferisce periodicamente al giudice e, a fine esecuzione, deposita la relazione finale.'],
  ['Fine: compenso OCC ed esdebitazione', '', VERDE],
], 8400);
box('Chi vota e chi no (art. 79)', [
  'Non votano: i privilegiati pagati per intero (salvo rinuncia alla prelazione); il coniuge, i parenti e affini fino al 4° grado e i cessionari dei loro crediti da meno di un anno.',
  'Silenzio-assenso: chi non risponde nel termine si considera favorevole alla proposta come gli è stata trasmessa. Per questo la comunicazione PEC dell’OCC deve essere perfetta e documentata (ricevute di consegna).',
  'Società: il concordato minore della società produce effetti anche per i soci illimitatamente responsabili.',
], AZZURRO);
box('Cram down fiscale – attenzione', [
  'Regola: il giudice omologa anche senza l’adesione dell’amministrazione finanziaria o degli enti previdenziali quando la loro adesione è determinante per la maggioranza e, anche in base alla specifica relazione dell’OCC, la proposta è per loro conveniente rispetto alla liquidazione (art. 80).',
  'In pratica: nella relazione serve un paragrafo apposito con il confronto “quanto prende il Fisco con il piano / quanto prenderebbe con la liquidazione”.',
  '[VERIFICARE]: dopo il correttivo D.Lgs. 136/2024 controllare il testo vigente dell’art. 80 e la giurisprudenza più recente; alcuni Tribunali sono restrittivi.',
], GIALLO);

C.push(h2('Indice tipo del RICORSO – Concordato minore'));
table(['Parte', 'Cosa scrivere', 'Norma'], [
  ['Intestazione', '“Ricorso per l’apertura della procedura di concordato minore” – Tribunale di ...', 'artt. 27, 74 CCII'],
  ['Ricorrente e difensore', 'Dati del debitore (P.IVA se c’è), avvocato con procura, OCC e Gestore.', 'art. 9; art. 76 c.1'],
  ['1. L’attività e il debitore', 'Che attività svolge, da quando, con che risultati. Requisiti di impresa minore (tabella con attivo, ricavi, debiti dei 3 anni).', 'art. 2 c.1 lett. d; art. 74'],
  ['2. Requisiti e assenza di cause di inammissibilità', 'Sovraindebitamento; nessuna esdebitazione nei 5 anni; nessun atto in frode.', 'art. 77'],
  ['3. Cause della crisi', 'Racconto documentato.', 'art. 76 c.2 lett. a'],
  ['4. Situazione economica, patrimoniale e finanziaria', 'Relazione aggiornata: bilanci/dati contabili, situazione debitoria (Tabella A), patrimonio.', 'art. 75 c.1 lett. b-c'],
  ['5. Atti degli ultimi 5 anni', 'Elenco o dichiarazione negativa.', 'art. 75 c.1 lett. d'],
  ['6. La proposta', 'Continuità o finanza esterna; classi (con criteri); percentuali e tempi per ogni classe; trattamento dei privilegiati; fonti dei pagamenti.', 'artt. 74, 75 c.2-3'],
  ['7. Il piano economico', 'Previsioni di incassi e spese anno per anno (business plan semplice per la continuità).', 'art. 75 c.1 lett. a'],
  ['8. Convenienza', 'Confronto con la liquidazione controllata, anche per Fisco e INPS se serve il cram down.', 'artt. 76, 80'],
  ['9. Istanze', 'Apertura; misure protettive; termine per le adesioni.', 'art. 78'],
  ['Conclusioni e allegati', 'Come nella RDC.', ''],
], [2200, 5126, 1700]);
C.push(h2('Indice tipo della RELAZIONE DEL GESTORE – Concordato minore'));
par('Stessa struttura della relazione RDC (cap. 6), con questi capitoli in più:');
list([
  'Requisiti di impresa minore verificati sui bilanci/dichiarazioni dei 3 anni.',
  'Atti impugnati dai creditori (art. 76 c.2 lett. c).',
  'Analisi della continuità: il piano economico è realistico? Su quali dati storici si basa?',
  'Finanza esterna: chi la dà, con quali mezzi (estratti conto, redditi), se è irrevocabile e se aumenta “in misura apprezzabile” quanto prendono i creditori.',
  'Classi: criteri di formazione (art. 76 c.2 lett. g).',
  'Percentuali, modalità e tempi per ogni classe (art. 76 c.2 lett. f).',
  'Convenienza rispetto alla liquidazione (art. 76 c.2 lett. d) e, se serve, paragrafo specifico per il cram down (art. 80).',
]);

// =====================================================================
// 8. LIQUIDAZIONE CONTROLLATA
// =====================================================================
C.push(h1('8. Procedura 3 – Liquidazione controllata (artt. 268-277 CCII)'));
C.push(h2('In una frase'));
par('Il debitore mette a disposizione dei creditori tutto il suo patrimonio (salvo i beni esclusi per legge) e la parte di reddito che supera il necessario per vivere; un liquidatore vende e distribuisce; dopo la chiusura o comunque dopo 3 anni dall’apertura arriva l’esdebitazione.');
C.push(h2('Chi la può chiedere (art. 268)'));
list([
  'Il debitore sovraindebitato (consumatore o no), con ricorso al Tribunale competente (art. 268 c.1).',
  'Un creditore, anche se ci sono esecuzioni in corso, e il PM se l’insolvenza riguarda un imprenditore (art. 268 c.2). Se il debitore è persona fisica e i debiti scaduti e non pagati sono inferiori a 50.000 €, su domanda del creditore la procedura non si apre.',
  'Su domanda di un creditore contro una persona fisica, la liquidazione non si apre se l’OCC, su richiesta del debitore, attesta che non è possibile acquisire attivo da distribuire, neppure con azioni giudiziarie (art. 268 c.3) [VERIFICARE testo vigente dopo il correttivo].',
]);
C.push(h2('Cosa resta fuori dalla liquidazione (art. 268 c.4)'));
list([
  'a) i crediti e i beni impignorabili (art. 545 c.p.c.);',
  'b) i crediti alimentari e di mantenimento, gli stipendi, le pensioni, i salari e ciò che il debitore guadagna con la sua attività, nei limiti di quanto occorre al mantenimento suo e della famiglia, fissati dal GIUDICE;',
  'c) i frutti dell’usufrutto legale sui beni dei figli, i beni del fondo patrimoniale e i loro frutti (salvo art. 170 c.c.);',
  'd) le cose che la legge dichiara non pignorabili.',
]);
C.push(h2('Il ricorso del debitore e la relazione (art. 269)'));
list([
  'Presentazione: il ricorso è presentato dal debitore con l’assistenza dell’OCC (art. 269 c.1). Sulla necessità dell’avvocato le prassi non sono uniformi [VERIFICARE la prassi del Tribunale].',
  'Relazione OCC (art. 269 c.2): deve esporre una valutazione sulla completezza e attendibilità della documentazione depositata e illustrare la situazione economica, patrimoniale e finanziaria del debitore.',
  'Comunicazione al fisco: anche qui l’OCC avvisa entro 7 giorni agente della riscossione e uffici fiscali.',
  'Documenti: in pratica gli stessi delle altre procedure (elenco creditori, beni, redditi e dichiarazioni dei 3 anni, atti dei 5 anni, spese familiari). Per le imprese: bilanci e scritture contabili.',
  'Buona pratica: nella relazione inserire anche cause dell’indebitamento e diligenza del debitore, perché servono poi per l’esdebitazione (art. 280) e il giudice le guarda già all’apertura.',
  'Limite di reddito: proporre al giudice, con calcolo motivato, la quota di reddito che resta al debitore per vivere (spese documentate) e quella che va ai creditori.',
]);
flow('Schema 6 – L’iter della liquidazione controllata', [
  ['Deposito del ricorso + relazione OCC', ''],
  ['Sentenza di apertura (art. 270)', 'Il Tribunale nomina il giudice delegato e il liquidatore (di regola il Gestore OCC che ha seguito il debitore).\nOrdina il deposito entro 7 giorni di bilanci, scritture ed elenco creditori (se mancano).\nAssegna ai creditori e ai terzi un termine (max 60 giorni) per inviare via PEC le domande di ammissione al passivo o di restituzione/rivendicazione.\nOrdina la consegna dei beni (il giudice può autorizzare il debitore a usarne alcuni).\nPubblicazione e trascrizione.'],
  ['Il liquidatore aggiorna l’elenco dei creditori', 'E comunica la sentenza (art. 272)'],
  ['Inventario e programma di liquidazione (art. 272)', 'Entro 90 giorni dall’apertura: inventario dei beni e programma su tempi e modi delle vendite, approvato dal giudice delegato.'],
  ['Stato passivo (art. 273)', 'Il liquidatore prepara il progetto, lo comunica; osservazioni dei creditori entro 15 giorni; se restano contestazioni decide il giudice.'],
  ['Vendite, incasso della quota di reddito, azioni (artt. 274-275)', 'Vendite con procedure competitive; eventuali azioni revocatorie o di recupero.\nIl liquidatore deposita rapporti periodici al giudice (di regola semestrali).'],
  ['Riparto e chiusura (artt. 275-276)', 'Distribuzione ai creditori secondo l’ordine delle prelazioni.', VERDE],
  ['Esdebitazione (artt. 278-282)', 'Opera di diritto con la chiusura o comunque decorsi 3 anni dall’apertura; dichiarata con decreto motivato del Tribunale, verificate le condizioni dell’art. 280. Reclamabile dai creditori.', VERDE],
], 8400);
box('Le condizioni per l’esdebitazione (art. 280) – da controllare GIÀ prima del deposito', [
  'Nessuna condanna definitiva per bancarotta fraudolenta o reati contro economia, commercio, patrimonio (salvo riabilitazione).',
  'Non aver distratto l’attivo o esposto passività inesistenti, né cagionato o aggravato il dissesto rendendo gravemente difficile la ricostruzione del patrimonio.',
  'Non aver ostacolato o rallentato la procedura e aver fornito agli organi le informazioni utili.',
  'Nessuna esdebitazione nei 5 anni precedenti; non più di 2 in tutto.',
  'Per questo serve il casellario giudiziale nel fascicolo fin dall’inizio.',
], GIALLO);

C.push(h2('Indice tipo del RICORSO – Liquidazione controllata (domanda del debitore)'));
table(['Parte', 'Cosa scrivere', 'Norma'], [
  ['Intestazione', '“Ricorso per l’apertura della liquidazione controllata del sovraindebitato” – Tribunale di ...', 'artt. 27, 268 CCII'],
  ['Ricorrente', 'Dati, OCC e Gestore, eventuale difensore.', 'art. 269 c.1'],
  ['1. Il debitore', 'Chi è, famiglia, lavoro.', ''],
  ['2. Stato di sovraindebitamento', 'Debiti vs patrimonio e reddito; perché non c’è altra soluzione (es. non c’è margine per un piano).', 'artt. 2, 268'],
  ['3. Cause', 'Racconto documentato (serve anche per l’esdebitazione).', 'art. 280'],
  ['4. Debiti', 'Tabella A.', ''],
  ['5. Patrimonio', 'Tutti i beni, compresi quelli che si chiede di poter continuare a usare, con il valore di realizzo.', 'art. 270'],
  ['6. Redditi e spese', 'Tabelle B e C; proposta motivata della somma da lasciare al debitore per il mantenimento.', 'art. 268 c.4 lett. b'],
  ['7. Atti dei 5 anni', 'Elenco o dichiarazione negativa.', ''],
  ['8. Istanze', 'Aprire la liquidazione; nominare liquidatore il Gestore; fissare il limite di reddito; autorizzare l’uso dell’auto o della casa fino alla vendita, se motivato.', 'art. 270'],
  ['Conclusioni e allegati', '', ''],
], [2200, 5126, 1700]);
C.push(h2('Indice tipo della RELAZIONE DEL GESTORE – Liquidazione controllata'));
list([
  'Incarico, indipendenza, attività svolta, comunicazioni a fisco/AdER.',
  'Il debitore e la famiglia.',
  'Situazione economica, patrimoniale e finanziaria (art. 269 c.2): debiti verificati, beni con valore di realizzo, redditi.',
  'Cause dell’indebitamento e comportamento del debitore (utile per l’art. 280).',
  'Valutazione di completezza e attendibilità della documentazione (art. 269 c.2): giudizio esplicito.',
  'Attivo prevedibilmente acquisibile: beni + quota di reddito per 3 anni + eventuali azioni (revocatorie, recupero crediti).',
  'Proposta di limite di reddito (spese documentate).',
  'Costi stimati della procedura.',
  'Conclusioni.',
]);

// =====================================================================
// 9. ESDEBITAZIONE INCAPIENTE
// =====================================================================
C.push(h1('9. Procedura 4 – Esdebitazione del sovraindebitato incapiente (art. 283 CCII)'));
C.push(h2('In una frase'));
par('La persona fisica meritevole che non può dare ai creditori NESSUNA utilità, diretta o indiretta, nemmeno in futuro, chiede di essere liberata dai debiti. Si può fare UNA SOLA VOLTA nella vita.');
C.push(h2('Requisiti'));
list([
  'Persona fisica (consumatore o no) in stato di sovraindebitamento.',
  'Incapienza: nessuna utilità per i creditori, nemmeno in prospettiva futura. I finanziamenti, in qualsiasi forma, non sono considerati utilità.',
  'Come si misura (art. 283 c.2): il reddito annuo, dedotte le spese di produzione del reddito e quanto serve al mantenimento della famiglia, va confrontato con la soglia pari all’assegno sociale aumentato della metà, moltiplicato per il parametro della scala ISEE del nucleo [VERIFICARE il testo vigente del comma 2 dopo il correttivo 2024].',
  'Meritevolezza: il giudice verifica l’assenza di atti in frode e la mancanza di dolo o colpa grave nella formazione dell’indebitamento.',
  'Una sola volta: non si può ripetere.',
]);
box('Esempio di calcolo dell’incapienza (dati inventati, assegno sociale 2026)', [
  'Caso: pensionata che vive sola, pensione 750 € × 13 = 9.750 € annui.',
  'Soglia: assegno sociale annuo 7.101,12 € × 1,5 = 10.651,68 € × parametro ISEE 1 (una persona) = 10.651,68 €.',
  'Confronto: 9.750 € < 10.651,68 € → reddito sotto la soglia: nessuna utilità da offrire.',
  'Beni: nessun immobile, nessun bene di valore. → Può chiedere l’esdebitazione dell’incapiente.',
  'Nucleo di 3 persone: la soglia diventa 10.651,68 × 2,04 = 21.729,43 € (si confronta con il reddito di tutta la famiglia, secondo l’orientamento prevalente [VERIFICARE]).',
], VERDE);
C.push(h2('Documenti da allegare (art. 283)'));
list([
  'Elenco di tutti i creditori, con le somme dovute.',
  'Elenco degli atti di straordinaria amministrazione degli ultimi 5 anni.',
  'Copia delle dichiarazioni dei redditi degli ultimi 3 anni.',
  'Indicazione di stipendi, pensioni, salari e di tutte le altre entrate del debitore e del nucleo familiare.',
]);
C.push(h2('La relazione del Gestore (art. 283)'));
list([
  'Cause dell’indebitamento e diligenza del debitore nell’assumere le obbligazioni.',
  'Ragioni dell’incapacità di adempiere.',
  'Eventuali atti del debitore impugnati dai creditori.',
  'Valutazione sulla completezza e attendibilità della documentazione.',
  'Merito creditizio dei finanziatori (come nelle altre procedure).',
  'In pratica, anche: calcolo dell’incapienza con la soglia del comma 2 e attestazione che non ci sono beni o utilità future prevedibili.',
]);
flow('Schema 7 – L’iter dell’esdebitazione dell’incapiente', [
  ['Domanda tramite OCC + relazione', 'Al giudice competente'],
  ['Decreto del giudice', 'Assunte le informazioni e valutata la meritevolezza, concede l’esdebitazione.\nIndica modalità e termine per la dichiarazione annuale delle sopravvenienze, a pena di revoca.', VERDE],
  ['Comunicazione ai creditori', 'Possono fare opposizione entro 30 giorni.'],
  ['Controllo negli anni successivi', 'Il debitore deposita ogni anno la dichiarazione sulle sopravvenienze; l’OCC vigila e verifica.\nSe arrivano utilità rilevanti entro il termine di legge, il debitore deve pagare i creditori.\n[VERIFICARE: termine di 3 anni dopo il correttivo 2024 (prima era di 4) e soglia del 10%]'],
  ['Fine del periodo: liberazione definitiva', '', VERDE],
], 8000);
box('Da sapere', [
  'Compenso OCC: per questa procedura è ridotto (la legge lo riduce della metà) [VERIFICARE].',
  'Errore tipico: chiedere l’esdebitazione dell’incapiente per chi ha una casa di proprietà o un reddito sopra soglia. In quel caso la strada è la liquidazione controllata o la RDC.',
], GIALLO);
C.push(h2('Indice tipo del RICORSO e della RELAZIONE – Incapiente'));
table(['Ricorso', 'Relazione del Gestore'], [
  ['Dati del debitore e dell’OCC', 'Incarico, indipendenza, attività, comunicazioni fiscali'],
  ['Stato di sovraindebitamento', 'Debiti verificati'],
  ['Storia e cause dei debiti', 'Cause, diligenza, ragioni dell’incapacità'],
  ['Redditi e famiglia; calcolo dell’incapienza', 'Verifica del calcolo; assenza di beni; prospettive future (età, salute, lavoro)'],
  ['Meritevolezza (assenza di frode, dolo, colpa grave)', 'Giudizio sulla meritevolezza; merito creditizio'],
  ['Elenco creditori, atti dei 5 anni, dichiarazioni dei 3 anni', 'Completezza e attendibilità della documentazione'],
  ['Richiesta di esdebitazione', 'Conclusioni'],
], [4513, 4513]);

// =====================================================================
// 10. FAMILIARE
// =====================================================================
C.push(h1('10. La procedura familiare (art. 66 CCII)'));
list([
  'Quando: membri della stessa famiglia conviventi, oppure con sovraindebitamento di origine comune (es. coniugi con mutuo cointestato, genitori fideiussori del figlio).',
  'Famiglia: oltre al coniuge, i parenti entro il 4° grado, gli affini entro il 2°, le parti di unione civile e i conviventi di fatto.',
  'Come: un unico progetto di risoluzione della crisi; masse attive e passive restano DISTINTE (si fanno tabelle separate per ciascun membro); un solo OCC.',
  'Vale per: RDC, concordato minore e, secondo la giurisprudenza, anche per la liquidazione controllata.',
  'Nel concordato minore familiare le maggioranze si calcolano per masse.',
]);

// =====================================================================
// 11. CONFRONTO
// =====================================================================
C.push(h1('11. Le 4 procedure a confronto'));
table(['', 'RDC consumatore', 'Concordato minore', 'Liquidazione controllata', 'Incapiente'], [
  ['Articoli', '67-73', '74-83', '268-277 (+278-282)', '283'],
  ['Chi', 'Solo consumatore', 'Non consumatore (impresa minore, professionista, agricolo...)', 'Tutti i sovraindebitati', 'Persona fisica senza alcuna utilità'],
  ['Avvocato', 'Non necessario (prevalente)', 'Sì (di regola)', 'Prassi non uniforme', 'Prassi non uniforme'],
  ['Voto creditori', 'No', 'Sì (silenzio = sì)', 'No', 'No (opposizione)'],
  ['Si tengono i beni?', 'Sì, se il piano lo prevede', 'Sì, se il piano lo prevede', 'No: tutto va venduto (salvo esclusi)', 'Non ci sono beni'],
  ['Relazione OCC', 'art. 68 c.2-3', 'art. 76 c.2-3 (più ampia: convenienza, classi)', 'art. 269 c.2', 'art. 283'],
  ['Atto del giudice', 'Decreto di ammissione, poi sentenza di omologa', 'Decreto di apertura, voto, sentenza di omologa', 'Sentenza di apertura', 'Decreto'],
  ['Meritevolezza', 'No colpa grave, malafede, frode', 'No atti in frode', 'Rileva per l’esdebitazione (art. 280)', 'No frode, dolo, colpa grave'],
  ['Esdebitazione', 'Con l’esecuzione del piano', 'Con l’esecuzione del piano', 'Dopo chiusura o 3 anni', 'Subito, con controllo successivo sulle sopravvenienze'],
], [1500, 1800, 2000, 1900, 1826]);

// =====================================================================
// 12. ERRORI E CHECKLIST
// =====================================================================
C.push(h1('12. Errori da evitare e checklist prima del deposito'));
C.push(h2('Gli errori più frequenti'));
table(['Errore', 'Perché è grave', 'Come evitarlo'], [
  ['Dimenticare un creditore', 'Il giudice può ritenere la documentazione non attendibile; il creditore dimenticato può contestare', 'Incrociare sempre: racconto del debitore + Centrale Rischi + CRIF + estratto di ruolo + cassetto fiscale'],
  ['Spese di mantenimento “gonfiate” e non documentate', 'Il giudice le taglia e il piano non regge, o sembra che si voglia pagare meno', 'Solo spese con ricevuta o comunque ragionevoli rispetto ai dati ISTAT'],
  ['Numeri diversi tra ricorso e relazione', 'Perdita di credibilità di tutto il fascicolo', 'Un solo file Excel “madre” da cui prendere tutti i numeri'],
  ['Relazione che copia il ricorso', 'La relazione deve essere una VERIFICA indipendente, non un riassunto', 'Per ogni affermazione scrivere “ho verificato su ...”'],
  ['Mettere in prededuzione l’avvocato o il consulente del debitore', 'La giurisprudenza lo esclude: hanno solo privilegio (art. 2751-bis n.2 c.c.)', 'In prededuzione solo il compenso OCC, da accantonare'],
  ['Ignorare gli atti dei 5 anni', 'Vendite o donazioni ai familiari fanno pensare a frode', 'Chiedere espressamente e spiegare ogni atto'],
  ['Scegliere la RDC per chi non è consumatore', 'Domanda inammissibile', 'Analizzare l’origine di ogni debito (fideiussioni per società, debiti da partita IVA)'],
  ['Citare norme o sentenze non verificate', 'Errore grave e imbarazzante davanti al giudice', 'Controllare sempre su Normattiva e sulle banche dati; con Claude chiedere la fonte e verificarla'],
], [2600, 3200, 3226]);
C.push(h2('Checklist finale (spunta tutto prima di darmi la bozza)'));
list([
  'Procedura scelta motivata e requisiti soggettivi verificati.',
  'Tribunale competente corretto (residenza / sede).',
  'Elenco creditori completo, incrociato con tutte le banche dati, con prelazioni corrette.',
  'Somma dei debiti uguale in ricorso, relazione e tabelle.',
  'Patrimonio completo con valori documentati.',
  'Atti di straordinaria amministrazione dei 5 anni indicati (o dichiarazione negativa).',
  'Dichiarazioni dei redditi dei 3 anni allegate.',
  'Redditi e spese del nucleo con documenti.',
  'Merito creditizio: tabella per ogni finanziamento.',
  'Confronto con la liquidazione fatto.',
  'Comunicazioni a fisco/AdER fatte e risposte allegate.',
  'Casellario e carichi pendenti allegati.',
  'Istanze (misure protettive, mutuo, limite di reddito) scritte nelle conclusioni.',
  'Allegati numerati in ordine e richiamati nel testo (“doc. 5”).',
  'Nessun [VERIFICARE] o “xxx” rimasto nel testo.',
  'Rilettura completa per refusi, nomi e date.',
]);

// =====================================================================
// 13. CLAUDE
// =====================================================================
C.push(h1('13. Come lavorare con Claude su una pratica'));
par('Claude ti aiuta a scrivere più in fretta e a non dimenticare nulla, ma la responsabilità del contenuto resta dello studio. Ecco come usarlo bene.');
flow('Schema 8 – Il metodo di lavoro con Claude', [
  ['1. Prepara i dati', 'Tabelle A-E complete in Excel e un riassunto della storia del debitore.'],
  ['2. Chiedi a Claude una cosa alla volta', 'Esempio: “Scrivi il paragrafo sulle cause dell’indebitamento usando questi fatti: ...”.\nNon chiedere “fammi tutto il ricorso” in un colpo solo.'],
  ['3. Indica sempre procedura e norma', '“Siamo in una RDC, art. 68 c.2 lett. a CCII.” Così il testo è preciso.'],
  ['4. Controlla ogni numero e ogni citazione', 'Se Claude cita una norma o una sentenza, verificala tu. Se qualcosa non torna, chiediglielo.'],
  ['5. Rileggi e porta a me la bozza', 'Con la checklist del cap. 12 compilata.', VERDE],
], 8000);
box('Esempi di richieste utili', [
  'Tabelle: “Ecco i debiti di Rossi (incollo i dati). Ordinali per prelazione e calcola i totali per categoria.”',
  'Calcoli: “Calcola la soglia dell’art. 283 c.2 per un nucleo di 3 persone con l’assegno sociale 2026 e confrontala con un reddito di 19.000 €.”',
  'Merito creditizio: “Per ogni finanziamento di questa lista calcola il reddito residuo e confrontalo con assegno sociale × scala ISEE dell’anno.”',
  'Controllo: “Ecco la bozza del ricorso e la bozza della relazione: trova le differenze nei numeri e le parti mancanti rispetto all’art. 76 CCII.”',
  'Spiegazioni: “Non ho capito la differenza tra privilegio generale e speciale: spiegamela con un esempio.”',
], AZZURRO);
box('Dati personali', [
  'Le pratiche contengono dati personali e sensibili (salute, condanne, famiglia). Usa solo i dati che servono e segui le regole dello studio sulla riservatezza.',
  'Non salvare documenti dei clienti in cartelle condivise o pubbliche. Gli esempi di questo vademecum sono inventati.',
], ROSA);

// =====================================================================
// APPENDICE: FAC-SIMILE
// =====================================================================
C.push(h1('Appendice A – Fac-simile di ricorso (RDC del consumatore)'));
nota('Traccia di partenza. Le parti tra [parentesi quadre] vanno sostituite. Ogni affermazione va adattata al caso e collegata ai documenti.');
facsimile([
  'TRIBUNALE DI [FORLÌ]',
  'RICORSO PER L’OMOLOGAZIONE DEL PIANO DI RISTRUTTURAZIONE',
  'DEI DEBITI DEL CONSUMATORE (artt. 67 e ss. D.Lgs. 14/2019)',
  '',
  'Il sig. [NOME COGNOME], nato a [...] il [...], C.F. [...], residente in [...],',
  'via [...], indirizzo PEC per le comunicazioni [...],',
  'con l’ausilio dell’Organismo di composizione della crisi [NOME OCC],',
  'Gestore della crisi nominato [dott./dott.ssa ...] (doc. 1 – nomina),',
  '',
  'PREMESSO CHE',
  '',
  '1. IL RICORRENTE E IL SUO NUCLEO FAMILIARE',
  'Il ricorrente, di anni [..], è [lavoratore dipendente presso ... con',
  'contratto a tempo indeterminato], coniugato con [...] e padre di [...],',
  'con cui convive in [...] (doc. 2 – stato di famiglia).',
  '',
  '2. LA SUSSISTENZA DEI REQUISITI DI ACCESSO',
  '2.1 Competenza. Il ricorrente risiede in [...], nel circondario di',
  'codesto Tribunale, competente ai sensi dell’art. 27, comma 2, CCII.',
  '2.2 Qualità di consumatore. I debiti sono stati contratti per esigenze',
  'personali e familiari, estranee a qualsiasi attività imprenditoriale o',
  'professionale (art. 2, comma 1, lett. e, CCII).',
  '2.3 Sovraindebitamento. A fronte di debiti per complessivi € [...],',
  'il ricorrente dispone di un reddito netto medio mensile di € [...] e di',
  'un patrimonio di € [...]: le rate dovute (€ [...]/mese) superano la',
  'capacità di pagamento (art. 2, comma 1, lett. c, CCII).',
  '2.4 Assenza di condizioni ostative. Il ricorrente non ha beneficiato di',
  'esdebitazione nei cinque anni precedenti né per due volte e non ha',
  'determinato il sovraindebitamento con colpa grave, malafede o frode',
  '(art. 69, comma 1, CCII), come illustrato al punto 3 e nella relazione.',
  '',
  '3. LE CAUSE DELL’INDEBITAMENTO',
  '[Racconto cronologico con richiamo ai documenti: doc. ...]',
  '',
  '4. LA SITUAZIONE DEBITORIA',
  '[Tabella dei creditori con importi e cause di prelazione – doc. ...]',
  '',
  '5. IL PATRIMONIO',
  '[Beni e valori – doc. ...]',
  '',
  '6. ATTI DI STRAORDINARIA AMMINISTRAZIONE NEGLI ULTIMI CINQUE ANNI',
  '[Elenco, oppure: “Il ricorrente non ha compiuto atti di straordinaria',
  'amministrazione negli ultimi cinque anni.”]',
  '',
  '7. REDDITI E SPESE DEL NUCLEO FAMILIARE',
  '[Tabella redditi; tabella spese; dichiarazioni dei redditi – doc. ...]',
  '',
  '8. LA PROPOSTA E IL PIANO',
  'Il ricorrente propone il pagamento di € [...] mensili per [..] mesi,',
  'per complessivi € [...], destinati: (i) al compenso dell’OCC in',
  'prededuzione, mediante accantonamento; (ii) ai creditori privilegiati',
  'nella misura del [..]%; (iii) ai creditori chirografari nella misura',
  'del [..]%. [Trattamento del mutuo – art. 67 c.5; cessione del quinto –',
  'art. 67 c.4.]',
  '',
  '9. FATTIBILITÀ E CONVENIENZA',
  '[Sostenibilità della rata; confronto con l’alternativa liquidatoria.]',
  '',
  'TUTTO CIÒ PREMESSO, il ricorrente',
  'CHIEDE',
  'che l’Ill.mo Tribunale, ritenuta l’ammissibilità della proposta e del',
  'piano, voglia:',
  'a) disporne la pubblicazione e la comunicazione ai creditori ai sensi',
  '   dell’art. 70, comma 1, CCII;',
  'b) disporre, ai sensi dell’art. 70, comma 4, CCII, [la sospensione della',
  '   procedura esecutiva n. ... / il divieto di azioni esecutive e',
  '   cautelari / la sospensione delle trattenute della cessione del quinto];',
  'c) [autorizzare il pagamento delle rate del mutuo ai sensi dell’art. 67,',
  '   comma 5, CCII];',
  'd) all’esito, omologare il piano con sentenza ai sensi dell’art. 70,',
  '   comma 7, CCII.',
  '',
  'Si allegano: 1) relazione dell’OCC; 2) [...]; 3) [...].',
  '',
  '[Luogo], [data]                     Firma del ricorrente',
]);

C.push(h1('Appendice B – Fac-simile di frasi per la relazione del Gestore'));
nota('Formule ricorrenti da adattare. Ogni giudizio deve essere motivato con i fatti e i documenti verificati.');
facsimile([
  'INCARICO',
  'Il sottoscritto [...], iscritto all’albo dei gestori della crisi presso',
  'l’OCC [...], nominato in data [...], ha accettato l’incarico dichiarando',
  'di essere in possesso dei requisiti di cui all’art. 358 CCII e di non',
  'trovarsi in situazioni di conflitto di interessi con il debitore.',
  '',
  'COMUNICAZIONI ALL’AGENTE DELLA RISCOSSIONE E AGLI UFFICI FISCALI',
  'In data [...], entro sette giorni dall’incarico, il sottoscritto ha',
  'dato notizia dell’incarico ad AdER, all’Agenzia delle Entrate e a [Comune],',
  'ai sensi dell’art. 68, comma 4, CCII. Hanno risposto: [...].',
  '',
  'CAUSE DELL’INDEBITAMENTO E DILIGENZA (art. 68, c.2, lett. a)',
  'Dalla documentazione esaminata emerge che l’indebitamento trae origine',
  'da [...]. Il debitore [ha / non ha] assunto obbligazioni sproporzionate',
  'rispetto alle proprie capacità reddituali al momento della loro',
  'assunzione. Non emergono elementi di colpa grave, malafede o frode.',
  '',
  'RAGIONI DELL’INCAPACITÀ DI ADEMPIERE (art. 68, c.2, lett. b)',
  'A fronte di rate mensili complessive pari a € [...], il reddito',
  'disponibile, dedotte le spese di mantenimento, è pari a € [...].',
  '',
  'MERITO CREDITIZIO (art. 68, c.3)',
  'Come risulta dalla tabella che segue, alla data di concessione dei',
  'finanziamenti [...] il reddito residuo del debitore era inferiore alla',
  'soglia dell’assegno sociale moltiplicato per il parametro della scala',
  'di equivalenza ISEE; il finanziatore [...] non risulta pertanto aver',
  'adeguatamente valutato il merito creditizio.',
  '',
  'COMPLETEZZA E ATTENDIBILITÀ (art. 68, c.2, lett. c)',
  'Il sottoscritto ha verificato la documentazione depositata confrontandola',
  'con [Centrale Rischi, CRIF, estratto di ruolo, cassetto fiscale ...] e',
  'la ritiene completa e attendibile.',
  '',
  'CONVENIENZA RISPETTO ALL’ALTERNATIVA LIQUIDATORIA',
  'In caso di liquidazione controllata i creditori chirografari otterrebbero',
  'il [..]%; con il piano proposto ottengono il [..]%.',
  '',
  'COSTI PRESUNTI DELLA PROCEDURA (art. 68, c.2, lett. d)',
  'Compenso OCC stimato in € [...] secondo il DM 202/2014, oltre spese.',
  '',
  'CONCLUSIONI',
  '[Sintesi dei giudizi.]',
  '[Luogo], [data]                     Il Gestore della crisi',
]);

C.push(h1('Appendice C – Fonti e aggiornamento', false));
list([
  'Testo di riferimento: D.Lgs. 12 gennaio 2019 n. 14 (CCII), come modificato dal D.Lgs. 83/2022 e dal D.Lgs. 136/2024 (correttivo-ter, in vigore dal 28 settembre 2024). Consultare sempre il testo vigente su Normattiva.',
  'Punti segnati [VERIFICARE]: avvocato nelle procedure diverse dalla RDC, testo vigente degli artt. 80, 268 c.3 e 283 dopo il correttivo, termine (3 o 4 anni) e soglia delle sopravvenienze nell’esdebitazione dell’incapiente, compenso OCC ridotto, importo dell’assegno sociale dell’anno.',
  'Per i criteri pratici dei giudici (Forlì e altri Tribunali) vedi anche il documento dello studio “Sovraindebitamento: i criteri dei giudici” nella stessa cartella.',
]);

const doc = new Document({
  creator: 'Studio – materiale di formazione',
  title: 'Vademecum del sovraindebitamento',
  features: { updateFields: true },
  background: { color: 'FDFBF5' },
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'punti', levels: [
    { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 300 } } } },
    { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1000, hanging: 300 } } } },
  ] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('Vademecum sovraindebitamento – uso interno – pag. ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => fs.writeFileSync(process.argv[2] || 'Vademecum_Sovraindebitamento_ricorsi_e_relazioni.docx', buf));
