// Bozza del Bilancio di genere 2025 – ODCEC Forlì.
// Legge dati_2025.json: i campi null diventano segnaposto evidenziati in giallo.
// Uso: python3 grafici.py && node genera_bilancio.js [file.docx]
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, Footer, PageNumber, ImageRun, TableOfContents,
} = require('docx');

const D = JSON.parse(fs.readFileSync(path.join(__dirname, 'dati_2025.json'), 'utf8'));
const FONT = 'Verdana', INK = '1F2A37', ACCENT = '2E5E8C', GREY = '4B5563';
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, color: INK, size: 22, ...o });
const daFare = (t) => new TextRun({ text: `[DA INTEGRARE: ${t}]`, font: FONT, color: INK, size: 22, shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'FFE873' } });
const p = (c, o = {}) => new Paragraph({ children: (Array.isArray(c) ? c : [c]).map(x => typeof x === 'string' ? run(x) : x), spacing: { after: 140, line: 360 }, ...o });
const b = (c) => new Paragraph({ numbering: { reference: 'punti', level: 0 }, spacing: { after: 80, line: 340 }, children: (Array.isArray(c) ? c : [c]).map(x => typeof x === 'string' ? run(x) : x) });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, spacing: { after: 200 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 32, color: ACCENT })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 300, after: 140 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: ACCENT })] });
const fonte = (t) => p(run('Fonte: ' + t, { size: 18, color: GREY }));
const it = (n, dec = 0) => n.toLocaleString('it-IT', { minimumFractionDigits: dec, maximumFractionDigits: dec });
const pct = (a, tot) => it(a / tot * 100, 1) + '%';
const img = (file, w = 600, h = 315) => new Paragraph({ spacing: { after: 120 }, children: [new ImageRun({ type: 'png', data: fs.readFileSync(path.join(__dirname, 'grafici', file)), transformation: { width: w, height: h }, altText: { title: file, description: 'Grafico: i valori sono riportati nella tabella che segue', name: file } })] });

function table(head, rows, widths) {
  const total = widths.reduce((a, c) => a + c, 0);
  const cell = (t, w, hd) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: hd ? { type: ShadingType.CLEAR, color: 'auto', fill: 'DCE7F2' } : undefined,
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    children: [new Paragraph({ children: [t instanceof TextRun ? t : run(String(t), { size: 19, bold: hd })] })],
  });
  const tr = [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], true)) })];
  rows.forEach(r => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false)) })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: tr });
}
// tabella vuota per dati in arrivo: una riga evidenziata con la richiesta
const attesa = (head, cosa, widths) => table(head, [[daFare(cosa), ...head.slice(1).map(() => '…')]], widths);

const C = [];
const I = D.iscritti, CA = D.cassa_cnpadc;
const [f25, m25] = I.per_genere['2025'], tot25 = f25 + m25;

// ---------------- copertina
C.push(p(run('Ordine dei Dottori Commercialisti e degli Esperti Contabili di Forlì', { size: 22, color: GREY })));
C.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { before: 1800, after: 200 }, children: [new TextRun({ text: 'Bilancio di genere 2025', font: FONT, bold: true, size: 56, color: ACCENT })] }));
C.push(p(run('A cura del Comitato Pari Opportunità', { size: 26 })));
C.push(p(run('Assemblea degli iscritti – Cesena, 26 novembre 2026', { size: 22, color: GREY })));
C.push(p([run('BOZZA DI LAVORO del 6 ottobre 2026. ', { bold: true }), run('Le parti evidenziate in giallo aspettano i dati richiesti all’Ordine il 28/9/2026 (scadenza 31 ottobre) o il questionario agli iscritti.')], { spacing: { before: 1200, after: 140, line: 360 } }));

C.push(h2('Come leggere questo documento'));
C.push(p('Abbiamo scritto il bilancio per essere letto da tutte e tutti, anche da chi ha un disturbo specifico dell’apprendimento (DSA): carattere senza grazie, frasi brevi, sfondo crema, ogni grafico accompagnato dalla tabella con i numeri. Usiamo il femminile e il maschile; quando per brevità c’è solo il maschile, si intende riferito a tutte le persone.'));
C.push(p('Il documento è pubblicato solo in formato digitale sul sito dell’Ordine, leggibile anche con la sintesi vocale.'));

C.push(h2('Il 2025 in sei numeri'));
C.push(table(['', ''], [
  [run(pct(f25, tot25), { bold: true, size: 24 }), `le donne tra gli iscritti (${f25} su ${tot25}). In Italia sono il 33,9%.`],
  [run('61%', { bold: true, size: 24 }), 'le donne tra gli iscritti di 36-45 anni. Sotto i 56 anni le donne sono la maggioranza.'],
  [run('61%', { bold: true, size: 24 }), `le donne tra i praticanti (${D.praticanti['2025'][0]} su ${D.praticanti['2025'][0] + D.praticanti['2025'][1]}).`],
  [run('54,3%', { bold: true, size: 24 }), 'il divario di reddito tra professioniste e professionisti iscritti alla Cassa: più alto della media regionale (48,2%) e nazionale (46,8%).'],
  [run('3 su 4', { bold: true, size: 24 }), 'le cariche del Consiglio 2026-2030 affidate a donne: Presidente, Segretaria, Tesoriera. Nel Consiglio le donne sono 5 su 11.'],
  [run('27,7%', { bold: true, size: 24 }), 'le relatrici negli eventi formativi del 2025: poco più di una su quattro.'],
], [2000, 7026]));
C.push(p(run('Nota: iscritti e praticanti al 31/12/2025; divario di reddito sui redditi 2024 comunicati alla Cassa nel 2025; Consiglio in carica dal 23/2/2026.', { size: 18, color: GREY })));

C.push(new Paragraph({ pageBreakBefore: true, children: [new TextRun({ text: 'Sommario', font: FONT, bold: true, size: 28, color: ACCENT })] }));
C.push(new TableOfContents('Sommario', { hyperlink: true, headingStyleRange: '1-2' }));
C.push(p(run('Il sommario si aggiorna in Word con tasto destro → Aggiorna campo.', { size: 18, color: GREY })));

// ---------------- 1 premessa e metodo
C.push(h1('1. Premessa e metodo'));
C.push(p('Il bilancio di genere fotografa la presenza di donne e uomini nella professione e nell’Ordine di Forlì e misura le differenze di reddito, di carriera e di rappresentanza. Serve a decidere dove intervenire e, l’anno dopo, a verificare se gli interventi hanno funzionato.'));
C.push(p('Seguiamo le Linee guida del Comitato Pari Opportunità nazionale del 14 marzo 2024 (Informativa CNDCEC n. 31/2024) e il Regolamento dei CPO, che all’art. 3 lett. i) affida al Comitato la redazione del bilancio di genere dell’Ordine.'));
C.push(h2('Fonti'));
C.push(table(['Dato', 'Fonte', 'Data di riferimento'], [
  ['Iscritti, Elenco speciale, praticanti', 'Albo e Registro del tirocinio (estrazione della Segreteria)', '31/12/2024 e 31/12/2025'],
  ['Serie storica 2020-2023', 'Bilanci di genere 2023 e 2024 dell’Ordine', '31/12 di ogni anno'],
  ['Redditi e volume d’affari', 'Cassa Dottori Commercialisti (CNPADC)', 'Redditi 2023 e 2024'],
  ['Confronto nazionale', 'Questionario CNPO 2025 e bilancio di genere del CNDCEC', '31/12/2024'],
  [daFare('Cassa Ragionieri, OCC, Camera di Commercio, spese'), 'Richieste di settembre-ottobre 2026', '…'],
], [3000, 4026, 2000]));
C.push(h2('Regole di calcolo'));
C.push(b('Divario di reddito = (reddito medio uomini − reddito medio donne) / reddito medio uomini. È lo stesso indicatore usato da Eurostat e dalle Linee guida.'));
C.push(b('Usiamo le medie ponderate fornite dalla Cassa, non la media delle fasce d’età.'));
C.push(b('Tutela dei dati: pubblichiamo solo dati aggregati e non mostriamo gruppi con meno di 5 persone quando potrebbero renderle riconoscibili.'));
C.push(h2('Nota di raccordo con il bilancio 2024'));
C.push(p('Rispetto al bilancio 2024 aggiungiamo tre informazioni che allora mancavano: i praticanti per genere (le donne erano il 60% già nel 2024), il divario di reddito degli iscritti di Forlì (49,2% sui redditi 2023, mentre il bilancio 2024 riportava solo i dati regionali e nazionali) e le sezioni dell’Albo per genere.'));

// ---------------- 2 contesto normativo
C.push(h1('2. Il quadro normativo in breve'));
C.push(b('Costituzione, artt. 3, 37, 51 e 117: uguaglianza, parità nel lavoro e nell’accesso alle cariche.'));
C.push(b('Codice delle pari opportunità (d.lgs. 198/2006) e legge 162/2021 sulla parità salariale, che ha introdotto la certificazione della parità di genere (UNI/PdR 125:2022).'));
C.push(b('D.lgs. 139/2005 sull’ordinamento della professione e Regolamento dei Comitati Pari Opportunità degli Ordini.'));
C.push(p([run('Direttiva (UE) 2023/970 del 10 maggio 2023 sulla trasparenza retributiva: gli Stati dovevano recepirla entro il 7 giugno 2026. '), daFare('stato del recepimento in Italia alla data del bilancio')]));
C.push(p('Per la nostra professione la direttiva conta due volte: negli studi come datori di lavoro e verso le imprese clienti, che dovranno misurare e comunicare il proprio divario retributivo.'));

// ---------------- 3 CPO
C.push(h1('3. Il Comitato Pari Opportunità'));
C.push(h2('Composizione'));
C.push(p(`Il Comitato in carica è stato nominato dal Consiglio dell’Ordine il ${D.cpo.nomina} per il quadriennio 2026-2030. È composto da 7 persone: ${D.cpo.componenti[0]} donne e ${D.cpo.componenti[1]} uomini.`));
C.push(D.cpo.nomi ? table(['Nome', 'Ruolo'], D.cpo.nomi, [5000, 4026]) : table(['Nome', 'Ruolo'], [['Emanuela Lupi', 'Presidente'], [daFare('altri 6 componenti e ruoli'), '…']], [5000, 4026]));
C.push(p('Il Comitato precedente (2022-2026) era presieduto da Monica Orecchioni ed era composto da 5 donne e 2 uomini. Ha curato i bilanci di genere 2023 e 2024.'));
C.push(h2('Cosa fa il Comitato'));
C.push(p('Il CPO promuove la parità tra iscritte, iscritti e tirocinanti: analizza i dati, propone iniziative al Consiglio, organizza eventi formativi, vigila contro le discriminazioni e redige il bilancio di genere.'));
C.push(h2('Risorse'));
C.push(p(['Budget assegnato dal Consiglio per il 2026: ', D.cpo.budget != null ? run(`${it(D.cpo.budget)} €`) : daFare('da chiedere al Consiglio; se non assegnato, motivazione come chiedono le Linee guida')]));
C.push(p(['Personale della Segreteria di supporto al CPO: ', D.cpo.personale || daFare('nominativo o ufficio di riferimento')]));
C.push(p('Fino al 2025 il Comitato ha operato senza budget, a titolo gratuito o con sponsor.'));

// ---------------- 4 attività
C.push(h1('4. Le attività'));
C.push(h2('2025'));
C.push(D.attivita_2025 ? p(D.attivita_2025) : p([daFare('attività del CPO 2022-2026 nel 2025: eventi, riunioni CPO regionale, convegno nazionale CPO')]));
C.push(h2('2026, primo anno del nuovo Comitato'));
C.push(b('Nomina del Comitato (23 febbraio 2026) e avvio della collaborazione stabile con il CPO dell’Ordine di Ravenna e con gli altri Ordini professionali della provincia.'));
C.push(b('Partecipazione al convegno “Zero disparità, più produttività – La certificazione della parità di genere” (Cesena, 4 giugno 2026), con Camera di Commercio della Romagna e Consigliera di Parità regionale.'));
C.push(b('Check-up dei bilanci di genere 2023 e 2024 e confronto con 17 bilanci di altri Ordini.'));
C.push(b('Richiesta dati all’Ordine e alla Fondazione (28 settembre 2026), call con la Presidente dell’Ordine (7 ottobre) e richieste a OCC Romagna (5 ottobre), Cassa Ragionieri, Tribunale e Camera di Commercio (12 ottobre).'));
C.push(b([run('Questionario breve agli iscritti su conciliazione, carriera, DSA e inclusione. '), daFare('data di invio e numero di risposte')]));
C.push(b(daFare('altre attività 2026 fino all’assemblea')));

// ---------------- 5 contesto iscritti
C.push(h1('5. Chi siamo: iscritte e iscritti'));
C.push(p(`Al 31/12/2025 l’Ordine conta ${tot25} iscritti: ${f25} donne (${pct(f25, tot25)}) e ${m25} uomini. La quota di donne è stabile da tre anni ed è molto più alta della media nazionale (33,9%) e regionale (34,2%).`));
C.push(h2('Albo ed Elenco speciale per genere'));
const sezRow = (k) => { const a = I.sezioni_2024[k], c = I.sezioni_2025[k]; return [k, `${a[0]} / ${a[1]}`, `${c[0]} / ${c[1]}`, pct(c[0], c[0] + c[1])]; };
C.push(table(['Sezione', '2024 donne / uomini', '2025 donne / uomini', '% donne 2025'], [
  ...Object.keys(I.sezioni_2025).map(sezRow),
  ['Totale', '299 / 348', `${f25} / ${m25}`, pct(f25, tot25)],
], [2600, 2200, 2200, 2026]));
fonte('Albo dell’Ordine, elaborazione CPO. Nella sezione B e nell’Elenco speciale i numeri sono piccoli: le percentuali vanno lette con prudenza.');

C.push(h2('Serie storica'));
const st = I.serie_totali;
C.push(table(['Anno', 'Albo', 'Elenco speciale', 'Totale', 'Donne', '% donne'], [
  ['2020', st['2020'][0], st['2020'][1], st['2020'][0] + st['2020'][1], '—', '—'],
  ...['2021', '2022', '2023', '2024', '2025'].map(y => { const g = I.per_genere[y]; const t = g[0] + g[1]; return [y, st[y][0], st[y][1], t, g[0], pct(g[0], t)]; }),
], [1200, 1300, 1700, 1300, 1700, 1826]));
fonte('bilanci di genere 2023 e 2024 dell’Ordine (2020); Albo dell’Ordine 2021-2025. Per il 2020 non è disponibile il dato per genere.');
C.push(p('In cinque anni la quota di donne è passata dal 44,9% al 45,7%, con un massimo del 46,2% nel 2024. Gli iscritti sono rimasti stabili intorno ai 650: le donne sono la stessa cifra di cinque anni fa (292), gli uomini sono 12 in meno.'));

C.push(h2('Età'));
C.push(img('eta_genere.png'));
const ce = I.classi_eta_2025;
C.push(table(['Classe d’età', 'Donne', 'Uomini', '% donne'], Object.entries(ce).map(([k, v]) => [k, v[0], v[1], pct(v[0], v[0] + v[1])]), [3000, 2000, 2000, 2026]));
fonte('Albo dell’Ordine al 31/12/2025, classi del questionario nazionale CNPO.');
C.push(p(`Le donne sono più giovani: età media ${it(I.eta_media['2025'][0], 1)} anni contro ${it(I.eta_media['2025'][1], 1)} degli uomini. Sotto i 56 anni sono la maggioranza in ogni classe; tra gli over 65 sono poco più di una su quattro. Con il ricambio generazionale l’Albo arriverà alla parità numerica nei prossimi anni.`));
C.push(table(['Generazione', 'Donne', 'Uomini', '% donne'], Object.entries(I.generazioni_2025).map(([k, v]) => [k, v[0], v[1], pct(v[0], v[0] + v[1])]), [4000, 1600, 1600, 1826]));
fonte('Albo dell’Ordine al 31/12/2025, generazioni come nei bilanci precedenti.');

C.push(h2('Praticanti'));
const pr = D.praticanti;
C.push(table(['Anno', 'Totale', 'Donne', 'Uomini', '% donne'], [
  ['2020', pr['2020'], '—', '—', '—'],
  ...['2021', '2022', '2023', '2024', '2025'].map(y => [y, pr[y][0] + pr[y][1], pr[y][0], pr[y][1], pct(pr[y][0], pr[y][0] + pr[y][1])]),
], [1500, 1500, 2000, 2000, 2026]));
fonte('Registro del tirocinio.');
C.push(p('Negli ultimi due anni sei praticanti su dieci sono donne; nel 2023 erano meno della metà. È la base della professione di domani: il punto è che restino e crescano.'));
fonte('Registro del tirocinio. Per il 2023 il bilancio 2024 riportava 28 praticanti, l’estrazione ne conta 29.');

C.push(h2('Chi entra e chi esce'));
const fs_ = D.flussi_stima_2025;
if (D.flussi_ufficiali) {
  C.push(table(['Anno', 'Iscrizioni donne / uomini', 'Cancellazioni donne / uomini', 'Motivi delle cancellazioni'], D.flussi_ufficiali, [1300, 2200, 2400, 3126]));
  fonte('Segreteria dell’Ordine. Iscrizioni all’Albo e all’Elenco speciale; il motivo della cancellazione manca per 8 casi del 2023.');
  const u = D.flussi_u50.totale;
  C.push(p(`In cinque anni sono entrate 29 donne e 23 uomini, ma sono uscite 29 donne e 34 uomini. Il 2025 è un’eccezione: entrano 2 donne e 8 uomini, mentre escono 9 donne e 9 uomini. Chi lascia la professione presto è soprattutto donna: tra le cancellazioni sotto i 50 anni le donne sono ${u[0]} e gli uomini ${u[1]}, e solo nel 2025 le donne sono 4. Quasi tutte si cancellano su richiesta.`));
}
else {
  C.push(table(['2024 → 2025', 'Donne', 'Uomini'], [['Nuovi ingressi nell’Albo', fs_.ingressi[0], fs_.ingressi[1]], ['Uscite dall’Albo', fs_.uscite[0], fs_.uscite[1]]], [4000, 2500, 2526]));
  C.push(p([run('Stima del CPO dal confronto degli elenchi anonimi 2024 e 2025. '), daFare('dato 1: iscrizioni e cancellazioni 2021-2025 per genere, età e motivo')]));
}
C.push(p('Anche a livello nazionale le donne sono passate dal 39% al 46% delle cancellazioni dalla Cassa tra il 2018 e il 2023. Il questionario chiederà a chi pensa di lasciare quali sono i motivi.'));

// ---------------- 6 governance
C.push(h1('6. Chi decide: Consiglio e organi'));
C.push(table(['Organo', 'Donne', 'Uomini', '% donne'], [
  ['Consiglio 2022-2026', 4, 7, pct(4, 11)],
  D.consiglio['2026-2030'] ? ['Consiglio 2026-2030', ...D.consiglio['2026-2030'], pct(D.consiglio['2026-2030'][0], D.consiglio['2026-2030'][0] + D.consiglio['2026-2030'][1])] : ['Consiglio 2026-2030', daFare('dato 3'), '…', '…'],
  ['CPO 2026-2030', D.cpo.componenti[0], D.cpo.componenti[1], pct(D.cpo.componenti[0], 7)],
  ...(D.organi ? D.organi.map(o => [o[0], o[1], o[2], pct(o[1], o[1] + o[2])]) : [['Consiglio di disciplina, revisori, commissioni', daFare('dato 3'), '…', '…']]),
], [3800, 1600, 1600, 2026]));
fonte('Segreteria dell’Ordine. Il Consiglio 2022-2026 è stato in carica dal 25/2/2022, il Consiglio 2026-2030 dal 23/2/2026.');
C.push(p('Nel Consiglio 2022-2026 le donne erano il 36%, dieci punti sotto la loro presenza tra gli iscritti. Nel Consiglio 2026-2030 sono il 45%, in linea con l’Albo, e tre cariche su quattro sono di donne: la Presidente, dott.ssa Sara Pennacchi, la Segretaria e la Tesoriera. A livello nazionale le Presidenti di Ordine sono passate da 19 a 30 su 132.'));
C.push(p(`Cariche 2022-2026: ${D.consiglio.cariche_2022}. Cariche 2026-2030: ${D.consiglio.cariche_2026_2030}.`));
C.push(p('Negli organi di controllo e di disciplina le donne sono la maggioranza o quasi, e presiedono sia il Consiglio di disciplina sia il Collegio dei revisori. Nelle commissioni invece il quadro è diverso: le componenti sono quasi la metà, ma nel 2026-2030 tre presidenze su quattro sono andate a uomini.'));

// ---------------- 7 redditi
C.push(h1('7. Quanto guadagniamo: il divario di reddito'));
C.push(p('I dati sono quelli della Cassa Dottori Commercialisti per gli iscritti attivi residenti nel territorio dell’Ordine. Non comprendono gli iscritti alla Cassa Ragionieri.'));
C.push(img('divario.png'));
const r = CA.reddito_medio, ia = CA.iscritti_attivi, dv = CA.divario_reddito;
const terr = ['Forlì', 'Emilia-Romagna', 'Italia'];
C.push(table(['', ...terr], [
  ['Iscritti attivi (% donne)', ...terr.map(t => `${it(ia[t][0])} (${it(ia[t][1], 1)}%)`)],
  ['Reddito medio donne', ...terr.map(t => `${it(r[t][0])} €`)],
  ['Reddito medio uomini', ...terr.map(t => `${it(r[t][1])} €`)],
  ['Divario di reddito, redditi 2024', ...terr.map(t => `${it(dv[t][1], 1)}%`)],
  ['Divario di reddito, redditi 2023', ...terr.map(t => `${it(dv[t][0], 1)}%`)],
  ['Divario di volume d’affari, 2024', ...terr.map(t => `${it(CA.divario_volume_affari[t], 1)}%`)],
], [3300, 1900, 1900, 1926]));
fonte('CNPADC, dati comunicati nel 2024 e nel 2025.');
C.push(p('A Forlì una professionista guadagna in media meno della metà di un collega. In un anno il reddito medio degli uomini è cresciuto dell’11,9%, quello delle donne dello 0,5%: il divario è salito di 5 punti e oggi supera la media regionale e nazionale.'));
C.push(h2('Il divario per età'));
const re = CA.rapporto_eta;
C.push(table(['Reddito delle donne in % di quello degli uomini', ...terr], Object.entries(re).map(([k, v]) => [k, ...v.map(x => x + '%')]), [3800, 1700, 1700, 1826]));
fonte('CNPADC, redditi 2024. Fino a 30 anni a Forlì gli iscritti attivi sono 19 (13 donne, 6 uomini); oltre 65 anni le donne sono 8: valori da leggere con prudenza.');
C.push(p('Fino a 50 anni le professioniste di Forlì stanno meglio che in regione e in Italia. Il divario complessivo così alto nasce soprattutto tra i 51 e i 65 anni, dove si concentrano gli uomini con i redditi più alti, e dal fatto che le donne iscritte sono più giovani. Il problema non è tanto l’ingresso nella professione, quanto la crescita a metà carriera.'));
C.push(h2('Cassa Ragionieri'));
C.push(D.cassa_cnpr ? table(['', 'Donne', 'Uomini', 'Divario'], D.cassa_cnpr, [3800, 1700, 1700, 1826]) : attesa(['Iscritti CNPR di Forlì', 'Donne', 'Uomini', 'Divario'], 'dato 8, reddito e volume d’affari medi per genere ed età', [3800, 1700, 1700, 1826]));

// ---------------- 8 formazione
C.push(h1('8. Formazione e conciliazione'));
C.push(h2('Relatrici e relatori'));
C.push(D.formazione_eventi ? table(['', 'Eventi', 'Relatrici', 'Relatori', '% relatrici', 'Panel di soli uomini'], D.formazione_eventi, [1800, 1300, 1400, 1400, 1500, 1626])
  : attesa(['Anno', 'Eventi', 'Relatrici', 'Relatori', '% relatrici', 'Panel di soli uomini'], 'dato 4, eventi 2025', [1800, 1300, 1400, 1400, 1500, 1626]));
if (D.formazione_eventi) { fonte('Segreteria dell’Ordine: 51 eventi del 2025 (convegni, webinar, seminari, corsi). Un evento in aula e in webinar conta due volte.'); C.push(p('Su 267 interventi di relatori, 74 sono di donne: poco più di uno su quattro, contro il 46% di donne tra gli iscritti. In 12 eventi su 51 hanno parlato solo uomini, mai il contrario.')); }
C.push(h2('Esoneri dalla formazione per maternità, paternità e cura'));
C.push(D.esoneri ? table(['Motivo', 'Donne', 'Uomini'], D.esoneri, [4000, 2500, 2526]) : attesa(['Motivo', 'Donne', 'Uomini'], 'dato 5', [4000, 2500, 2526]));
C.push(p('Questi dati dicono quanto il lavoro di cura pesa ancora sulle donne della professione. Il confronto con le indennità di maternità e paternità della Cassa completerà il quadro nel bilancio 2026.'));

// ---------------- 9 incarichi
C.push(h1('9. Incarichi e cariche: una parte nuova'));
C.push(p('Le Linee guida suggeriscono di guardare oltre l’Albo: chi riceve gli incarichi dal Tribunale e dagli organismi di composizione della crisi, chi siede nei collegi sindacali delle imprese. Pochi Ordini lo fanno (Reggio Emilia, Venezia, Parma). Abbiamo chiesto i dati al Tribunale, all’OCC Romagna e alla Camera di Commercio, anche con il valore degli incarichi e non solo il numero.'));
C.push(h2('Tribunale di Forlì'));
if (D.tribunale_nota && !D.tribunale) C.push(p(D.tribunale_nota));
else C.push(D.tribunale ? table(['Tipo di incarico', 'Donne', 'Uomini', '% donne'], D.tribunale, [3800, 1600, 1600, 2026]) : attesa(['Tipo di incarico 2021-2025', 'Donne', 'Uomini', '% donne'], 'dato 9: curatele, liquidazioni, CTU, custodie, delegati alle vendite', [3800, 1600, 1600, 2026]));
C.push(h2('OCC Romagna'));
if (D.occ_nota && !D.occ) C.push(p(D.occ_nota));
C.push(D.occ ? table(['', 'Donne', 'Uomini', '% donne'], D.occ, [3800, 1600, 1600, 2026]) : attesa(['Gestori iscritti a Forlì e incarichi 2021-2025', 'Donne', 'Uomini', '% donne'], 'dato 10', [3800, 1600, 1600, 2026]));
C.push(h2('Collegi sindacali e revisione nelle imprese'));
C.push(D.cciaa ? table(['Carica al 31/12/2025', 'Donne', 'Uomini', '% donne'], D.cciaa, [3800, 1600, 1600, 2026]) : attesa(['Carica al 31/12/2025', 'Donne', 'Uomini', '% donne'], 'dato 11, Camera di Commercio della Romagna', [3800, 1600, 1600, 2026]));
C.push(p('Se i dati non arriveranno in tempo, questa parte sarà completata nel bilancio 2026 e resterà qui come impegno.'));

// ---------------- 10 spese
C.push(h1('10. Le spese dell’Ordine con la lente di genere'));
C.push(p('È il cuore del “gender budgeting” secondo il Consiglio d’Europa, ripreso dalle Linee guida, ma nei 17 bilanci di Ordini che abbiamo esaminato nessuno lo applica. Dividiamo le spese 2025 dell’Ordine e della Fondazione in tre gruppi:'));
C.push(b('neutre: non toccano in modo diverso donne e uomini (affitto, utenze, software, quote al Consiglio Nazionale);'));
C.push(b('dirette a ridurre le diseguaglianze di genere: attività del CPO, progetti nelle scuole;'));
C.push(b('attività socio-culturali: visite guidate, eventi sportivi e altre iniziative per gli iscritti, da leggere guardando chi vi partecipa.'));
if (D.spese_note) C.push(p(D.spese_note));
C.push(D.spese ? table(['Gruppo', 'Importo', '% sul totale'], D.spese, [4000, 2500, 2526]) : attesa(['Gruppo', 'Importo 2025', '% sul totale'], 'consuntivo 2025 per capitolo dalla Tesoriera', [4000, 2500, 2526]));

// ---------------- 11 questionario
C.push(h1('11. La voce degli iscritti: il questionario'));
C.push(p('Il questionario breve chiede di conciliazione tra lavoro e famiglia, di carriera, di esperienze di discriminazione e, per la prima volta tra i bilanci di genere degli Ordini che abbiamo esaminato, di DSA e inclusione. È anonimo e i risultati sono pubblicati solo in forma aggregata.'));
C.push(D.questionario ? p(D.questionario) : p([daFare('numero di risposte per genere, principali risultati, grafici')]));

// ---------------- 12 obiettivi
C.push(h1('12. Obiettivi per il 2027'));
C.push(p('Nessuno dei bilanci che abbiamo esaminato fissa obiettivi numerici da verificare l’anno dopo. Noi proviamo: ogni obiettivo ha un indicatore, il valore di partenza e un semaforo che nel bilancio 2026 dirà se l’abbiamo raggiunto. Gli obiettivi sono una proposta del CPO da condividere con il Consiglio.'));
C.push(table(['Obiettivo', 'Indicatore', 'Partenza 2025', 'Traguardo 2027'], [
  ['Bilancio completo', 'Contenuti minimi delle Linee guida coperti', 'circa 50% nel 2024', '100%'],
  ['Più relatrici', '% relatrici negli eventi dell’Ordine', D.eventi_kpi.pct_relatrici, 'almeno 40%'],
  ['Nessun panel di soli uomini', 'Eventi con soli relatori uomini', D.eventi_kpi.solo_uomini, '0 tra gli eventi con 3 o più relatori'],
  ['Presidenze delle commissioni', 'Commissioni presiedute da donne', '1 su 4', 'almeno 2 su 4 al prossimo rinnovo'],
  ['Il CPO ha risorse', 'Budget assegnato', 'nessuno', daFare('importo da concordare')],
  ['Ascoltare gli iscritti', 'Risposte al questionario', daFare('dato del questionario'), 'almeno 25% degli iscritti'],
  ['Divario di reddito', 'Divario Forlì rispetto alla media regionale', '+6,1 punti', 'sotto la media regionale entro il 2028'],
  ['Bilancio accessibile', 'Sintesi facile da leggere e PDF per sintesi vocale', 'sì (2025)', 'sì, anche versione audio'],
], [2300, 2700, 2000, 2026]));
C.push(p('Il divario di reddito dipende da molti fattori fuori dal controllo dell’Ordine. Il CPO lo segue con azioni mirate alla fascia di metà carriera: formazione su specializzazioni e incarichi, mentoring, rete con la Consigliera di Parità.'));

// ---------------- 13 conclusioni
C.push(h1('13. Conclusioni'));
C.push(p('Forlì è un Ordine con molte donne, più giovani dei colleghi, e con un’ampia base di praticanti. La parità nei numeri è vicina. Non lo è nei redditi: il divario è alto e cresce, soprattutto a metà carriera. Non lo è ancora nella rappresentanza, anche se il nuovo Consiglio è guidato da una Presidente.'));
C.push(p('Con questo bilancio abbiamo voluto fare due passi in più: misurare dove si decidono le carriere (incarichi, cariche, formazione) e leggere con la lente di genere anche le spese dell’Ordine. Il prossimo anno sapremo se gli obiettivi fissati sono stati raggiunti.'));
C.push(p(daFare('data e firme dei componenti del CPO')));

// ---------------- appendice
C.push(h1('Appendice – Stato dei dati'));
const stato = (v, k) => v ? 'inserito' : (D.ricevuti && D.ricevuti[k]) ? `ricevuto ${D.ricevuti[k]}, da inserire` : 'in attesa';
C.push(table(['N.', 'Dato richiesto', 'A chi', 'Stato'], [
  ['1', 'Iscrizioni e cancellazioni 2021-2025', 'Segreteria', stato(D.flussi_ufficiali, 'flussi')],
  ['2', 'Iscritti e praticanti per genere 2021-2023', 'Segreteria', stato(D.iscritti.per_genere_2021_2022, 'serie')],
  ['3', 'Consiglio, cariche, organi, commissioni', 'Segreteria', stato(D.consiglio['2026-2030'] && D.organi, 'organi')],
  ['4', 'Eventi formativi 2025, relatrici e relatori', 'Segreteria', stato(D.formazione_eventi, 'eventi')],
  ['5', 'Esoneri per maternità, paternità e cura', 'Segreteria', stato(D.esoneri, 'esoneri')],
  ['6', 'Consuntivo 2025 per capitolo', 'Tesoriera', stato(D.spese, 'spese')],
  ['7', 'Spese socio-culturali della Fondazione', 'Fondazione', stato(D.spese_note, 'fondazione')],
  ['8', 'Redditi Cassa Ragionieri', 'CNPR', stato(D.cassa_cnpr)],
  ['9', 'Incarichi del Tribunale', 'Tribunale di Forlì', D.tribunale ? 'inserito' : D.tribunale_nota ? 'non concessi' : 'in attesa'],
  ['10', 'Gestori e incarichi', 'OCC Romagna', stato(D.occ)],
  ['11', 'Sindaci e revisori', 'Camera di Commercio', stato(D.cciaa)],
  ['–', 'Questionario agli iscritti', 'CPO', stato(D.questionario)],
], [700, 4300, 2300, 1726]));

const doc = new Document({
  creator: 'CPO ODCEC Forlì', title: 'Bilancio di genere 2025 – ODCEC Forlì', description: 'Bozza di lavoro',
  features: { updateFields: true },
  background: { color: 'FDFBF5' },
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } } },
  numbering: { config: [{ reference: 'punti', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1400, right: 1400 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run('Bilancio di genere 2025 – ODCEC Forlì – pag. ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[2] || path.join(__dirname, 'Bilancio_di_genere_2025_ODCEC_Forli_BOZZA.docx'), buf); console.log('ok'); });
