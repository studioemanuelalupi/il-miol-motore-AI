const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, Footer, PageNumber,
} = require('docx');

const FONT = 'Verdana', INK = '1F2A37', ACCENT = '2E5E8C', GREY = '4B5563';
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, color: INK, size: 22, ...o });
const p = (c, o = {}) => new Paragraph({ children: Array.isArray(c) ? c : [typeof c === 'string' ? run(c) : c], spacing: { after: 120, line: 340 }, ...o });
const b = (t) => new Paragraph({ numbering: { reference: 'punti', level: 0 }, spacing: { after: 60, line: 320 }, children: typeof t === 'string' ? [run(t)] : t });
const h1 = (t, brk = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: brk, spacing: { after: 160 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 30, color: ACCENT })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 280, after: 120 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: ACCENT })] });
const nota = (t) => p(run(t, { size: 18, color: GREY }));

function table(head, rows, widths, opts = {}) {
  const total = widths.reduce((a, c) => a + c, 0);
  const cell = (t, w, hd, fill) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: hd ? { type: ShadingType.CLEAR, color: 'auto', fill: 'DCE7F2' } : fill ? { type: ShadingType.CLEAR, color: 'auto', fill } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [run(String(t), { size: 18, bold: hd })] })],
  });
  const tr = [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], true)) })];
  rows.forEach(r => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => {
    const fill = opts.color && typeof c === 'string' ? (c.startsWith('✔') ? 'E3F1E4' : c.startsWith('✘') ? 'F8E1E1' : c.startsWith('◐') ? 'FBF0D9' : undefined) : undefined;
    return cell(c, widths[i], false, fill);
  }) })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: tr });
}

const C = [];
// ---------------- copertina
C.push(p(run('Comitato Pari Opportunità – ODCEC di Forlì', { size: 20, color: GREY })));
C.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 120 }, children: [new TextRun({ text: 'Check-up del Bilancio di genere', font: FONT, bold: true, size: 40, color: ACCENT })] }));
C.push(p(run('Analisi dei bilanci di genere di Forlì 2023 e 2024, dei dati della Cassa e dell’Albo 2024-2025, delle Linee guida nazionali e di 17 bilanci di altri Ordini', { size: 22, color: GREY })));
C.push(p(run('Documento di lavoro riservato al CPO – settembre 2026', { size: 20, color: GREY })));

C.push(h2('In sintesi'));
[
  'Il bilancio 2024 di Forlì rispetta circa metà dei contenuti minimi delle Linee guida: mancano metodo e fonti, la serie storica per genere, le cariche, le commissioni, i praticanti per genere e il programma per l’anno successivo.',
  'Il capitolo sul divario di reddito usa solo dati Emilia-Romagna e Italia, pur avendo nella cartella i dati della Cassa specifici per Forlì. Contiene medie calcolate in modo non corretto e testo copiato da un assistente automatico, rimasto anche nel PDF pubblicato.',
  'I dati di Forlì, calcolati qui per la prima volta, sono forti: le donne sono il 45,7% degli iscritti (Italia 34%), ma il divario di reddito è del 54,3%, più alto di Emilia-Romagna (48,2%) e Italia (46,8%), ed è cresciuto di 5 punti in un anno.',
  'Tra i 17 bilanci confrontati nessuno riclassifica le spese dell’Ordine, fissa obiettivi numerici da verificare, rende accessibile il documento o raccoglie dati su DSA e inclusione: su questi temi Forlì può essere la prima.',
].forEach(t => C.push(b(t)));

// ---------------- 1 documenti
C.push(h1('1. Documenti analizzati'));
C.push(table(['Gruppo', 'Documenti'], [
  ['Ordine di Forlì', 'Bilancio di genere 2023 (aprile 2024); Bilancio di genere 2024 (Word e PDF, novembre 2025); estrazione Albo e Registro tirocinio 2024-2025 (file Excel, 1.286 righe)'],
  ['Cassa Dottori Commercialisti', 'Iscritti attivi, reddito netto (Irpef) e volume d’affari (IVA) per genere ed età – Italia, Emilia-Romagna, Ordine di Forlì – comunicazioni 2024 e 2025'],
  ['Consiglio Nazionale', 'Informativa 31/2024 e Linee guida al bilancio di genere (14/3/2024); Bilancio di genere del CNDCEC anno 2023; Questionario CNPO 2025'],
  ['Emilia-Romagna', 'Bologna 2024, Modena 2024, Parma 2025, Reggio Emilia 2025'],
  ['Altri Ordini', 'Venezia 2024, Padova 2024, Genova 2024 (slide), Roma 2025, Firenze preventivo 2026, Arezzo preventivo 2026, Ancona 2025, Lodi 2024, Bari 2025, Trani 2025, Matera 2025, Nola 2022'],
  ['Non analizzati', 'Reputational Report 2025 (oltre 10 MB, non scaricabile dal connettore); documenti sulla certificazione di parità del Comune di Forlì (già noti)'],
], [2300, 6726]));

// ---------------- 2 conformità
C.push(h1('2. Il bilancio 2024 di Forlì rispetto alle Linee guida'));
C.push(p('Legenda: ✔ presente   ◐ parziale   ✘ assente'));
C.push(table(['Contenuto minimo (Linee guida 14/3/2024)', 'Forlì 2024', 'Cosa manca'], [
  ['CPO: composizione, anno di elezione, mandato, obiettivi, funzioni', '✔', '—'],
  ['Attività svolte nell’anno', '◐', 'Solo 2 eventi e le riunioni regionali'],
  ['Attività programmate per l’anno successivo', '✘', 'Presente nel 2023, sparita nel 2024'],
  ['Budget del CPO e, se negato, motivazioni del Consiglio', '◐', 'Dichiarato “nessun budget”, senza le motivazioni del Consiglio'],
  ['Personale dell’Ordine assegnato al CPO', '✘', 'Non indicato'],
  ['Metodo di redazione e fonti in calce a tabelle e grafici', '✘', 'Nessun paragrafo sulle fonti'],
  ['Iscritti sez. A, B, Elenco speciale al 31/12, per genere e generazione', '◐', 'Genere solo sul totale; sezioni e generazioni senza genere'],
  ['Serie storica di almeno 5 anni', '◐', 'Solo totali, non per genere'],
  ['Consiglio per genere e cariche (Presidente, Vice, Segretario)', '◐', 'Solo 4 donne e 7 uomini, senza cariche'],
  ['Confronto con le due consigliature precedenti', '✘', '—'],
  ['Commissioni per genere', '✘', '—'],
  ['Praticanti con la stessa analisi degli iscritti', '◐', 'Solo totali: manca il genere (le donne sono il 60%)'],
  ['Divario di reddito con dati Cassa per genere ed età, rapportati agli iscritti', '✘', 'Usati solo dati regionali e nazionali, non quelli di Forlì'],
  ['Sondaggio (facoltativo)', '✘', 'Nel 2023 citato il sondaggio regionale SAF-ER'],
  ['Analisi facoltative (genitorialità, molestie, Tribunale, CCIAA…)', '✘', '—'],
  ['Conclusioni con la progettualità futura', '◐', 'Conclusioni generiche, senza azioni'],
], [4200, 1300, 3526], { color: true }));

C.push(h2('Errori da non ripetere'));
[
  'Nel capitolo sul divario di reddito è rimasto testo generato da un assistente automatico (“Di seguito trovi medie, osservazioni e una lettura comparativa”, simboli ed emoji nel file Word). Il testo è presente anche nel PDF pubblicato.',
  'Le “medie generali” sono medie semplici delle fasce d’età, senza pesare il numero di iscritti in ciascuna fascia. Vanno usati i totali forniti dalla Cassa.',
  'La frase “il gender gap è più ampio in Emilia-Romagna (+71%)” non corrisponde a un indicatore definito. Il divario va espresso sempre come (media uomini − media donne) / media uomini.',
  'Non è chiaro se la tabella riporti reddito o volume d’affari, né se includa la Cassa Ragionieri. Le Linee guida chiedono il volume d’affari.',
  'La direttiva UE 2023/970 sulla trasparenza retributiva è del 10 maggio 2023 (non del 20) e andava recepita entro il 7 giugno 2026: il passaggio “in Italia non c’è alcun dibattito” va aggiornato. Il dato “gap nel settore privato 20%” non ha fonte.',
  'Piccoli refusi (“cosiglieri”, “obbiettivi”). Il bilancio 2024 è datato 21 novembre 2025: le Linee guida prevedono la presentazione all’assemblea degli iscritti.',
  'Differenza minima tra bilancio ed estrazione: 634 Albo + 13 Elenco speciale nel bilancio, 633 + 14 nel file Excel (totale 647 in entrambi).',
].forEach(t => C.push(b(t)));

// ---------------- 3 dati Forlì
C.push(h1('3. I dati di Forlì, calcolati sui file della cartella'));
C.push(h2('3.1 Iscritti per genere (Albo + Elenco speciale al 31/12)'));
C.push(table(['', '2023', '2024', '2025'], [
  ['Totale', '654', '647', '639'],
  ['Donne', '301 (46,0%)', '299 (46,2%)', '292 (45,7%)'],
  ['Uomini', '353', '348', '347'],
  ['Elenco speciale (di cui donne)', '13', '14 (11)', '11 (8)'],
  ['Età media donne / uomini', '—', '50,3 / 55,5', '51,0 / 55,7'],
], [3000, 2000, 2000, 2026]));
C.push(nota('Fonte: 2023 bilancio di genere 2023; 2024-2025 elaborazione CPO su estrazione Albo. Confronto: donne iscritte Italia 33,9%, Emilia-Romagna 34,2% (CNPO 2025, dati 31/12/2024).'));

C.push(h2('3.2 Donne per classe d’età (2025, classi del questionario nazionale)'));
C.push(table(['Classe d’età', 'Donne', 'Uomini', '% donne'], [
  ['Fino a 35', '28', '25', '52,8%'], ['36-45', '60', '38', '61,2%'], ['46-55', '99', '92', '51,8%'],
  ['56-65', '81', '127', '38,9%'], ['Oltre 65', '24', '65', '27,0%'],
], [3000, 2000, 2000, 2026]));
C.push(p('Sotto i 56 anni le donne sono già maggioranza in tutte le classi. La parità numerica dell’Albo è una questione di anni: il dato giustifica la proiezione demografica proposta nel capitolo 5.'));

C.push(h2('3.3 Praticanti e movimenti 2024-2025'));
C.push(table(['', '2024', '2025'], [
  ['Praticanti totali', '25', '33'], ['di cui donne', '15 (60,0%)', '20 (60,6%)'],
  ['Uscite dall’Albo 2024→2025 (stima)', '—', '17: 9 donne, 8 uomini'],
  ['Ingressi nell’Albo 2024→2025 (stima)', '—', '9: 2 donne, 7 uomini'],
], [4000, 2500, 2526]));
C.push(p('Due segnali da verificare con la Segreteria. Le praticanti sono il 60%, ma tra i nuovi iscritti del 2025 le donne sono solo 2 su 9. Tra le uscite, 4 donne hanno meno di 50 anni, mentre gli uomini usciti sono in gran parte oltre i 57 anni. È il fenomeno segnalato anche dal Consiglio Nazionale: le donne sono passate dal 39% al 46% delle cancellazioni dalla Cassa tra il 2018 e il 2023.'));
C.push(nota('Stima ottenuta confrontando gli elenchi anonimi dei due anni (genere, anno di nascita, sezione). Per il bilancio serve il dato ufficiale di iscrizioni e cancellazioni con il motivo.'));

C.push(h2('3.4 Divario di reddito: Forlì, Emilia-Romagna, Italia'));
C.push(table(['Cassa Dottori – iscritti attivi', 'Forlì', 'Emilia-Romagna', 'Italia'], [
  ['Iscritti attivi (donne %)', '486 (45,3%)', '6.098 (40,8%)', '73.959 (33,4%)'],
  ['Reddito medio uomini', '129.687 €', '133.509 €', '115.046 €'],
  ['Reddito medio donne', '59.218 €', '69.199 €', '61.184 €'],
  ['Divario di reddito', '54,3%', '48,2%', '46,8%'],
  ['Divario di volume d’affari', '57,0%', '52,2%', '50,9%'],
  ['Divario di reddito anno precedente', '49,2%', '47,8%', '46,2%'],
], [3300, 1900, 1900, 1926]));
C.push(nota('Fonte: CNPADC, dati comunicati nel 2025 (redditi 2024) e nel 2024 (redditi 2023). Divario = (media uomini − media donne) / media uomini. Esclusi gli iscritti alla Cassa Ragionieri.'));
C.push(p('A Forlì, in un anno, il reddito medio degli uomini è cresciuto dell’11,9% e quello delle donne dello 0,5%. Il divario è salito di 5 punti ed è oggi più alto della media regionale e nazionale.'));
C.push(table(['Reddito donne in % di quello degli uomini', 'Forlì', 'Emilia-Romagna', 'Italia'], [
  ['Fino a 30 anni (6 uomini, 13 donne)', '133%', '93%', '90%'],
  ['31-40', '73%', '66%', '69%'], ['41-50', '63%', '55%', '57%'],
  ['51-65', '46%', '52%', '52%'], ['Oltre 65 (8 donne)', '49%', '61%', '87%'],
], [3800, 1700, 1700, 1826]));
C.push(p('Lettura: fino a 50 anni le professioniste di Forlì stanno meglio che in regione e in Italia. Il divario complessivo così alto dipende soprattutto dalla fascia 51-65, dove si concentrano gli uomini con i redditi più alti, e dal fatto che le donne iscritte sono più giovani. È un messaggio utile e meno scontato di “le donne guadagnano la metà”.'));

C.push(h2('3.5 Governance'));
[
  'Consiglio 2022-2025: 11 componenti, 4 donne (36%), con iscritte al 46%: 10 punti di sotto-rappresentanza. Mancano le cariche per genere.',
  'Consiglio 2026-2030: presieduto da una donna, la dott.ssa Sara Pennacchi. Composizione e cariche per genere vanno acquisite. A livello nazionale le Presidenti di Ordine sono passate da 19 a 30 su 132.',
  'CPO 2026-2030 (nominato il 23/2/2026): 7 componenti, 5 donne e 2 uomini.',
].forEach(t => C.push(b(t)));

// ---------------- 4 benchmark
C.push(h1('4. Cosa fanno gli altri Ordini'));
C.push(table(['Pratica', 'Chi la fa già', 'Forlì'], [
  ['Incarichi dal Tribunale per genere', 'Reggio Emilia (2020-2025: cancellerie civili, curatori, commissari), Venezia, Parma. Suggerita anche dalle Linee guida', '✘'],
  ['Cariche di sindaco, revisore, amministratore dal Registro imprese', 'Venezia, Padova, Bologna. Suggerita dalle Linee guida', '✘'],
  ['Iscrizioni e cancellazioni per genere', 'Consiglio Nazionale, Modena, Bologna, Roma', '✘'],
  ['Dati Cassa provinciali per età', 'Modena, Venezia, Bologna (anche pensionati), Ancona, Matera', '✘'],
  ['Pensioni e indennità di maternità per genere', 'Genova (slide), Bologna', '✘'],
  ['Relatrici e relatori degli eventi formativi', 'Padova (profilazione eventi); Bari e Firenze (monitoraggio panel)', '✘'],
  ['Questionario proprio', 'Padova, Venezia, Genova (tra più Ordini), Consiglio Nazionale', '◐ (2023: SAF-ER)'],
  ['Consiglio di disciplina e revisori per genere', 'Reggio Emilia, Venezia', '✘'],
  ['Personale dipendente e incarichi conferiti dall’Ordine', 'Roma 2025 (dentro il bilancio di sostenibilità)', '✘'],
  ['Intelligenza artificiale e pari opportunità', 'Bari, Trani', '✘'],
  ['Carta etica, “no woman no panel”', 'Firenze, Reggio Emilia', '✘'],
  ['Sportello di ascolto, Banca del tempo', 'Firenze, Bari, Trani', '✘'],
  ['Evento su pari opportunità e DSA', 'Firenze (19/3/2025)', '✘'],
], [3300, 4626, 1100], { color: true }));

// ---------------- 5 innovazione
C.push(h1('5. Cosa non ha fatto ancora nessuno'));
C.push(p('Verificato su tutti i 17 bilanci della cartella. Questi elementi non compaiono in nessuno:'));
C.push(table(['Proposta per Forlì', 'Perché è nuova'], [
  ['Obiettivi numerici con verifica l’anno dopo (semaforo verde, giallo, rosso)', 'Tutti elencano attività, nessuno misura risultati rispetto a un obiettivo'],
  ['Riclassificazione delle spese dell’Ordine: neutre, sensibili, dirette a ridurre le diseguaglianze', 'È il cuore del “gender budget” nella definizione del Consiglio d’Europa ripresa dalle Linee guida, ma nessuno la applica'],
  ['Bilancio accessibile: sintesi facile da leggere, PDF leggibile dalla sintesi vocale, versione audio', 'Nessun bilancio è pensato per chi ha un DSA; quello di Nola è solo immagini'],
  ['Dati su DSA e inclusione raccolti con il questionario', 'Firenze ha fatto un evento, nessuno ha dati'],
  ['Scomposizione del divario tra età e condizioni reali (il caso Forlì 51-65)', 'Tutti riportano il divario medio; nessuno distingue quanto dipende dall’età degli iscritti'],
  ['Proiezione dell’Albo al 2035 e ricambio generazionale con lente di genere', 'I dati di Forlì mostrano donne in maggioranza sotto i 56 anni'],
  ['Valore economico degli incarichi, non solo il numero', 'Reggio Emilia e Venezia contano gli incarichi, non il loro valore'],
  ['Indicatori ordinati sui 6 ambiti UNI/PdR 125, come pre-valutazione per il progetto regionale', 'Coerente con la proposta “ODCEC Emilia-Romagna verso la UNI/PdR 125”'],
  ['Modello comune per i 10 Ordini dell’Emilia-Romagna', 'I CPO toscani lo hanno proposto al nazionale; in Emilia-Romagna ognuno usa uno schema diverso'],
], [4500, 4526]));

// ---------------- 6 cosa chiedere
C.push(h1('6. Cosa chiedere, a chi'));
C.push(table(['Dato', 'A chi', 'Uso nel bilancio'], [
  ['Iscrizioni e cancellazioni 2021-2025 per genere, età e motivo', 'Segreteria Ordine', 'Flussi e “tubatura che perde”'],
  ['Serie iscritti per genere e sezione 2021-2025', 'Segreteria Ordine', 'Serie storica minima di 5 anni'],
  ['Consiglio, cariche, commissioni, Consiglio di disciplina, revisori: 2017-2021, 2022-2025, 2026-2030', 'Segreteria Ordine', 'Governance e confronto tra consigliature'],
  ['Budget del CPO e personale assegnato', 'Consiglio', 'Paragrafo obbligatorio'],
  ['Eventi formativi 2025 con genere di relatrici e relatori', 'Ordine / Fondazione', 'Panel e “no woman no panel”'],
  ['Dati Cassa Ragionieri per Forlì', 'CNPR', 'Divario di reddito completo'],
  ['Pensioni medie e indennità di maternità e paternità per genere', 'CNPADC, CNPR', 'Genitorialità e pensioni'],
  ['Incarichi 2021-2025 per genere e valore', 'Tribunale di Forlì, OCC', 'Incarichi giudiziali'],
  ['Sindaci, revisori, amministratori iscritti a Forlì, per genere', 'Camera di Commercio della Romagna', 'Cariche societarie'],
  ['Spese dell’Ordine per capitolo', 'Tesoriere', 'Riclassificazione di genere'],
], [3900, 2300, 2826]));

C.push(h2('Prossimi passi'));
[
  'Decidere se il bilancio 2025 correggerà anche i dati 2024 (divario di Forlì, praticanti per genere) in una nota di raccordo.',
  'Inviare le richieste della tabella qui sopra entro ottobre.',
  'Lanciare il questionario breve: i risultati coprono le parti su conciliazione, DSA e inclusione.',
  'Preparare lo schema del bilancio 2025 con le parti obbligatorie e le nuove.',
].forEach(t => C.push(b(t)));

const doc = new Document({
  creator: 'CPO ODCEC Forlì', title: 'Check-up Bilancio di genere ODCEC Forlì',
  background: { color: 'FDFBF5' },
  styles: { default: { document: { run: { font: FONT, size: 22, color: INK } } } },
  numbering: { config: [{ reference: 'punti', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run('Pagina ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[2], buf); console.log('ok'); });
