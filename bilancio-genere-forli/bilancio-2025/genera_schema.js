// Punto della situazione e schema del Bilancio di genere 2025 (documento di lavoro CPO, 17/11/2026)
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, LevelFormat, PageOrientation } = require('docx');

const FONT = 'Verdana', INK = '1F2A37', ACCENT = '2E5E8C', GREY = '4B5563';
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, color: INK, size: 19, ...o });
const p = (c, o = {}) => new Paragraph({ children: (Array.isArray(c) ? c : [c]).map(x => typeof x === 'string' ? run(x) : x), spacing: { after: 120, line: 300 }, ...o });
const b = (c) => new Paragraph({ numbering: { reference: 'punti', level: 0 }, spacing: { after: 60, line: 290 }, children: (Array.isArray(c) ? c : [c]).map(x => typeof x === 'string' ? run(x) : x) });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 140 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 28, color: ACCENT })] });
const FILL = { '✔': 'DDF0DD', '◐': 'FFF4C2', '✘': 'F8D7D7' };

function table(head, rows, widths) {
  const cell = (t, w, hd) => {
    const s = String(t);
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: hd ? { type: ShadingType.CLEAR, color: 'auto', fill: 'DCE7F2' } : (FILL[s] ? { type: ShadingType.CLEAR, color: 'auto', fill: FILL[s] } : undefined),
      margins: { top: 50, bottom: 50, left: 90, right: 90 },
      children: s.split('\n').map(l => new Paragraph({ children: [run(l, { size: 17, bold: hd })] })),
    });
  };
  const tr = [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], true)) })];
  rows.forEach(r => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false)) })));
  return new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths, rows: tr });
}

const C = [];
C.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 100 }, children: [new TextRun({ text: 'Bilancio di genere 2025 – punto della situazione', font: FONT, bold: true, size: 36, color: ACCENT })] }));
C.push(p(run('Documento di lavoro del CPO, 17 novembre 2026. Da qui all’assemblea del 26/11 non si chiedono altri dati all’esterno: si lavora su quelli che abbiamo.', { color: GREY })));

C.push(h1('In sintesi'));
C.push(b([run('Dati: ', { bold: true }), run('abbiamo 9 fonti su 10. Manca solo il consuntivo 2025 della Tesoriera; il Tribunale ha negato i dati.')]));
C.push(b([run('Bozza: ', { bold: true }), run('13 capitoli scritti, 10 parti ancora in giallo. 5 si chiudono con materiale che abbiamo già, 3 vanno riformulate, 2 sono le firme e una tua conferma.')]));
C.push(b([run('Rispetto ai 17 bilanci di altri Ordini esaminati, ', { bold: true }), run('Forlì ha già 6 contenuti che nessuno o quasi ha: OCC, Camera di Commercio, relatrici, obiettivi numerici, spese con la lente di genere, confronto nazionale CNPO.')]));

C.push(h1('1. Schema del bilancio e stato dei dati'));
C.push(p(run('Legenda: ✔ completo   ◐ parziale, si chiude con quello che abbiamo   ✘ manca', { color: GREY })));
const W = [2300, 3500, 700, 3600];
C.push(table(['Capitolo', 'Dati che abbiamo (fonte)', 'Stato', 'Cosa manca / cosa possiamo aggiungere'], [
  ['Sintesi: il 2025 in sei numeri', 'Iscritti, divario Cassa, Consiglio, relatrici, OCC, CCIAA', '✔', 'Aggiornare con budget CPO 1.500 € (primo stanziamento)'],
  ['1. Premessa e metodo', 'Fonti, regole di calcolo, raccordo con bilancio 2024', '◐', 'Riga sulle spese: senza consuntivo si cita il preventivo 2026'],
  ['2. Quadro normativo', 'Linee guida CN 2024, Regolamento CPO 2023, d.lgs. 139/2005', '◐', 'Direttiva UE 2023/970: recepimento da verificare sulla Gazzetta Ufficiale (lo controllo io)'],
  ['3. Il CPO', 'Composizione e ruoli, nomina 23/2/2026, budget 1.500 €', '◐', 'Ufficio di supporto: Segreteria (Sara Maraldi, Monica Valmori), dalle mail\nConferma genere di un componente (file Segreteria)'],
  ['4. Attività 2025', '—', '✘', 'Cerco nel Drive e in Gmail le attività del CPO 2022-2026 nel 2025 (bilancio 2024, riunioni regionali, informativa 88/2025)'],
  ['4. Attività 2026', 'Verbali 1-2, convegno 19/6 (con Consulenti del Lavoro, 100 posti, crediti CN, SEAC 500 €), CPO ER e voto 23/9, coordinamento 16 e 24/11, rete con Ravenna', '◐', 'Da scrivere ora: tutto il materiale è nel fascicolo'],
  ['5. Iscritte e iscritti', 'Albo e praticanti 2021-2025, età, flussi, cancellazioni (Segreteria); confronto Italia/ER (CNPO)', '✔', 'Aggiungere: proiezione “anno della parità” dell’Albo, già calcolabile dall’estrazione'],
  ['6. Chi decide', 'Consiglio 2022-26 e 2026-30, Disciplina, Revisori, commissioni, CPO', '✔', 'Aggiungere: candidature alle elezioni 2026 per genere (elenco candidati)'],
  ['7. Divario di reddito', 'Cassa Dottori 2024 e 2025 per età; Cassa Ragionieri 2025', '✔', '—'],
  ['8. Formazione e conciliazione', '51 eventi e 267 interventi; esoneri 2025', '✔', 'Correggere il doppio conteggio aula/webinar'],
  ['9. Incarichi e cariche', 'OCC 2021-2025; CCIAA 2021-2025 e imprese attive 2025; nota sul diniego del Tribunale', '✔', 'OCC 2022 da segnare come dato non confermato'],
  ['10. Spese con lente di genere', 'Fondazione 2025 (bilancio e scheda eventi); budget CPO 2026', '◐', 'Senza consuntivo: metodo + Fondazione + preventivo 2026; tabella completa diventa obiettivo 2026'],
  ['11. Questionario', 'Questionario breve pronto; risultati nazionali CNPO 2025', '◐', 'Se il nostro non è partito: presentarlo come azione 2026-27 e tenere il confronto nazionale'],
  ['12. Obiettivi 2027', '8 obiettivi con indicatore, partenza e traguardo', '◐', 'Partenza del questionario: “non rilevato” se non è partito'],
  ['13. Conclusioni', 'Testo presente', '◐', 'Data e firme del CPO'],
], W));

C.push(h1('2. Le 10 parti in giallo e come chiuderle'));
C.push(table(['Parte', 'Come si chiude', 'Chi'], [
  ['Consuntivo spese 2025 (2 punti)', 'Si scrive “il consuntivo per capitolo sarà analizzato nel bilancio 2026”; resta Fondazione + budget CPO', 'Claude'],
  ['Direttiva UE 2023/970', 'Verifica del recepimento sulle fonti ufficiali; frase datata “alla data del bilancio”', 'Claude'],
  ['Ufficio di supporto al CPO', 'Segreteria dell’Ordine', 'Claude, con tua conferma'],
  ['Attività CPO 2025', 'Ricerca in Drive/Gmail, testo breve', 'Claude'],
  ['Questionario: invio e risposte (3 punti)', 'Dipende da te: è stato inviato agli iscritti?', 'Emanuela'],
  ['Altre attività 2026', 'Dal fascicolo (verbali, mail, Drive)', 'Claude'],
  ['Data e firme', 'Alla chiusura, prima del 26/11', 'CPO'],
], [3000, 5300, 1800]));

C.push(h1('3. Cosa possiamo aggiungere con i dati che abbiamo'));
C.push(p('In ordine di utilità per l’assemblea. Le prime quattro richiedono poco lavoro.'));
C.push(table(['Proposta', 'Dati', 'Perché', 'Altri Ordini'], [
  ['1. Attività 2026 complete, con convegno e rete', 'Verbali, mail, Drive', 'È il primo anno del nuovo CPO: va raccontato', 'Tutti'],
  ['2. Candidature alle elezioni 2026 per genere', 'Elenco candidati (gennaio 2026)', 'Mostra quante donne si candidano, non solo quante sono elette', 'Nessuno dei 17'],
  ['3. In quanti anni l’Albo arriva alla parità', 'Estrazione Albo per età', 'Sotto i 56 anni le donne sono già maggioranza: messaggio forte e positivo', 'Nessuno dei 17'],
  ['4. Riquadro “Forlì e l’Italia”', 'CNPO 2025, Cassa', 'Iscritte 46% contro 34%, ma divario di reddito più alto della media', 'Pochi'],
  ['5. Sintesi “facile da leggere” di 2 pagine', 'Bozza', 'Già tra gli obiettivi (accessibilità); serve anche per l’assemblea', 'Nessuno'],
  ['6. Slide per l’assemblea del 26/11', 'Bozza e grafici', '10-12 slide con i numeri chiave', 'Padova, Genova, Brescia'],
], [2700, 2100, 3400, 1900]));

C.push(h1('4. Controlli prima della chiusura'));
C.push(b('Genere di un componente del CPO: nel file della Segreteria il conteggio dà 6 donne e 1 uomo, il check-up di settembre ne contava 5 e 2.'));
C.push(b('OCC 2022: nel file compilato la cella dell’anno contiene “53”; il dato 53/88 va presentato come non confermato.'));
C.push(b('Praticanti 2023: 28 nel bilancio 2024, 29 nell’estrazione.'));
C.push(b('Cancellazioni 2023: 8 senza motivo.'));
C.push(b('Eventi: chi è in aula e in webinar conta due volte (51 righe).'));
C.push(b('Nota della Segreteria sulle regole elettorali: verificarla sul Regolamento CN prima di citarla.'));
C.push(b('Camera di Commercio: avvisarla prima della pubblicazione, come richiesto.'));
C.push(b('Privacy: nel bilancio solo dati aggregati; classi sotto 5 persone non pubblicate.'));

C.push(h1('5. Cosa mi serve da te'));
C.push(b('Il questionario breve è stato inviato agli iscritti? Se sì, quando e con quante risposte.'));
C.push(b('Via libera alle proposte 1-4 della sezione 3 (e se vuoi anche 5 e 6).'));
C.push(b('Conferma: Segreteria come ufficio di supporto al CPO.'));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 19 } } } },
  numbering: { config: [{ reference: 'punti', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 900, right: 900 } } }, children: C }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(path.join(__dirname, process.argv[2] || 'Punto_situazione_Bilancio_genere_2025.docx'), buf); console.log('ok'); });
