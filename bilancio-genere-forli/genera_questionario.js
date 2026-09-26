const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow,
  TableCell, WidthType, ShadingType, BorderStyle, LevelFormat, Footer, PageNumber,
} = require('docx');

const FONT = 'Verdana';
const INK = '1F2A37';
const ACCENT = '2E5E8C';
const CNPO = 'CNPO 2025';
const NUOVA = 'Nuova – Forlì-Cesena';
const PNR = 'Preferisco non rispondere';

// ---------- contenuti ----------
const sezioni = [
  {
    titolo: 'Sezione A – Chi sei',
    intro: 'Poche domande per leggere i risultati per gruppi. Le classi sono ampie per non renderti riconoscibile.',
    domande: [
      { id: 'A1', t: 'Genere', src: `${CNPO} – Graf. 1`, k: 's',
        o: ['Donna', 'Uomo', 'Non binario', 'Altro', PNR] },
      { id: 'A2', t: 'Età', src: `${CNPO} – Graf. 2`, k: 's',
        o: ['Meno di 36 anni', 'Tra 36 e 45', 'Tra 46 e 55', 'Tra 56 e 65', 'Oltre 65'] },
      { id: 'A3', t: 'In quale sezione sei iscritta/o?', src: NUOVA, k: 's',
        o: ['Albo – Sezione A', 'Albo – Sezione B', 'Elenco speciale'] },
      { id: 'A4', t: 'Da quanti anni sei iscritta/o all’Albo?', src: NUOVA, k: 's',
        o: ['Meno di 5', 'Da 5 a 10', 'Da 11 a 20', 'Da 21 a 30', 'Oltre 30'] },
      { id: 'A5', t: 'Dove si trova il tuo studio o luogo di lavoro principale?', src: NUOVA, k: 's',
        o: ['Comprensorio forlivese', 'Comprensorio cesenate', 'Rubicone e costa', 'Fuori provincia'] },
      { id: 'A6', t: 'A quale Cassa sei iscritta/o?', src: NUOVA, k: 's',
        o: ['Cassa Dottori Commercialisti', 'Cassa Ragionieri', 'Nessuna / altra'] },
    ],
  },
  {
    titolo: 'Sezione B – Famiglia e cura',
    intro: 'Queste domande sono uguali a quelle del questionario nazionale. Così possiamo confrontare Forlì-Cesena con l’Italia.',
    domande: [
      { id: 'B1', t: 'Hai figli?', src: `${CNPO} – Graf. 4`, k: 's', o: ['Sì', 'No'] },
      { id: 'B2', t: 'Quanti anni ha il figlio o la figlia più piccola?', cond: 'Solo se hai risposto Sì a B1.',
        src: `${CNPO} – Graf. 5`, k: 's', o: ['0-6 anni', '7-13 anni', '14-18 anni', '19-25 anni', 'Oltre 25 anni'] },
      { id: 'B3', t: 'Dopo la nascita di un figlio o una figlia (anche fino a 15 anni dopo), hai avuto la percezione che il tuo reddito o la tua carriera ne abbiano risentito?',
        cond: 'Solo se hai risposto Sì a B1.', src: `${CNPO} – Graf. 6`, k: 's',
        o: ['Sì, il reddito è diminuito',
          'Sì, il reddito non è cresciuto come negli anni precedenti',
          'Sì, la carriera si è fermata (es. non sono diventata/o socia/o)',
          'Sì, ho dovuto rinunciare a incarichi importanti',
          'No',
          'No, anzi il reddito è aumentato (promozione, incarichi importanti)'] },
      { id: 'B4', t: 'Hai genitori o parenti anziani o con disabilità da assistere?', src: `${CNPO} – Graf. 8`, k: 's', o: ['Sì', 'No'] },
      { id: 'B5', t: 'Ti occupi tu direttamente dell’assistenza (di figli o familiari)?', src: `${CNPO} – Graf. 9`, k: 's', o: ['Sì', 'No'] },
      { id: 'B6', t: 'Quante ore al giorno dedichi in media alla cura dei familiari?', src: `${CNPO} – Graf. 10`, k: 's',
        o: ['Fino a 2', 'Da 3 a 4', '5 e oltre'] },
      { id: 'B7', t: 'Se non te ne occupi tu, chi se ne prende cura principalmente?', cond: 'Solo se hai risposto No a B5.',
        src: `${CNPO} – Graf. 11`, k: 's', o: ['Partner', 'Altri familiari', 'Aiuto a pagamento'] },
      { id: 'B8', t: 'Negli ultimi 5 anni hai usato uno di questi strumenti?', src: NUOVA, k: 'm',
        o: ['Indennità di maternità o paternità della Cassa',
          'Esonero dalla formazione continua per maternità, paternità o cura',
          'Riduzione volontaria dell’attività per motivi di cura',
          'Nessuno',
          'Non conoscevo questi strumenti'] },
      { id: 'B9', t: 'Quali iniziative dell’Ordine ti aiuterebbero a conciliare lavoro e cura?', src: NUOVA, k: 'm', max: 3,
        o: ['Eventi formativi in orari compatibili con scuola e cura',
          'Più formazione e-learning registrata',
          'Rete di colleghe e colleghi che si sostituiscono durante maternità, paternità o malattia',
          'Convenzioni con nidi, centri estivi, servizi di assistenza',
          'Sportello informativo su tutele Cassa e agevolazioni',
          'Altro (specificare)'] },
    ],
  },
  {
    titolo: 'Sezione C – Attività professionale',
    intro: 'Servono per capire le cause del divario di reddito tra donne e uomini.',
    domande: [
      { id: 'C1', t: 'In quale fascia è il tuo ultimo reddito professionale dichiarato?', src: `${CNPO} – Graf. 12 e 24`, k: 's',
        o: ['Fino a 15.000 €', 'Da 15.001 a 30.000 €', 'Da 30.001 a 60.000 €', 'Da 60.001 a 85.000 €', 'Da 85.001 a 110.000 €', 'Oltre 110.000 €', PNR] },
      { id: 'C2', t: 'Qual è il tuo ruolo nello studio o nella struttura in cui lavori?', src: `${CNPO} – Graf. 13 (+ 2 voci locali)`, k: 's',
        o: ['Titolare', 'Contitolare', 'Collaborazione presso altro studio', 'Condivisione della struttura con altri colleghi',
          'Socia/o di STP [voce locale]', 'Dipendente di impresa o ente [voce locale]'] },
      { id: 'C3', t: 'Secondo te, perché non sei ancora contitolare?', cond: 'Solo se collabori presso un altro studio.',
        src: `${CNPO} – Graf. 14`, k: 's',
        o: ['Nello studio non investono abbastanza su di me (incarichi, formazione)',
          'Non riesco a dedicare abbastanza ore alla professione',
          'Sono giovane, ma ho possibilità di crescita',
          'Non ho interesse',
          'La struttura dello studio non lo consente',
          'Altro (specificare)'] },
      { id: 'C4', t: 'Quante ore al giorno lavori in media?', src: `${CNPO} – Graf. 15 e 26`, k: 's',
        o: ['Meno di 5', 'Da 5 a 7', '8 ore', '9-10 ore', 'Oltre 10 ore'] },
      { id: 'C5', t: 'In quali settori lavori principalmente?', src: `${CNPO} – Graf. 16 e 29`, k: 'm', max: 3,
        o: ['Contabilità, fisco, paghe, dichiarazioni', 'Consulenza fiscale e societaria', 'Procedure concorsuali',
          'Revisione, sindaco, membro CdA', 'Ristrutturazione aziendale, advisor', 'Procedure di sovraindebitamento', 'Finanza aziendale'] },
      { id: 'C6', t: 'Quando prepari un preventivo, come ti comporti?', src: `${CNPO} – Graf. 17`, k: 's',
        o: ['Chiedo il compenso minimo, sotto il quale non lavoro',
          'Chiedo un compenso alto, per avere margine di trattativa',
          'Chiedo un compenso giusto, non trattabile e modificabile se c’è lavoro extra'] },
      { id: 'C7', t: 'Durante un incarico, in quali comportamenti ti riconosci?', src: `${CNPO} – Graf. 18 e 30`, k: 'm',
        o: ['Faccio lavoro extra senza chiedere un compenso aggiuntivo',
          'Faccio lavoro extra solo dopo aver concordato un compenso aggiuntivo',
          'Faccio fatica a rifiutare lavori pagati male o quando sono sovraccarica/o',
          'Non faccio fatica a rifiutare lavori pagati male',
          'Sono in difficoltà negli ambiti nuovi perché non mi sento all’altezza',
          'Mi piace mettermi in gioco in ambiti nuovi'] },
      { id: 'C8', t: 'Negli ultimi 3 anni hai ricevuto incarichi di nomina pubblica o giudiziale?', src: NUOVA, k: 'm',
        o: ['Curatore/curatrice o commissario/a giudiziale', 'Gestore della crisi (OCC)', 'Revisore di enti locali',
          'CTU o CTP', 'Custode o delegato/a alle vendite', 'Nessuno, ma sono iscritta/o agli elenchi', 'Nessuno e non sono iscritta/o agli elenchi'] },
      { id: 'C9', t: 'Quanto usi strumenti digitali avanzati o di intelligenza artificiale nel tuo lavoro?', src: NUOVA, k: 's',
        o: ['Mai', 'Raramente', 'Qualche volta', 'Spesso', 'Ogni giorno'] },
    ],
  },
  {
    titolo: 'Sezione D – Conciliazione, pari opportunità, violenza economica',
    domande: [
      { id: 'D1', t: 'Ritieni adeguato il tuo equilibrio tra vita privata e lavoro?', src: `${CNPO} – Graf. 19`, k: 's',
        o: ['Assolutamente no', 'No', 'Abbastanza', 'Sì, abbastanza', 'Assolutamente sì'] },
      { id: 'D2', t: 'Nel tuo lavoro, ti è capitato di vedere queste situazioni?', src: `${CNPO} – Graf. 20`, k: 'm',
        o: ['Coniuge o partner escluso dalla gestione dei conti o del patrimonio comune',
          'Socia o socio titolare di quote, ma escluso dalle decisioni',
          'Clienti in dipendenza economica, senza accesso alle proprie risorse',
          'Donazioni o cessioni simulate per sottrarre beni a un familiare o socio più debole',
          'Redditi nascosti per ridurre il mantenimento o impedire l’indipendenza del coniuge',
          'Altro (specificare)',
          'Mai visto / non saprei riconoscerle'] },
      { id: 'D3', t: 'Sai che questi comportamenti sono forme di violenza economica?', src: `${CNPO} – Graf. 21`, k: 's',
        o: ['Sì', 'Sì, e vorrei formazione su questo tema', 'No', 'No, e vorrei formazione su questo tema'] },
      { id: 'D4', t: 'Cosa significa per te “pari opportunità” nella nostra professione?', src: `${CNPO} – Graf. 22`, k: 'm',
        o: ['Pari opportunità di lavoro e di incarichi', 'Pari compenso', 'Stesso trattamento indipendentemente da genere ed età',
          'Pari dignità professionale', 'Pari tutele e servizi di assistenza'] },
      { id: 'D5', t: 'Ritieni utile la formazione sulle pari opportunità organizzata dagli Ordini e dal Consiglio Nazionale?', src: `${CNPO} – Graf. 23`, k: 's',
        o: ['Sì', 'Sì, soprattutto per cambiare abitudini che penalizzano il mio reddito', 'Abbastanza', 'No', 'Non so, non la seguo'] },
      { id: 'D6', t: 'Negli ultimi 3 anni, nel lavoro, hai subito o visto comportamenti discriminatori o molestie legati al genere?', src: NUOVA, k: 'm',
        o: ['Sì, li ho subiti', 'Sì, li ho visti', 'No', PNR] },
      { id: 'D7', t: 'Se sì, da parte di chi?', cond: 'Solo se hai risposto Sì a D6.', src: NUOVA, k: 'm',
        o: ['Clienti', 'Colleghe o colleghi', 'Titolari o soci dello studio', 'Uffici pubblici o tribunali', 'Altro (specificare)'] },
    ],
  },
  {
    titolo: 'Sezione E – L’Ordine e la formazione',
    domande: [
      { id: 'E1', t: 'Come preferisci seguire la formazione continua?', src: NUOVA, k: 's',
        o: ['In presenza', 'Online in diretta', 'E-learning registrato', 'Misto'] },
      { id: 'E2', t: 'Quale fascia oraria ti è più comoda per gli eventi in diretta?', src: NUOVA, k: 'm',
        o: ['Mattina 9-13', 'Primo pomeriggio 14-16', 'Tardo pomeriggio 16-19', 'Sera dopo le 19'] },
      { id: 'E3', t: 'Conosci il Comitato Pari Opportunità dell’Ordine di Forlì?', src: NUOVA, k: 's',
        o: ['Sì, e ho partecipato a sue iniziative', 'Sì, ma non ho partecipato', 'No'] },
      { id: 'E4', t: 'Partecipi o hai partecipato a una commissione di studio dell’Ordine?', src: NUOVA, k: 's',
        o: ['Sì, come presidente o coordinatrice/coordinatore', 'Sì, come componente', 'No'] },
      { id: 'E5', t: 'Hai mai pensato di candidarti al Consiglio dell’Ordine o ad altre cariche? Se no, perché?', src: NUOVA, k: 'm',
        o: ['Mi sono candidata/o', 'Non ho tempo', 'Non so come funziona', 'Non mi sento rappresentata/o',
          'L’ambiente mi sembra poco accogliente', 'Non mi interessa'] },
    ],
  },
  {
    titolo: 'Sezione F – Accessibilità e DSA (facoltativa)',
    intro: 'Questa sezione è facoltativa. Alcune domande riguardano la salute: sono “dati particolari” (art. 9 GDPR). Rispondi solo se vuoi. Le risposte restano anonime e sono pubblicate solo in forma aggregata.',
    domande: [
      { id: 'F0', t: 'Vuoi rispondere a questa sezione?', src: 'Consenso esplicito', k: 's',
        o: ['Sì, acconsento al trattamento dei dati di questa sezione', 'No, passo alla sezione G'] },
      { id: 'F1', t: 'Quali strumenti renderebbero più facili da usare le comunicazioni e la formazione dell’Ordine?', src: NUOVA, k: 'm',
        o: ['Sintesi brevi all’inizio di circolari e news', 'Versione audio dei documenti', 'Sottotitoli nei video e nei webinar',
          'Slide scaricabili prima dell’evento', 'Caratteri più leggibili e testi meno densi', 'Test finali FPC senza limite di tempo',
          'Mappe concettuali e schemi', 'Nessuno, va bene così'] },
      { id: 'F2', t: 'Hai una diagnosi di DSA (dislessia, disortografia, disgrafia, discalculia)?', src: NUOVA, k: 's',
        o: ['Sì', 'Non ho una diagnosi, ma penso di sì', 'No', PNR] },
      { id: 'F3', t: 'Hai altre condizioni (es. ADHD, disabilità visiva o uditiva) che incidono su come studi o lavori?', src: NUOVA, k: 's',
        o: ['Sì', 'No', PNR] },
      { id: 'F4', t: 'All’esame di Stato o durante il tirocinio hai chiesto strumenti compensativi (tempo aggiuntivo, calcolatrice, sintesi vocale…)?', src: NUOVA, k: 's',
        o: ['Sì, e li ho ottenuti', 'Sì, ma non li ho ottenuti', 'No, non sapevo di poterli chiedere', 'No, non ne avevo bisogno'] },
      { id: 'F5', t: 'Quali attività della professione ti risultano più faticose?', src: NUOVA, k: 'm', max: 3,
        o: ['Leggere normativa e circolari lunghe', 'Scrivere relazioni e pareri', 'Calcoli e controlli numerici',
          'Rispettare molte scadenze insieme', 'Test a tempo della formazione online', 'Nessuna in particolare'] },
      { id: 'F6', t: 'Hai mai parlato della tua condizione in ambito professionale?', src: NUOVA, k: 's',
        o: ['Sì, con colleghe/i o nello studio', 'Sì, anche con clienti', 'No, per timore di pregiudizi', 'No, non mi sembrava necessario', 'Non applicabile'] },
    ],
  },
  {
    titolo: 'Sezione G – La tua proposta',
    domande: [
      { id: 'G1', t: 'Se potessi chiedere una sola cosa all’Ordine per le pari opportunità e l’inclusione, cosa chiederesti?', src: NUOVA, k: 'o' },
    ],
  },
];

// ---------- helpers ----------
const run = (text, opts = {}) => new TextRun({ text, font: FONT, color: INK, size: 24, ...opts });
const p = (children, opts = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children], spacing: { after: 120, line: 360 }, ...opts });
const bullet = (text) => new Paragraph({ numbering: { reference: 'punti', level: 0 }, spacing: { after: 80, line: 360 }, children: [run(text)] });

function domanda(d) {
  const out = [];
  out.push(new Paragraph({
    keepNext: true, spacing: { before: 280, after: 80, line: 360 },
    children: [run(`${d.id}. `, { bold: true, color: ACCENT }), run(d.t, { bold: true })],
  }));
  const tipo = d.k === 's' ? 'Una sola risposta' : d.k === 'm' ? (d.max ? `Massimo ${d.max} risposte` : 'Più risposte possibili') : 'Risposta libera';
  const meta = [run(`${tipo}`, { size: 20, color: '4B5563' })];
  if (d.cond) meta.push(run(`  ·  ${d.cond}`, { size: 20, color: '4B5563' }));
  out.push(new Paragraph({ keepNext: true, spacing: { after: 80, line: 300 }, children: meta }));
  const box = d.k === 's' ? '○' : '☐';
  (d.o || []).forEach((o, i, a) => out.push(new Paragraph({
    keepNext: i < a.length - 1, indent: { left: 360 }, spacing: { after: 40, line: 340 },
    children: [run(`${box}  ${o}`)],
  })));
  if (d.k === 'o') {
    for (let i = 0; i < 3; i++) out.push(new Paragraph({
      spacing: { before: 200, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '9CA3AF', space: 1 } }, children: [run('')],
    }));
  }
  out.push(new Paragraph({ spacing: { before: 40, after: 0 }, children: [run(`Fonte: ${d.src}`, { size: 16, color: '6B7280' })] }));
  return out;
}

const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 160 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 32, color: ACCENT })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, pageBreakBefore: true, spacing: { before: 0, after: 160 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 28, color: ACCENT })] });

// Tabella di raccordo con il questionario nazionale
function tabellaRaccordo() {
  const W = [1500, 4526, 3000];
  const cell = (text, w, head = false) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, color: 'auto', fill: 'DCE7F2' } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [run(text, { size: 20, bold: head })] })],
  });
  const rows = [new TableRow({ tableHeader: true, children: [cell('Domanda', W[0], true), cell('Tema', W[1], true), cell('Confronto con', W[2], true)] })];
  sezioni.flatMap(s => s.domande).forEach(d => {
    rows.push(new TableRow({ cantSplit: true, children: [cell(d.id, W[0]), cell(d.t.length > 70 ? d.t.slice(0, 67) + '…' : d.t, W[1]), cell(d.src, W[2])] }));
  });
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: W, rows });
}

// ---------- documento ----------
const children = [
  new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: 'Ordine dei Dottori Commercialisti e degli Esperti Contabili di Forlì', font: FONT, size: 22, color: '4B5563' })] }),
  new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 120 }, children: [new TextRun({ text: 'Questionario per il Bilancio di genere e inclusione', font: FONT, bold: true, size: 40, color: ACCENT })] }),
  p(run('Bozza per il Comitato Pari Opportunità – da validare prima della diffusione', { size: 22, color: '4B5563' })),

  h1('Prima di iniziare'),
  p(run('Cara collega, caro collega,')),
  p(run('l’Ordine sta preparando il suo Bilancio di genere e inclusione. Ti chiediamo circa 12 minuti.')),
  bullet('Il questionario è anonimo: non chiediamo nome, email o codice fiscale.'),
  bullet('Ogni domanda è facoltativa. Puoi sempre saltarla.'),
  bullet('Molte domande sono uguali a quelle del questionario nazionale CNPO 2025. Così possiamo confrontare il nostro territorio con l’Italia.'),
  bullet('Pubblicheremo solo dati aggregati. Non mostreremo gruppi con meno di 5 persone.'),
  bullet('Non c’è limite di tempo. Puoi usare la sintesi vocale del tuo dispositivo.'),
  p([run('Legenda: ', { bold: true }), run('○ una sola risposta   ☐ più risposte possibili')]),
];

sezioni.forEach((s, i) => {
  children.push(h2(s.titolo));
  if (s.intro) children.push(p(run(s.intro, { color: '374151' })));
  s.domande.forEach(d => children.push(...domanda(d)));
});

children.push(h2('Grazie!'));
children.push(p(run('Grazie per il tuo tempo. I risultati saranno presentati nel Bilancio di genere e inclusione dell’Ordine, anche in versione breve e facile da leggere.')));

// ---------- note per il CPO ----------
children.push(h2('Note per il Comitato Pari Opportunità (da non pubblicare)'));
children.push(h1('Privacy'));
[
  'Informativa ex art. 13 GDPR in apertura: titolare (Ordine), finalità (Bilancio di genere), base giuridica, conservazione, diritti.',
  'Sezione F: dati relativi alla salute (art. 9 GDPR). Serve consenso esplicito (domanda F0) e la sezione deve restare facoltativa.',
  'Nella piattaforma: disattivare raccolta email, login e indirizzi IP. Preferire server nell’Unione europea.',
  'Nelle tabelle pubblicate: oscurare le celle con meno di 5 risposte. Non incrociare più di due variabili del profilo (es. genere × età, non genere × età × zona).',
  'Sentire il DPO dell’Ordine prima della diffusione.',
].forEach(t => children.push(bullet(t)));

children.push(h1('Accessibilità del questionario'));
[
  'Una domanda per schermata, con barra di avanzamento e possibilità di salvare e riprendere.',
  'Carattere senza grazie da almeno 12 punti, testo allineato a sinistra, niente corsivi o maiuscole lunghe.',
  'Nessun limite di tempo. Verificare che la piattaforma funzioni con la sintesi vocale.',
  'Offrire un recapito (telefono o email del CPO) per chi preferisce rispondere con assistenza.',
].forEach(t => children.push(bullet(t)));

children.push(h1('Diffusione e analisi'));
[
  'Periodo suggerito: 4-6 settimane, con due promemoria.',
  'Obiettivo: almeno il 20% degli iscritti. Il campione nazionale CNPO in Emilia-Romagna è stato molto piccolo (0,8% degli uomini, 2,8% delle donne).',
  'Tenere traccia della quota di donne e uomini tra chi risponde. Se è molto diversa da quella dell’Albo, pesare i risultati per genere ed età.',
  'Per le domande con fonte “CNPO 2025” usare le stesse classi del documento nazionale, così il confronto è diretto.',
].forEach(t => children.push(bullet(t)));

children.push(h1('Raccordo con il questionario nazionale'));
children.push(tabellaRaccordo());

const doc = new Document({
  creator: 'Bilancio di genere ODCEC Forlì',
  title: 'Questionario Bilancio di genere e inclusione – ODCEC Forlì',
  background: { color: 'FDFBF5' },
  styles: { default: { document: { run: { font: FONT, size: 24, color: INK } } } },
  numbering: { config: [{ reference: 'punti', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run('Pagina ', { size: 18, color: '6B7280' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: '6B7280' })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(process.argv[2], b);
  console.log('scritto', process.argv[2], sezioni.flatMap(s => s.domande).length, 'domande');
});
