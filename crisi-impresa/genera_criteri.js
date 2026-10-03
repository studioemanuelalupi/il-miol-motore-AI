const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, Footer, PageNumber, AlignmentType,
} = require('docx');

const FONT = 'Verdana', INK = '1F2A37', ACCENT = '2E5E8C', GREY = '4B5563';
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, color: INK, size: 22, ...o });
const p = (c, o = {}) => new Paragraph({ children: Array.isArray(c) ? c : [typeof c === 'string' ? run(c) : c], spacing: { after: 120, line: 340 }, ...o });
// punto elenco: "Titolo: testo" mette in grassetto la parte prima dei due punti
const b = (t) => {
  const i = t.indexOf(': ');
  const kids = i > 0 && i < 60 ? [run(t.slice(0, i + 1), { bold: true }), run(t.slice(i + 1))] : [run(t)];
  return new Paragraph({ numbering: { reference: 'punti', level: 0 }, spacing: { after: 60, line: 320 }, children: kids });
};
const h1 = (t, brk = true) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: brk, spacing: { after: 160 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 30, color: ACCENT })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 280, after: 120 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: ACCENT })] });
const nota = (t) => p(run(t, { size: 18, color: GREY }));
const list = (arr) => arr.forEach(t => C.push(b(t)));

function table(head, rows, widths) {
  const total = widths.reduce((a, c) => a + c, 0);
  const cell = (t, w, hd) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: hd ? { type: ShadingType.CLEAR, color: 'auto', fill: 'DCE7F2' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [run(String(t), { size: 18, bold: hd })] })],
  });
  const tr = [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], true)) })];
  rows.forEach(r => tr.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], false)) })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: tr });
}

const C = [];
// ---------------- copertina
C.push(p(run('Materiale di studio interno – OCC / Tribunale di Forlì', { size: 20, color: GREY })));
C.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { after: 120 }, children: [new TextRun({ text: 'Sovraindebitamento: i criteri dei giudici', font: FONT, bold: true, size: 40, color: ACCENT })] }));
C.push(p(run('Riassunto per procedura di 46 provvedimenti 2023-2026 (Forlì, Genova, Verona, Savona, Benevento)', { size: 22, color: GREY })));
C.push(p(run('Ottobre 2026 – schede anonime: nessun dato dei debitori, dei creditori persone fisiche o dei professionisti', { size: 20, color: GREY })));

C.push(h2('In sintesi'));
list([
  'Ristrutturazione dei debiti del consumatore: il giudice controlla sempre ammissibilità e fattibilità. La convenienza la controlla solo se un creditore la contesta. A Forlì il confronto con la liquidazione controllata si fa con 36 mesi di quota di stipendio, più i beni al netto dei costi di vendita, meno il compenso del liquidatore.',
  'Concordato minore: se non c’è prosecuzione dell’attività serve finanza esterna vera, che aumenti l’attivo in modo apprezzabile rispetto alla liquidazione (compresi 3 anni di reddito). Se il voto del Fisco è determinante e la proposta gli conviene, il giudice omologa lo stesso.',
  'Liquidazione controllata: il piano del debitore non vincola. Tutto l’attivo va al liquidatore, anche quello che arriva dopo, fino all’esdebitazione. Il giudice fissa la parte di reddito che resta al debitore: sono considerate solo le spese documentate e la quota non scende sotto il quinto pignorabile.',
  'Esdebitazione: non serve una percentuale minima ai creditori. È negata in caso di condanne ostative senza riabilitazione, distrazioni, assenza di contabilità o aggravamento del dissesto.',
  'Merito creditizio (art. 124-bis TUB): il finanziatore che non ha verificato il merito creditizio non può opporsi, e la sua colpa “assorbe” quella del consumatore.',
  'Ludopatia: non è colpa grave se è una vera patologia certificata dal SerD con percorso di cura.',
]);

// ---------------- 0 fonti
C.push(h1('1. Fonti e metodo'));
C.push(table(['Tribunale', 'Provvedimenti letti', 'Tipo'], [
  ['Forlì', '11', 'Omologhe RDC, aperture di liquidazione controllata (anche familiare), decreti di esdebitazione'],
  ['Genova', '21', 'Omologhe RDC e concordato minore, decreti sui limiti di reddito in liquidazione, esdebitazioni concesse e negate, chiusure'],
  ['Verona', '9', 'Aperture e omologhe di concordato minore, RDC (anche un rigetto), decreto di correzione'],
  ['Savona', '3', 'Omologa RDC, due aperture di liquidazione controllata'],
  ['Benevento', '2', 'Ammissione RDC familiare, concordato minore'],
], [1600, 1700, 5726]));
nota('I provvedimenti sono nella cartella Drive “raccolta sentenze”. Altri 2 file sono scansioni senza testo (nomine di liquidatore) e non sono stati usati. Le citazioni indicano solo tribunale e anno. I provvedimenti originali contengono dati personali: per citarli in atti usare la massima anonima, non la copia integrale.');

// ---------------- 2 criteri comuni
C.push(h1('2. Criteri comuni a tutte le procedure'));
C.push(h2('Accesso'));
list([
  'Sovraindebitamento (art. 2 lett. c): confronto tra debiti, patrimonio prontamente liquidabile e reddito al netto del mantenimento. Una disponibilità mensile insufficiente a pagare le rate basta a dimostrarlo.',
  'Consumatore (art. 2 lett. e): conta la natura dei debiti, non la professione attuale. Il dipendente con debiti personali è consumatore. Chi ha debiti da fideiussioni per società o da socio illimitatamente responsabile non lo è e deve usare il concordato minore o la liquidazione.',
  'Ex socio di società fallita e cancellata: può accedere al concordato minore. Il divieto dell’art. 33 c.4 non si applica perché l’impresa era della società, non sua (Verona 2026).',
  'Precedenti: nessuna esdebitazione nei 5 anni precedenti e mai due volte; nessuna procedura revocata per sua colpa (art. 72); nessuna domanda pendente del titolo IV.',
  'Competenza (art. 27 c.2 e 28): residenza del debitore. Un trasferimento nell’anno precedente non sposta la competenza (Forlì).',
]);
C.push(h2('Colpa grave, malafede, frode e merito creditizio'));
list([
  'Valutazione: si guardano le cause del debito ricostruite dall’OCC. Contano crisi della società partecipata, malattia, separazione con mantenimento non versato, aumento dei tassi, Covid. Tutte queste sono cause esterne o al più colpa lieve.',
  'Silenzio reticente: non dichiarare altri debiti nel modulo di finanziamento non è frode, perché la frode richiede artifici. La colpa del consumatore è assorbita da quella della finanziaria che poteva consultare le banche dati in pochi secondi (Genova, conforme a Messina 2023).',
  'Art. 69 c.2: il creditore che ha violato l’art. 124-bis TUB non può opporsi né contestare la convenienza o le spese di mantenimento (Genova 2024). L’OCC deve sempre fare la tabella rata/reddito alla data di ogni finanziamento.',
  'Mutuo prima casa preso con un familiare che poi non paga, e banca che ha concesso il credito: colpa lieve (Verona 2026).',
  'Dolo: è ostativo sospendere volontariamente il mutuo contestandone la validità con argomenti artificiosi (Verona 2026, rigetto).',
  'Concordato minore: rileva solo la frode (art. 77), non la colpa grave. I debiti per imposte dichiarate e non versate non impediscono l’accesso (Genova 2025).',
]);
C.push(h2('Ludopatia'));
list([
  'Regola: non è colpa grave se è un disturbo patologico (non semplice abitudine al gioco), documentato e seguito dal SerD (Savona 2026, che richiama Ravenna 2021, Catania 2020 e Oristano 2023; conforme Forlì).',
  'Prova: certificato SerD con percorso iniziato o concluso, e misure familiari di controllo (carte tolte, denaro limitato).',
  'Ricaduta: una ricaduta episodica che il debitore ha denunciato da sé non è ostativa (Forlì).',
  'Prescrizioni: Forlì impone la certificazione SerD in ogni relazione semestrale. Se il percorso si interrompe la procedura può essere revocata (art. 72). Savona aggiunge il divieto di nuove carte e di accesso al credito per tutta la durata del piano, con comunicazione alla Centrale Rischi.',
]);
C.push(h2('Spese di mantenimento e quota per i creditori'));
list([
  'Prova: valgono solo le spese documentate. Ad esempio un contributo a un figlio maggiorenne non documentato viene escluso (Forlì). Le spese si possono verificare con i dati ISTAT, e Genova applica un margine prudenziale di circa il 30%.',
  'Reddito del coniuge: non è acquisibile, ma serve a valutare il tenore di vita della famiglia. Le spese non vanno divise 50/50 se un coniuge guadagna molto di più (Genova 2024, Verona).',
  'Soglia minima: in liquidazione ai creditori non può andare meno del quinto pignorabile (Genova 2024).',
  'Tredicesima e quattordicesima: Genova le acquisisce al 50%.',
  'Entrate future: il debitore deve versare ogni entrata oltre una soglia annua fissata dal giudice (Genova: oltre 32.000 € netti annui).',
]);
C.push(h2('Confronto con la liquidazione controllata (metodo Forlì, seguito anche altrove)'));
list([
  'Reddito: quota di stipendio acquisibile per 36 mesi (art. 272 c.3), di regola fino al quinto, considerando un mantenimento dignitoso. Verona la calcola su 13 mensilità.',
  'Beni: valore di realizzo al netto di trascrizione, pubblicità sul PVP e costi di vendita. Ribassi d’asta realistici: Genova ha stimato un realizzo di 45.000 € su una stima di 60.000 €.',
  'Costi: compenso del liquidatore secondo il DM 202/2014 (a Forlì minimo circa 1.514 € più accessori) e costi indiretti, ad esempio la perdita dell’auto necessaria per lavorare.',
  'Esclusioni: nel confronto entra solo il patrimonio del debitore. Non entrano quello del coniuge né l’eredità futura dei genitori, che non è un’aspettativa giuridica (Verona 2026). La finanza esterna condizionata all’omologa sparisce nello scenario liquidatorio.',
  'Esecuzione pendente: se un’esecuzione in corso pagherebbe integralmente l’opponente in meno tempo, il piano non è conveniente (Verona 2026, rigetto).',
]);
C.push(h2('Compensi e spese'));
list([
  'OCC: è prededucibile, ma lo liquida il giudice solo a fine esecuzione, dopo la relazione finale (artt. 71 c.4 e 81 c.4). Nel piano va accantonato pro quota e non pagato subito. Gli acconti si autorizzano solo nei limiti di quanto accantonato.',
  'Legale e advisor del debitore: non sono prededucibili, hanno solo il privilegio dell’art. 2751-bis n.2 c.c. (Forlì per la liquidazione; Genova 2026 per il concordato minore). Attenzione ai piani che li mettono in prededuzione.',
  'Opposizione respinta in un caso di rigetto: Verona ha condannato il debitore alle spese di lite (art. 91 c.p.c.).',
]);

// ---------------- 3 RDC
C.push(h1('3. Ristrutturazione dei debiti del consumatore (artt. 67-73)'));
C.push(h2('Ammissione (art. 70 c.1)'));
list([
  'Documenti: quelli dell’art. 67 c.2 e la relazione OCC dell’art. 68 c.2-3. La relazione deve trattare cause, diligenza, completezza dei documenti, convenienza, costi e merito creditizio.',
  'Proposta: deve essere almeno in parte satisfattiva e non solo dilatoria.',
  'Contenuto libero: è possibile la soddisfazione parziale e differenziata, anche dei crediti da cessione del quinto. Il debitore non è obbligato a offrire di più o più a lungo (Forlì).',
  'Effetti del decreto: pubblicazione, comunicazione ai creditori entro 30 giorni, 20 giorni per osservazioni e poi 10 giorni all’OCC per la sua relazione e le eventuali modifiche (art. 70 c.3 e c.6). Il decreto vieta azioni esecutive e atti di straordinaria amministrazione e sospende le trattenute della cessione del quinto.',
  'Primo vaglio: Verona fa già in apertura un primo confronto di convenienza.',
]);
C.push(h2('Fase delle osservazioni'));
list([
  'Modifiche: sono ammesse prima dell’omologa ma vanno ricomunicate ai creditori, con nuovi termini di 20 + 10 giorni (Verona).',
  'Nuovo debito o variazione del passivo dopo l’apertura: serve una nuova comunicazione, perché la convenienza può cambiare (Forlì).',
  'Correzioni dei crediti: se l’OCC le accoglie, si aggiornano le percentuali e si omologa il piano modificato (Genova).',
  'Ruolo dell’OCC: se dopo le osservazioni l’OCC cambia giudizio su completezza o meritevolezza, la domanda è rigettata (Verona 2026).',
]);
C.push(h2('Omologa (art. 70 c.7)'));
list([
  'Sempre: il giudice verifica ammissibilità giuridica e fattibilità.',
  'Fattibilità: guarda la durata (di solito 4-5 anni, anche 5 anni e 8 mesi a Genova; un piano di oltre 11 anni è stato ammesso in apertura a Benevento), la rata rispetto a reddito netto e spese, la stabilità del lavoro e l’effetto della sospensione di quinto e pignoramenti.',
  'Convenienza: si verifica solo se un creditore la contesta. Il credito dell’opponente non deve essere pagato meno che in liquidazione.',
  'Reddito familiare: conta per la sostenibilità (art. 67 c.2 lett. e) senza rendere la procedura familiare.',
  'Mutuo prima casa in regolare ammortamento: può continuare fuori dal piano (Verona, Savona).',
  'Morte di uno dei codebitori: la procedura diventa improseguibile per lui e prosegue per l’altro (Genova 2025).',
  'Finanziamento garantito da una fondazione antiusura: è ammesso come unica eccezione al divieto di nuovo credito (Verona 2026).',
]);
C.push(h2('Esecuzione e chiusura'));
list([
  'Dispositivo tipo: pubblicazione entro 48 ore (Verona 6 anni con anonimizzazione dei terzi; Genova per tutta la durata), PEC ai creditori entro 30 giorni, trascrizione sui beni da cedere e chiusura della procedura con l’omologa.',
  'Conto dedicato: Verona lo vincola all’ordine del giudice e prevede pagamenti almeno annuali e proporzionali tra i chirografari.',
  'Relazioni: semestrali (Forlì, Verona al 30/6 e al 31/12, trasmesse ai creditori), ogni 5 mesi a Genova, la prima entro 6 mesi a Savona. Alla fine serve una relazione finale.',
  'Fine piano: decreto di esecuzione, con inesigibilità dei debiti residui e cancellazione della pubblicazione (Genova 2026).',
  'Revoca (art. 72): per inadempimento o per interruzione degli obblighi prescritti, ad esempio il percorso SerD.',
]);
C.push(h2('Quando la RDC è stata respinta (Verona 2026)'));
list([
  'Motivi: la completezza documentale è diventata negativa dopo le osservazioni; c’è stato dolo nella crisi del mutuo; l’esecuzione pendente avrebbe soddisfatto il creditore prima e per intero.',
  'Conseguenze: rigetto con sentenza (art. 70 c.8), revoca delle misure protettive e condanna alle spese.',
]);

// ---------------- 4 CM
C.push(h1('4. Concordato minore (artt. 74-83)'));
C.push(h2('Accesso e apertura (art. 78)'));
list([
  'Legittimati: debitore non consumatore e non assoggettabile a liquidazione giudiziale, cioè impresa minore, professionista anche forfettario, agricoltore, garante o socio di società.',
  'Continuità (art. 74 c.1): il professionista anche forfettario e l’impresa minore possono continuare l’attività; la finanza esterna non è obbligatoria ma rafforza la proposta. Le rate di un leasing rinegoziato del bene strumentale possono entrare tra le spese incomprimibili per proseguire l’attività (Genova 2026). Una continuità senza cambiamenti nell’attività non basta a escludere il concordato, anche se l’Erario obietta che il debitore tornerà a non pagare le imposte (Genova 2024).',
  'Concordato liquidatorio senza continuità (art. 74 c.2): serve finanza esterna senza diritto di restituzione. Deve aumentare l’attivo “in misura apprezzabile” rispetto a quello della liquidazione, compresi 3 anni di quote di reddito ed eredità acquisibili (Verona 2026: 37.348 € contro 28.412 €).',
  'Finanza esterna: servono la dichiarazione di impegno dei terzi, la prova della loro capacità attestata dal Gestore e preferibilmente la rinuncia a surroga e regresso. Verona l’ha fatta depositare prima, con assegni circolari affidati al Gestore.',
  'Documenti (art. 75 c.1): dichiarazioni fiscali, IVA e IRAP degli ultimi 3 anni, scritture contabili, atti straordinari degli ultimi 5 anni, entrate e spese della famiglia.',
  'Relazione particolareggiata (art. 76 c.2): cause e diligenza, motivi dell’incapacità, atti impugnati, completezza dei documenti, convenienza rispetto alla liquidazione, costi, percentuali e tempi, criteri delle classi. Deve contenere anche il giudizio sul merito creditizio (artt. 76 c.3 e 80 c.4).',
  'Privilegiati pagati in parte: serve l’attestazione dell’art. 75 c.2 che non ricevono meno che in liquidazione. Il pagamento della parte capiente deve avvenire entro 6 mesi dall’omologa (art. 86 richiamato dall’art. 74 c.4, Verona 2026).',
  'Classi (art. 74 c.3): devono essere omogenee per posizione giuridica e interessi, ad esempio una classe separata per i crediti con garanzia pubblica.',
  'Commissario: viene nominato solo se richiesto o se c’è pregiudizio per i creditori.',
  'Misure protettive: valgono fino all’omologa definitiva. La sospensione degli interessi è un effetto di legge (art. 76 c.5).',
]);
C.push(h2('Voto (art. 79)'));
list([
  'Termine: 30 giorni via PEC. Il silenzio vale assenso, e Verona ha omologato senza alcun voto espresso.',
  'Maggioranze: serve la maggioranza dei crediti ammessi al voto e, se ci sono classi, la maggioranza delle classi. Se un creditore ha più del 50% dei crediti serve anche la maggioranza per teste.',
  'Chi non vota: i privilegiati pagati al 100%. Non conta nemmeno il voto contrario di chi nel frattempo non è più creditore, ad esempio per un leasing rinegoziato (Genova 2026).',
  'Voto condizionato: il voto “favorevole condizionato” vale come voto negativo (Genova 2026).',
  'Coobbligati e garanti: il creditore conserva le sue azioni contro di loro (art. 79 c.5), quindi il suo pregiudizio è limitato.',
]);
C.push(h2('Omologa (art. 80)'));
list([
  'Cram-down fiscale (art. 80 c.3): se il voto del Fisco è determinante e la proposta per il Fisco è più conveniente della liquidazione, il giudice omologa lo stesso. A Genova gli orientamenti sono diversi. Nel 2024 un giudice ha “sterilizzato” il voto del Fisco e ricalcolato le classi senza di esso. Nel 2025 un altro giudice ha escluso la sterilizzazione dell’art. 88 c.4 e ha applicato la lettera dell’art. 80, omologando anche con una convenienza di misura (21.879 € contro 20.661 €).',
  'Convenienza: il giudice ricalcola i conteggi. Per le quote societarie usa il pro quota di patrimonio netto più l’utile; per il reddito usa la quota pignorata per 36 mesi.',
  'Vendita di immobili: il giudice può imporre una gara sulla proposta d’acquisto, con almeno 30 giorni di pubblicità. La cessione diretta è possibile solo con attestazione di congruità del prezzo (art. 81 c.1).',
  'Cessione del quinto: il giudice può scioglierla (art. 97 richiamato) se serve al piano, notificando il cessionario (Verona).',
  'Sanzioni amministrative pecuniarie: restano escluse dall’esdebitazione (art. 278 c.7).',
  'Dopo l’omologa: piano dei pagamenti aggiornato entro 20 giorni (Genova); relazioni ogni 3 o 6 mesi; pubblicazione da 1 anno (Verona) a tutta la durata (Genova); possibile ordine al datore di lavoro di versare la quota sul conto della procedura (Verona); revoca dell’art. 82; inefficacia degli atti contrari al piano (art. 81 c.3).',
]);

// ---------------- 5 LC
C.push(h1('5. Liquidazione controllata (artt. 268-277)'));
C.push(h2('Apertura'));
list([
  'Soggetti: persona fisica non fallibile, ex imprenditore cancellato da oltre un anno o impresa minore. L’impresa individuale ancora attiva si liquida anch’essa, anche se il Gestore sostiene il contrario (Forlì 2026).',
  'Documenti (art. 39 richiamato dall’art. 270 c.5): dichiarazioni dei redditi degli ultimi 3 anni, beni, elenco dei creditori con le prelazioni. Savona ordina il deposito delle scritture entro 7 giorni.',
  'Relazione OCC (art. 269 c.2): completezza e attendibilità dei documenti, situazione economica e patrimoniale, cause e diligenza. Per la persona fisica serve l’attestazione dell’art. 268 c.3 sull’attivo acquisibile; se manca ma l’attivo è evidente, non serve integrarla. Se la relazione è incompleta, il giudice emette un decreto di integrazione (Forlì).',
  'Colpa grave: Savona la verifica già in apertura, guardando alle cause e alla ludopatia curata. Il giudizio è utile in vista dell’esdebitazione.',
  'Piano del debitore: è irrilevante. Vale il concorso su tutto il patrimonio, compresi i beni sopravvenuti fino all’esdebitazione (art. 272 c.3-bis), e il giudice decide modi e tempi.',
  'Beni: nessuno è escluso, nemmeno l’auto, di cui al massimo si autorizza l’uso fino alla vendita. La vendita si evita solo versando dall’esterno il controvalore (Forlì). In un concordato minore familiare a Genova l’auto strumentale di modico valore è rimasta fuori dall’attivo.',
  'Creditore fondiario: può iniziare o proseguire l’esecuzione individuale (Cass. 22914/2024, Savona).',
  'Liquidatore: di regola è il Gestore stesso (art. 270 c.2 lett. b). Savona lo autorizza ad accedere alle banche dati (artt. 155-quater e seguenti disp. att. c.p.c.).',
  'Sostituzione del liquidatore (Forlì 2026): revoca d’urgenza se il liquidatore perde i requisiti, ad esempio per sospensione disciplinare dall’albo divenuta efficace. Il nuovo liquidatore verifica lo stato della procedura, si fa consegnare tutta la documentazione dal precedente, ne valuta la regolarità e riferisce al giudice. Nelle sentenze di apertura Forlì avverte che il mancato deposito delle relazioni semestrali è causa di revoca e incide sul compenso.',
]);
C.push(h2('Limite di reddito escluso (art. 268 c.4 lett. b)'));
C.push(table(['Caso', 'Reddito e spese', 'Quota alla procedura'], [
  ['Forlì – famiglia', 'Netto 1.730 €; famiglia 3.300 € contro spese 2.815 €', '400 €/mese'],
  ['Forlì – single', '50.000 € annui', '1.000 €/mese'],
  ['Forlì – familiare', 'Entrate 11.400 €/mese; spese 4.298 €', '6.800 €/mese divisi per masse'],
  ['Genova – coniugi', 'Due stipendi 2.800 €; l’OCC proponeva 1.620 € intangibili', 'Il collegio fissa 1.300 € intangibili: 500 €/mese più il 50% di 13ª e 14ª'],
  ['Genova – coniugi', 'Un coniuge senza lavoro', '600 €/mese, più ogni entrata oltre 32.000 € netti annui; da rivedere quando l’altro torna a lavorare'],
  ['Savona – famiglia di 5', 'Netto 3.000 € + 380 € di assegni; coniuge 900 €', '1.380 €/mese per 36 mesi, come offerto'],
], [2100, 3900, 3026]));
nota('Indirizzo comune: il limite è provvisorio, si fissa in sentenza o con decreto subito dopo e si può rivedere. A Genova la quota non scende sotto il quinto pignorabile.');
C.push(h2('Svolgimento'));
list([
  'Termini: 90 giorni per le domande dei creditori; 30 giorni per aggiornare l’elenco dei creditori; 90 giorni per inventario e programma di liquidazione; relazioni semestrali; reclamo contro lo stato passivo entro 8 giorni (art. 133).',
  'Durata: almeno 3 anni (art. 272 c.3). Chiusura anticipata se non c’è più nulla da acquisire.',
  'Compenso OCC (artt. 6 e 275): se Gestore e Liquidatore sono la stessa persona, lo liquida il giudice dopo il rendiconto e non va insinuato. Se sono persone diverse, l’OCC si insinua in prededuzione.',
]);

// ---------------- 6 esdebitazione
C.push(h1('6. Esdebitazione (artt. 278-283)'));
C.push(h2('Nella liquidazione controllata'));
list([
  'Due strade: di diritto, con il decreto di chiusura (artt. 276 e 233 c.1 lett. c/d); oppure su istanza dopo 3 anni dall’apertura, anche a procedura ancora aperta (artt. 281 c.2 e 282).',
  'Procedura: istanza del debitore o segnalazione del liquidatore; relazione del liquidatore; comunicazione ai creditori ammessi con 15 giorni per le osservazioni. Osservazioni generiche non bastano.',
  'Controlli (artt. 280 e 282 c.2): nessuna condanna definitiva per i reati dell’art. 280 lett. a) senza riabilitazione; nessuna distrazione, passività inesistenti, aggravamento del dissesto o ricorso abusivo al credito; collaborazione e versamenti puntuali delle somme dovute; nessuna frode, malafede o colpa grave; mai esdebitato prima.',
  'Soddisfazione dei creditori: non serve una soglia minima. A Forlì l’esdebitazione è stata concessa con riparti di circa il 10%.',
  'Effetti: i crediti anteriori non pagati diventano inesigibili, salvi i limiti dell’art. 278 c.7. Pubblicazione per 12 mesi a Forlì e 5 anni a Genova. Iscrizione nel registro imprese solo per l’imprenditore. Reclamo (art. 124) entro 30 giorni.',
]);
C.push(h2('Casi di diniego'));
list([
  'Genova 2026: il debitore aveva condanne ostative senza riabilitazione (art. 280 lett. a). Rigetto immediato su relazione del liquidatore.',
  'Genova: distrazioni, nessuna contabilità che impediva di ricostruire il patrimonio, aggravamento del dissesto (art. 280 lett. b). Pubblicazione per 5 anni.',
]);
C.push(h2('Dopo RDC o concordato minore'));
list([
  'Effetto: l’esdebitazione deriva dall’esatta esecuzione del piano. Il decreto che dichiara eseguito il piano rende inesigibili i debiti residui e cancella la pubblicità.',
  'Esclusioni: restano escluse le sanzioni amministrative pecuniarie e quanto previsto dall’art. 278 c.7.',
  'Coobbligati e fideiussori: restano obbligati (artt. 79 c.5 e 67).',
]);
C.push(h2('Debitore incapiente (art. 283)'));
C.push(p('Nella raccolta non ci sono provvedimenti sull’art. 283. Per questa procedura servono altre fonti.'));

// ---------------- 7 familiare
C.push(h1('7. Procedura familiare (art. 66)'));
list([
  'Quando: membri della stessa famiglia conviventi, oppure con sovraindebitamento di origine comune. Esempi: coniugi garanti della figlia, coniugi con mutuo cointestato.',
  'Struttura: masse attive e passive distinte, proposta e piano unitari, un solo OCC. Nel concordato minore le maggioranze si calcolano per masse.',
  'Liquidazione controllata familiare: è ammessa (Forlì 2025) e il limite di reddito si ripartisce per masse.',
  'Attenzione: in una RDC non familiare il reddito del coniuge conta per la sostenibilità ma non rende familiare la procedura.',
]);

// ---------------- 8 confronto tribunali
C.push(h1('8. Forlì e gli altri tribunali: le differenze pratiche'));
C.push(table(['Tema', 'Forlì', 'Altri tribunali'], [
  ['Convenienza in RDC', 'Solo se contestata; confronto con la liquidazione su 36 mesi, costi inclusi', 'Verona: primo vaglio già in apertura; Genova: realizzo d’asta ridotto'],
  ['Beni in liquidazione', 'Nessuna esclusione; uso temporaneo o versamento del controvalore', 'Genova (concordato minore familiare): auto strumentale di modico valore lasciata fuori'],
  ['Limite di reddito', 'Fissato provvisoriamente in sentenza su spese documentate', 'Genova: decreto separato, minimo il quinto, 13ª e 14ª al 50%; Savona: accetta la quota offerta'],
  ['Pubblicazione esdebitazione', '12 mesi', 'Genova: 5 anni'],
  ['Pubblicazione omologa RDC o CM', 'Per tutta la durata', 'Verona: 6 anni per la RDC, 1-4 anni per il CM'],
  ['Relazioni del Gestore', 'Semestrali; SerD in ogni relazione se c’è ludopatia', 'Genova ogni 5 mesi; Verona 30/6 e 31/12 o trimestrali'],
  ['Compenso OCC', 'Liquidato dal giudice a fine procedura', 'Uguale ovunque; Benevento autorizza acconti nei limiti accantonati'],
  ['Divieto di credito', 'Non previsto di regola', 'Savona e Verona: divieto di carte e credito per tutto il piano, con comunicazione alla Centrale Rischi'],
], [2300, 3300, 3426]));

// ---------------- 9 checklist
C.push(h1('9. Checklist per il Gestore'));
C.push(h2('Relazione per RDC o concordato minore'));
list([
  'Visure: Centrale Rischi, CRIF, cassetto fiscale, estratti di ruolo AdER, Anagrafe dei rapporti finanziari, casellario giudiziale e carichi pendenti.',
  'Merito creditizio: tabella rata/reddito alla data di ogni finanziamento. Se il merito non è stato verificato, citare l’art. 69 c.2.',
  'Cause del debito: ricostruite con documenti (malattia, separazione, crisi d’impresa). Per la ludopatia, certificato SerD.',
  'Spese di mantenimento: documentate e confrontate con i dati ISTAT; reddito del coniuge indicato; margine prudenziale.',
  'Confronto con la liquidazione: 36 mesi di quota di reddito (13 mensilità), beni netti, costi del liquidatore, finanza esterna esclusa.',
  'Privilegiati pagati in parte: attestazione dell’art. 75 c.2. Nel concordato minore, pagamento entro 6 mesi della parte capiente.',
  'Prededuzione: solo il compenso OCC, da accantonare. Il legale e l’advisor del debitore hanno solo privilegio.',
  'Finanza esterna: impegno scritto, prova della capacità, rinuncia al regresso, meglio se depositata prima.',
  'Concordato minore: classi motivate, regola delle maggioranze con la verifica per teste, ipotesi di cram-down fiscale.',
]);
C.push(h2('Relazione per la liquidazione controllata'));
list([
  'Contenuti: giudizio di completezza e attendibilità dei documenti, cause e diligenza, attestazione dell’attivo acquisibile (art. 268 c.3).',
  'Limite di reddito: proposta motivata con spese documentate, mai sotto il quinto.',
  'Impresa: verificare se l’attività è ancora in corso e se è cessata da meno di un anno.',
  'Beni: elenco completo, compresi quelli che il debitore vorrebbe tenere, con il valore di realizzo.',
]);
C.push(h2('Dopo l’apertura o l’omologa'));
list([
  'Adempimenti iniziali: pubblicazione entro 48 ore, PEC ai creditori entro 30 giorni, trascrizioni, conto dedicato, comunicazione al datore di lavoro.',
  'Durante l’esecuzione: relazioni periodiche alle scadenze fissate, certificati SerD, segnalazione immediata dei fatti che possono portare alla revoca (artt. 72 e 82).',
  'In liquidazione: relazione per l’esdebitazione entro il terzo anno, con i 15 giorni ai creditori per le osservazioni.',
]);

const doc = new Document({
  creator: 'OCC – materiale di studio',
  title: 'Sovraindebitamento: i criteri dei giudici',
  background: { color: 'FDFBF5' },
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  numbering: { config: [{ reference: 'punti', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 500, hanging: 300 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('Uso interno – schede anonime – pag. ', { size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
    children: C,
  }],
});
Packer.toBuffer(doc).then(buf => fs.writeFileSync(process.argv[2] || 'Criteri_Sovraindebitamento.docx', buf));
