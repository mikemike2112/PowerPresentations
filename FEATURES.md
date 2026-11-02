# FEATURES — inventario e stato

## Metodo dell'inventario

Lo strumento computer use non è disponibile, quindi l'app vera è stata ispezionata così:

1. `tools/ribbon-inventory.ps1` avvia PowerPoint con la **finestra visibile** su una presentazione vuota creata dallo script
   (nessun documento dell'utente), legge il ribbon con **UI Automation** (`System.Windows.Automation`: schede, gruppi, comandi)
   e scrive `tools/ribbon-inventory.json` (11 schede, circa 230 controlli in italiano).
2. Valori predefiniti e comportamento reale (layout, stili, tabelle, transizioni, note, autofit, formato dei file) sono stati osservati
   generando il corpus con `tools/make-corpus.ps1` (35 file creati dal vero PowerPoint via COM) e confrontando i PDF esportati da PowerPoint
   con quelli del clone (`tools/visual-compare.ps1`).
3. Nessuna decompilazione: solo comportamento osservabile, file prodotti e specifica ECMA-376.

Priorità: **P0** = richiesto per la prima consegna, **P1** = importante, **P2** = utile. Si spunta `[x]` solo ciò che è implementato **e** coperto da
un test automatico (nome del test o dello script tra parentesi). `[ ]` = non fatto, oppure fatto ma senza test automatico (indicato).

## File

- [x] P0 Apri `.pptx` (Open XML SDK, file non fidato: errori → messaggio chiaro) — `EditTests.MalformedInput_*`, `RoundTripTests`
- [x] P0 Salva / Salva con nome, atomico (file temporaneo + sostituzione) — `EditTests.Save_IsAtomic_*`
- [x] P0 Nuova presentazione da modello generato da codice (`TemplateBuilder`: tema "Prisma", master, 6 layout originali, nessun file Office) — `EditTests.New_FromTemplate_*`, `OfficeEditTests.NewFromTemplate_OpensInOffice`
- [x] P0 Round-trip senza perdite (part, relazioni, elementi/attributi sconosciuti, `mc:Ignorable`, `extLst`, part custom) — `RoundTripTests.*` su 35 file + `UnknownContent_*`
- [x] P0 Esporta PDF (PDFsharp, slide nascoste escluse come in PowerPoint) — `PdfTests.*`, confronto visivo
- [ ] P1 Salva come `.pptm`/`.potx`, apertura `.ppt` / `.pptm` (oggi: messaggio di formato non valido)
- [x] P1 Annulla / Ripristina (pulsanti, Ctrl+Z/Ctrl+Y) — `UiTests.RealInput_UndoRedo_*`
- [x] P2 Stampa (una slide per pagina, nascoste escluse; finestra di stampa di Windows) — `AppCommandTests.Print_WritesOnePagePerVisibleSlide_ToXps` (file XPS, nessuna stampante)
- [ ] P2 Esporta come video/immagini, protezione con password

## Home

- [x] P0 Nuova diapositiva con scelta del layout — `OfficeEditTests.AddSlide_*`
- [x] P0 Duplica / Elimina / Nascondi diapositiva (note e sezioni coerenti) — `EditTests.DeleteSlide_*`, `Duplicate_*`, `OfficeEditTests.Unhide_*`
- [x] P0 Grassetto, Corsivo, Sottolineato, dimensione (A+/A−), colore carattere (tavolozza) — `EditTests.SetShapeText_KeepsRunFormat*` (XML); i pulsanti sono provati a mano
- [ ] P1 Tipo di carattere, spaziatura caratteri, barrato, apice/pedice (solo rendering: `RendererTests`/confronto visivo, nessun comando)
- [x] P1 Elenco puntato / numerato (pulsante a due stati, tabelle comprese) e allineamento — `CommandTests.Bullets_*`, `OfficeEditTests.Bullets_ReadByOffice_*`, `UiTests.RealInput_Commands_*`
- [ ] P1 Layout della slide esistente, Ripristina, Sezioni (le sezioni si conservano e si aggiornano all'eliminazione)
- [x] P1 Copia/Taglia/Incolla di forme (immagini, grafici con workbook, SmartArt, media e collegamenti copiati con le loro part; anche fra documenti; testo e immagini dagli appunti di Windows), Ctrl+C/X/V, slide copia/taglia/incolla dalle miniature, menu tasto destro — `CommandTests.CopyPaste_*`, `OfficeEditTests.Paste_*`, `UiTests.RealInput_Commands_*`
- [x] P1 Copia formato (carattere, paragrafo, riempimento e contorno; clic sulla forma di destinazione) — `TextCommandTests.CopyFormat_*`, `OfficeEditTests.CopyFormat_*`
- [x] P2 Trova e sostituisci (Ctrl+F/H, tutta la presentazione, tabelle comprese), Seleziona (tutto, Ctrl+A, scelta da elenco, Ctrl+clic) — `TextCommandTests.FindReplace_*`, `UiTests.RealInput_Commands_*`
- [ ] P2 Dettatura, Designer

## Inserisci

- [x] P0 Casella di testo — `OfficeEditTests.InsertedObjects_AreRealOfficeShapes`
- [x] P0 Forme base (18 geometrie: rettangolo, arrotondato, ellisse, triangoli, rombo, parallelogramma, trapezio, esagono, ottagono, croce, 4 frecce, stella, cheuron, pentagono) — idem
- [x] P0 Immagini PNG/JPEG/GIF/BMP — idem
- [x] P0 Tabella (dimensioni libere, stile "Medio 2 – Colore 1") e modifica del testo di cella da API — `OfficeEditTests.TableCellEdit_*`; UI: inserimento 3×3, modifica celle non ancora disponibile da interfaccia
- [ ] P1 Altre forme predefinite (callout, flowchart, cuore, fulmine… oggi disegnate come rettangolo), connettori con aggancio, linee
- [x] P1 Video e audio incorporati (MP4/MOV/WMV/AVI, MP3/WAV/M4A/WMA, tetto 50 MB, anteprima generica) come forme multimediali vere — `MediaCommandTests.*`, `OfficeEditTests.Media_*` (non riprodotti nel clone)
- [ ] P1 Grafico, SmartArt, equazioni, WordArt, icone, 3D (si **conservano** nel file e il rendering mostra un segnaposto: grafico con workbook incorporato e SmartArt provati dopo modifica e salvataggio, `ChartSmartArtTests`)
- [x] P1 Commenti: il comando Commento apre il riquadro Commenti (elenco della diapositiva corrente con autore, data e testo, aggiornato al cambio slide); aggiunta (autore generico «Autore», moderno o classico se la slide ha già commenti classici), modifica del testo ed eliminazione (con conferma se ci sono risposte, che spariscono con il commento) per tutti i commenti del file, moderni con risposte e classici; un solo Ctrl+Z per operazione — `CommentEditTests`, `ReviewCommandTests.Comment_*`, `OfficeEditTests.Comments_*`, `UiTests.RealInput_CommentsPane_*`; i commenti non sono disegnati sulla slide
- [x] P1 Collegamento ipertestuale su tutto il testo di una forma (solo http, https, mailto) — `TextCommandTests.Hyperlink_*`, `OfficeEditTests.Hyperlink_*`; selezione parziale di testo non supportata
- [x] P2 Intestazione e piè di pagina (numero, data fissa/automatica, piè; slide corrente o tutte; solo dove il layout ha il segnaposto), Simboli (elenco di 28) — `ReviewCommandTests.HeaderFooter_*`, `TextCommandTests.AppendText_*`, `OfficeEditTests.HeaderFooter_*`

## Progettazione (temi e schemi)

- [x] P0 Slide basate su master e layout; ereditarietà slide → layout → master → tema (posizioni, stili di testo per livello, `clrMap`, sfondo, forme del master/layout, `showMasterSp`) — `RendererTests.*`, confronto visivo
- [x] P0 Temi letti dal file: schema colori (con trasformazioni lumMod/lumOff/tint/shade/alpha) e font maggiore/minore — `RendererTests.ThemeColors_*`, `ThemeFonts_*`, `ColorTransforms_*`
- [x] P0 Rimpicciolimento automatico del testo (`normAutofit`: usa `fontScale` memorizzato; ricalcolo e scrittura dopo la modifica) — `RendererTests.Autofit_*`, `EditTests.SetShapeText_OnAutofitPlaceholder_*`
- [x] P0 Formato 4:3 / 16:9 e dimensioni personalizzate lette dal file — `PdfTests` (4:3), `OfficeRoundTripTests`
- [ ] P1 Cambio tema/varianti/colori/font, Dimensioni diapositiva, Stili sfondo (editor dello schema diapositiva)
- [ ] P2 Designer

## Visualizza

- [x] P0 Vista Normale con miniature a sinistra, area di modifica e note — `UiTests.*` (finestra reale)
- [x] P0 Vista a miniature "Sequenza diapositive" — `UiTests.MainWindow_RendersCorpusFile_AndSlideSorterView`
- [x] P0 Riordino delle slide con drag & drop (tocca solo `sldIdLst`; part, relazioni, note e commenti restano) — `EditTests.Reorder_*`, `UiTests.RealWindow_*` (trascinamento reale col mouse)
- [x] P0 Note del relatore (lettura, scrittura, creazione di notesMaster/notesSlide) — `OfficeEditTests.Notes_*`
- [x] P1 Tema scuro (predefinito) / chiaro commutabile — screenshot `docs/app-light.png` (nessun test automatico)
- [x] P1 Righello e griglia (cm, solo a schermo), Pagina note (immagine + note modificabili, scorri le slide) — `UiTests.RealInput_Commands_*` (griglia/righello)
- [ ] P1 Zoom, guide, Schema diapositiva

## Transizioni e presentazione

- [x] P0 Transizioni base lette/scritte: nessuna, dissolvenza, spingi, cancella, copri, scopri (direzione e velocità) — `OfficeEditTests.Transition_*`, `RendererTests.Transitions_AreRead`
- [ ] P0 Riproduzione animata delle transizioni nella presentazione (fade/push/cover/pull/wipe): implementata (`ShowWindows.cs`), **senza test automatico** dell'animazione
- [x] P0 Modalità presentazione a schermo intero (F5, Maiusc+F5, frecce, Esc, B per schermo nero, salta le slide nascoste) — `UiTests.RealWindow_*` (F5 + freccia reali)
- [x] P0 Vista relatore (slide corrente, successiva, note, timer, contatore) — `UiTests.PresentationAndPresenterWindows_Render` (rendering; **non verificata** su due monitor: un solo schermo sul PC di sviluppo)
- [ ] P1 Altri effetti (cubo, vortice, p14…): letti come "altro", riprodotti come dissolvenza; avanzamento automatico, suoni, anteprima
- [x] P1 Presentazione personalizzata (crea/modifica/elimina/avvia), Prova intervalli (tempi per slide salvati come avanzamento automatico, usato dalla presentazione), Applica a tutte, Imposta presentazione (tipo, ciclo, intervallo di slide, senza narrazione/animazioni, usa intervalli) — `CustomShowTests`, `TimingTests`, `ReviewCommandTests.ShowSettings_*`, `OfficeEditTests.*`
- [ ] P1 Animazioni degli oggetti (si conservano nel file, non riprodotte), registrazione di narrazione e video (**Registra** resta disattivato: richiede microfono e un formato audio da incorporare), Icone (**Icone** resta disattivato: servirebbe una libreria di icone vettoriali da convertire in geometrie)

## Revisione / altro

- [ ] P1 Controllo ortografico, traduzione
- [ ] P2 Accessibilità, protezione, firma, macro (mai eseguite: i file con macro non si aprono)

## Finestre di dialogo e salvataggio

- [x] P1 Dialoghi propri a tema (`ThemedDialog`, niente `MessageBox`): richiesta di salvataggio «Vuoi salvare le modifiche a …?» con Salva / Non salvare / Annulla; per un documento senza file, campo Nome file e Percorso con "Sfoglia…"; Salva con nome con "Altre opzioni…"; conferma di sostituzione — logica in `SavePlanTests` (nome non valido, sovrascrittura, mappatura pulsanti); aspetto: `docs/dialog-salva-scuro.png`, `docs/dialog-salva-chiaro.png`. La verifica con input reale (`RealInput_UndoRedo_*`, adattato a "Non salvare") non è stata rieseguita dopo questa modifica.
- [x] Manifest con Common Controls v6 e DPI PerMonitorV2 (`app.manifest`).

## Scudo privacy, finestra e ribbon adattivo

- [x] P1 Scudo privacy (scheda Powers): analisi con valori reali (autore, ultimo autore, società, manager, proprietà personalizzate, autori di commenti, percorsi locali, miniatura), copia `nome (pulito).pptx` con scelta anonimizza/rimuovi commenti, ricontrollo della copia, avvisi per ciò che non pulisce (EXIF, oggetti incorporati, macro, slide nascoste, note, collegamenti locali) — `PrivacyTests`, `PrivacyOfficeTests` (copia aperta nel vero Office senza ripristino, stesso contenuto). Screenshot: `docs/scudo-privacy.png`.
- [x] Pulsante ingrandisci/ripristina con icona e tooltip che seguono lo stato ("Ingrandisci" / "Ripristina giù") — `UiTests.WindowResize_*` (UI Automation).
- [x] Ribbon adattivo: gruppi compatti (solo icone) e poi menu "Altri comandi"; barra del titolo e schede si stringono fino a 640 DIP — `UiTests.WindowResize_*`, screenshot `docs/resize-*.png`.
- Voci tolte dall'interfaccia (il contenuto già nei file resta conservato): Inserisci → Schermata, Album foto, SmartArt, Grafico, Azione, WordArt, Registrazione schermo; Home → Layout, Ripristina (gruppo Diapositive), Sezione, Disponi. Annulla/Ripristina in alto restano.
