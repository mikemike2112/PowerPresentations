# COMPAT — compatibilità con PowerPoint (ECMA-376)

Tutti i numeri qui sotto vengono dall'ultima esecuzione reale dei test (`dotnet test`) e degli script in `tools/`.

## Corpus

`tests/corpus/office-made/`: **35 file** generati dal vero PowerPoint (Microsoft 365, interfaccia italiana, tema "Office" con Aptos) da
`tools/make-corpus.ps1` (rigenerabile; Author/LastAuthor = `Corpus`, verificato da `Corpus_Metadata_*`; nessun dato personale nelle part).
Un file per funzionalità (titolo, 11 layout, master personalizzato, colori/font del tema, testo, elenchi, forme, rotazioni, connettori, immagini
con ritaglio, tabelle e celle unite, note, 6 transizioni, animazioni, slide nascosta, sezioni, sfondi, autofit, gruppi, collegamenti, commenti, piè
di pagina, layout personalizzato, 4:3, ombre/effetti, WordArt) più due file "misti" realistici.

**Non nel corpus**: grafici (non aggiunti) e SmartArt: su questa installazione `Application.SmartArtLayouts.Count` vale 0, quindi `Shapes.AddSmartArt(SmartArtLayouts.Item(1), …)` fallisce con "Impossibile chiamare un metodo su un'espressione con valore null". Questi oggetti sono comunque conservati dal round-trip (provato con part custom e con `ChartSmartArtTests`: file generato con l'SDK con grafico + workbook incorporato e SmartArt a 4 part; dopo modifica e salvataggio le part e il binario incorporato restano identici, senza modifiche il package è identico, l'export PDF non fallisce; non provato con file reali di Office).

## Round-trip (per ognuno dei 35 file)

| Criterio | Esito |
|---|---|
| a. salva col clone → `OpenXmlValidator` (Microsoft365): nessun errore nuovo | **35/35** (gli originali di PowerPoint hanno 0 errori del validatore: nessun falso positivo da elencare) |
| b. il salvato si riapre nel vero PowerPoint via COM, senza errori né ripristino | **35/35** (100%) |
| c. contenuto identico letto da Office (n. slide/forme, testo, tabelle, note, transizione, nascosta, dimensioni) | **35/35** |
| d. stessa lista di part e relazioni, ogni part XML semanticamente identica (canonicalizzata), binari identici | **35/35** |
| e. controllo negativo | vedi sotto |
| f. test con modifica (`OfficeEditTests`: testo, riordino, slide, note, transizioni, oggetti, incolla di forme/grafico/media, elenchi, collegamento, sostituzione, copia formato, commento, intestazione e piè di pagina, imposta presentazione, intervalli, presentazione personalizzata, media) | tutti letti correttamente da PowerPoint |

Modifiche verificate in Office: testo cambiato (resto intatto), riordino (note ancora agganciate alla slide giusta), nuova slide da layout,
eliminazione con sezioni, note create/modificate, transizione, casella/forma/immagine/tabella inserite, duplicazione (con note), slide
nascosta → visibile, cella di tabella unita, nuova presentazione dal modello.

Contenuto sconosciuto (`UnknownContent_SurvivesEditAndSave`): part custom con relazione, elemento di namespace estraneo dichiarato in `mc:Ignorable`,
attributo sconosciuto, `extLst` — tutto presente dopo modifica, riordino, note e salvataggio, validatore senza errori nuovi.

### Controllo negativo del rilevatore

`tools/office-inspect.ps1` apre con `Presentations.Open(..., ReadOnly, WithWindow=msoFalse)`. Il rilevatore **fallisce** su:
XML rotto in una slide (`PowerPoint could not open the file`), relazione/part mancante (layout cancellato: `0x80070570`), attributo numerico invalido
(`<a:off x="abc…`). Il file buono dello stesso test passa. In questa configurazione PowerPoint **rifiuta** i file corrotti con un'eccezione COM
invece di mostrare il dialogo di ripristino; il caso del dialogo modale è comunque coperto dal watchdog per PID (`tools/watchdog.ps1`):
`tools/test-watchdog.ps1` dimostra che, fermando il battito, viene ucciso **solo** il PID di PowerPoint avviato dallo script (`OfficeNegativeControlTests`).
Un eventuale ripristino silenzioso che alteri il contenuto verrebbe colto dal confronto (c). Esito della ricerca di un file "riparato": con id di slide o di layout duplicati PowerPoint rifiuta il file (0x80070570); con larghezza di colonna negativa lo apre correggendolo in silenzio (contenuto diverso, colto dal confronto (c): `Detector_CatchesSilentRepair_*`); altre alterazioni (embed inesistente, enum non valido, id forma duplicato, content type mancante, elemento sconosciuto) sono aperte senza differenze nel contenuto letto. Un dialogo di ripristino non è mai comparso.

## Confronto visivo (PDF PowerPoint vs PDF clone, rasterizzati con `Windows.Data.Pdf`)

`tools/visual-compare.ps1` produce per ogni file l'immagine affiancata. Differenza media di pixel (a 160×90) ≤ 0,5% su 30 file su 35; i valori più alti
sono tutti spiegati da differenze reali osservate guardando le immagini:

- **Ombre, bagliore, smusso** (`shadow-effects`, 1,3%): non disegnati (nessun supporto a `effectLst`/`effectRef`).
- **Riempimento a motivo** (`shape-fills-lines`): il clone usa il colore di primo piano pieno invece del motivo.
- **Geometrie predefinite non implementate** (`shapes-basic`, cuore e fulmine): disegnate come rettangolo; vale per tutte le geometrie fuori dalle 18 supportate.
- **WordArt/contorno del testo** (`wordart`): il clone disegna il testo con riempimento pieno, senza contorno/gradiente del testo.
- **Gradienti** (`slide-background` slide 2, `shape-fills-lines`): interpolazione lineare in sRGB tra gli stop; i toni intermedi differiscono leggermente da PowerPoint.
- **Rimpicciolimento automatico** (`autofit-shrink`, 2,8%): PowerPoint, aprendo in sola lettura senza finestra, non ricalcola `fontScale`; il clone fa lo stesso
  (usa il valore memorizzato) e lo ricalcola solo quando si modifica il testo. Differenza residua = larghezza dei caratteri.
- **Interlinea multipla** (`paragraph-align-spacing`): ~1–2 pt di scarto verticale con interlinea 2,0.
- **Font**: il tema Office usa Aptos (font cloud). Il clone lo trova in `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts` e lo usa sia per l'impaginazione
  sia per il disegno (GlyphRun dal file): larghezze identiche, niente crenatura (kerning) → scarti sub-pixel sul testo lungo. Se il font non è installato si
  ripiega su Segoe UI poi Arial.
- **Colori `tint`/`shade`**: Office li calcola in RGB lineare (verificato sui pixel del PDF: 204,210,216 per accent1 tint 40%); il clone fa lo stesso.

## Limiti noti (non perdono dati: gli elementi non gestiti si conservano)

- Grafici, SmartArt, OLE, video/audio: segnaposto (i video/audio inseriti dal clone sono forme multimediali vere di PowerPoint, con anteprima generica e senza riproduzione nel clone) con etichetta; `p:oleObj` e `mc:AlternateContent` con fallback immagine usano il ramo `Fallback`.
- Testo: niente giustificazione, niente colonne, niente testo verticale/ruotato di 90°, niente crenatura, niente `a:arcTo` nei freeform, small caps ≈ normale.
- Gruppi: la rotazione/ribaltamento del gruppo non si applica ai figli (la posizione sì).
- Immagini: EMF/WMF/SVG senza PNG di ripiego → riquadro grigio; maschere di ritaglio a forma ignorate.
- Slide numero/data: campo `slidenum` calcolato, `datetime` mostra il testo memorizzato.
- Transizioni: gli effetti p14 (vortice, cubo…) si leggono come "altro" e si riproducono in dissolvenza. Una transizione che l'utente non cambia resta identica (`mc:AlternateContent` e `p14:dur` conservati, test `TransitionDataLossTests`). Se la cambia, il clone riscrive l'`mc:AlternateContent` con la nuova transizione conservando `p14:dur` (o 500/1000/2000 ms se cambia la velocità); si perdono solo gli attributi p14 specifici dell'effetto precedente, e passando a "Nessuna" l'intero blocco viene rimosso.
- Vista relatore e secondo monitor: logica implementata (presentazione sul secondo schermo, relatore sul primo), **non verificata** (PC con un solo schermo).
- Duplicazione slide: grafici e SmartArt (con workbook incorporato) vengono copiati in part distinte, non condivise con l'originale (`ChartSmartArtTests`); la copia non ha commenti (le note sì) né il riferimento `p188:commentRel` dell'originale; i commenti dell'originale restano intatti (`CommentsTests`, `CommentEditTests`).
- `.pptm`, `.ppt`, file protetti da password: non si aprono (messaggio chiaro).

## Comandi aggiunti (formato scritto)

- **Appunti**: le forme copiate viaggiano come JSON con XML + part collegate (immagini, grafico con workbook, SmartArt, media, collegamenti); incollando, id forma e relazioni sono rinumerati e le part ricreate (grafico e SmartArt con part proprie). I segnaposto incollati diventano forme libere con posizione esplicita (perdono la formattazione ereditata dal layout). I collegamenti con schemi diversi da http/https/mailto sono scartati; gli appunti sono trattati come input non fidato.
- **Commento**: moderno (`ppt/comments/modernComment_*.xml`, `ppt/authors.xml`, `p188:commentRel`, `p14:creationId` della slide); l'autore è sempre «Autore» (mai il nome di Windows). Se la slide ha già commenti classici si aggiunge un commento classico. Il riquadro Commenti legge, modifica ed elimina anche i commenti creati da PowerPoint: moderni (`p188:cm` con `p188:replyLst`) e classici (`p:cm`, risposte via `p15:parentCm`). La modifica riscrive solo il testo (autore, data, id, posizione e gli altri commenti non cambiano); l'eliminazione toglie il commento con le sue risposte e, se era l'ultimo della slide, anche la part, la relazione, l'override del content type e `p188:commentRel`. Gli autori (`authors.xml`, `commentAuthors.xml` con `lastIdx`) restano com'erano, come fa PowerPoint. Verificato con validatore Open XML e con PowerPoint (`OfficeEditTests.Comments_EditedAndDeletedByClone_*`); i commenti classici sono provati solo su file costruiti nei test (il corpus contiene solo un commento moderno).
- **Intestazione e piè di pagina**: crea i segnaposto `dt`/`ftr`/`sldNum` solo se il layout li prevede.
- **Imposta presentazione**: `p:showPr` in `presProps.xml` (modalità, ciclo, narrazione/animazioni/intervalli, intervallo `sldRg`). Gli effetti sulla riproduzione del clone: intervallo di slide e ciclo; narrazione assente.
- **Prova intervalli / Applica a tutte**: attributo `advTm` di `p:transition`, mantenuto anche quando cambia l'effetto (anche dentro `mc:AlternateContent`).
- **Presentazione personalizzata**: `p:custShowLst`; eliminando una slide la voce esce dall'elenco.
- **Video/Audio**: `p:pic` con `a:videoFile`/`a:audioFile` e `p14:media`, relazioni video/audio/media sulla stessa part dati; tetto 50 MB (annulla/ripristina copiano il documento in memoria).
- **Non implementati (restano disattivati)**: Registra (narrazione: microfono + incorporamento audio) e Icone (libreria vettoriale).

## Finestre di dialogo sgranate (causa reale)

Misurata con catture dello schermo vero a scala Windows 125% (`tools/dialog-screen.ps1`, CopyFromScreen): la finestra di salvataggio era nitida solo senza `AllowsTransparency`.
Una finestra WPF a strati (`AllowsTransparency=true`) disegna il testo in software, senza ClearType, e Windows la riscala sopra il 100%. Spostare il `DropShadowEffect` su un elemento
separato non bastava (provato: testo ancora morbido). Soluzione: finestra normale con `WindowChrome` e angoli/ombra del sistema (DWM). Dopo la modifica il testo è nitido a 125%.
