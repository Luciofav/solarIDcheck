# solarIDcheck

Web app leggera per il manutentore di impianti fotovoltaici: scansiona con la fotocamera dello smartphone i barcode dei pannelli e li confronta con il file Excel dell'impianto.

Nessun server, nessun account: i dati restano sul telefono e l'app funziona anche offline dopo la prima apertura.

## Come si usa

0. **Accesso**: al primo avvio l'app chiede la chiave di accesso. Su quel dispositivo viene ricordata, quindi non va reinserita ogni volta.
1. **Inizio giornata**: inserisci il **nome dell'impianto** (obbligatorio) e, se lo hai, carica il file Excel con le colonne `N° SERIALE | MODELLO`. Il file è opzionale e si può caricare anche alla fine, dal report.
2. **Scansione**: inquadra le etichette. L'app cerca sempre **due barcode** nell'inquadratura:
   - se li trova, li registra subito come **due pannelli distinti**;
   - se dopo circa 0,7 s ne vede uno solo, registra quello.

   Un codice viene accettato solo se è letto in almeno 2 fotogrammi, per evitare letture errate. Un codice appena registrato non viene contato di nuovo finché resta inquadrato.
   - Verde / bip acuto: seriale presente nel file
   - Arancione / doppio bip: seriale **non** presente nel file
   - Rosso / bip grave: seriale **già scansionato**

   Se l'etichetta è rovinata, puoi scrivere a mano il seriale stampato sotto il barcode.
3. **Termina**: il report mostra tutte le incongruenze e permette di scaricare (o condividere) i due file CSV.

La sessione viene salvata a ogni lettura: se il telefono si blocca o la pagina si chiude, alla riapertura trovi **Riprendi**.

## Incongruenze rilevate

| Tipo | Significato |
|---|---|
| `NON_IN_ELENCO` | Scansionato in campo ma assente dal file impianto. Se esiste un seriale del file che differisce di un solo carattere, viene suggerito come possibile lettura errata. |
| `NON_SCANSIONATO` | Presente nel file ma mai trovato in campo |
| `SCANSIONE_RIPETUTA` | Stesso seriale letto in momenti diversi |
| `DUPLICATO_IN_FILE` | Stesso seriale su più righe dell'Excel (segnala se i modelli sono diversi) |
| `DATI_FILE_INCOMPLETI` | Riga dell'Excel senza seriale o senza modello |
| `FORMATO_ANOMALO` | Lunghezza del seriale diversa da quella prevalente (probabile lettura parziale) |
| `INSERITO_A_MANO` | Seriale digitato a mano: da ricontrollare |

## File esportati

I file sono CSV con separatore `;` e codifica UTF-8 (BOM), così si aprono direttamente in Excel in italiano.

- `<impianto>_<data>_scansioni.csv`: N, Seriale, Modello, Esito, Data, Ora, Origine (Barcode/Manuale), Lettura n., Barcode nella lettura (1 o 2), Letture ripetute, Impianto
- `<impianto>_<data>_incongruenze.csv`: Tipo, Descrizione, Seriale, Modello, Dettaglio, Impianto

## Formato del file Excel

Primo foglio, riga di intestazione con i nomi delle colonne. La colonna del seriale viene riconosciuta da *seriale / serial / S/N / matricola*, quella del modello da *modello / model*. Se l'intestazione non viene riconosciuta, si usano la prima e la seconda colonna. Sono accettati anche `.xls` e `.csv`.

Un esempio è in [`esempio/impianto_esempio.xlsx`](esempio/impianto_esempio.xlsx).

## Pubblicazione su GitHub Pages

La fotocamera funziona solo in **HTTPS**: GitHub Pages va benissimo.

1. Crea un repository su GitHub (es. `solarIDcheck`) e carica tutti i file di questa cartella.
2. Vai su **Settings → Pages → Build and deployment**, scegli *Deploy from a branch*, branch `main`, cartella `/ (root)`.
3. Dopo un minuto l'app è su `https://<utente>.github.io/solarIDcheck/`.
4. Sul telefono apri il link e scegli **Aggiungi a schermata Home**: si apre a tutto schermo e funziona offline.

Ogni volta che modifichi i file, aumenta `VERSION` in `sw.js`: così i telefoni scaricano la nuova versione.

### Prova in locale

```bash
python3 -m http.server 8000
```

Poi apri `http://localhost:8000`. La fotocamera del computer funziona su `localhost`; per provare col telefono serve HTTPS (GitHub Pages).

## Chiave di accesso

Nel codice è salvata solo l'impronta SHA-256 della chiave, non la chiave in chiaro. Per cambiarla:

```bash
printf '%s' 'NUOVA-CHIAVE' | shasum -a 256
```

Copia il risultato in `ACCESS_HASH` dentro `app.js` e aumenta `VERSION` in `sw.js`. I dispositivi già sbloccati dovranno inserire la nuova chiave.

Nota: è un'app statica e pubblica. La chiave tiene fuori chi non la conosce, ma non è una protezione forte come un vero login con server.

## Struttura

```
index.html            interfaccia (3 schermate: avvio, scansione, report)
style.css             stile (tema chiaro/scuro automatico)
app.js                logica: Excel, scansione, analisi, export CSV
sw.js                 cache offline
manifest.webmanifest  installazione come app
vendor/               librerie incluse (nessuna dipendenza esterna a runtime)
  xlsx.full.min.js      SheetJS 0.18.5 – lettura Excel
  barcode-detector.js   barcode-detector 3.2.2 – rilevatore barcode
  zxing_reader.wasm     zxing-wasm 3.1.3 – motore di decodifica
```

**Lettura barcode**: su Android (Chrome) viene usato il rilevatore nativo del sistema, il più veloce. Su iPhone e sugli altri browser si usa ZXing in WebAssembly. Entrambi leggono più barcode nello stesso fotogramma. Sono abilitati solo i formati lineari (Code 128, Code 39, Code 93, Codabar, ITF, EAN, UPC).
