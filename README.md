# Lettore Presenze AI

App web installabile (PWA) che legge con l'intelligenza artificiale (Google Gemini) i rapportini di cantiere cartacei
(Mod. 22.03, fronte e retro) e scrive commesse e ore direttamente nel file Excel delle presenze.

Funziona interamente nel browser (Microsoft Edge o Google Chrome): non serve installare programmi né avviare un server.

## Come si usa

1. Apri l'indirizzo dell'app con Edge o Chrome e clicca **Installa app** nella barra degli indirizzi.
2. Al primo avvio clicca **Collega cartella Excel** e scegli la cartella che contiene i file `.xlsm` delle presenze.
3. In **Configura API** inserisci la tua chiave Google Gemini (resta salvata solo nel browser di quel PC).
4. Trascina il PDF del rapportino, premi **Analizza con AI Vision**, controlla la tabella e premi **Conferma e scrivi su Excel**.

## Dati e privacy

- I file Excel, i rapportini e la chiave API **restano sul PC**: non vengono caricati su questo sito.
- Le immagini dei rapportini vengono inviate solo a Google Gemini per la lettura.
- Il codice non contiene nomi di persone: l'elenco degli operai e delle commesse viene letto dal file Excel.

## Protezioni

- Prima di ogni scrittura viene salvata una copia del file Excel con data e ora nella sottocartella `Backup` (ultime 20 copie).
- Se in Excel ci sono già dati diversi per gli stessi operai e turni, l'app chiede cosa fare prima di sovrascrivere.
- Se il file è aperto in Excel, la scrittura viene fermata.
- Nella scansione multipla i rapportini con dubbi non vengono scritti finché non li controlli.
