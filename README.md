# Database gestione associazione

Questo progetto contiene un database SQLite gia pronto per gestire:

- associati e anagrafica
- quote annuali di tesseramento
- tipologie di corsi e anagrafica corsi
- iscrizioni ai corsi
- quote mensili dei corsi
- iscritti al campo estivo e pagamento quota di partecipazione una tantum
- eventi multipli con iscrizioni e pagamento quota di partecipazione una tantum

In piu, la web app locale ora include anche:

- modifica ed eliminazione di tutti i dati inseriti
- export dei report in Excel e PDF
- numerazione automatica progressiva e univoca per associati, corsi, campi estivi ed eventi
- generazione massiva delle quote mensili corsi per mese selezionato
- controllo anti-duplicazione in fase di generazione massiva quote mensili
- report partecipanti filtrabili per corso, campo estivo o evento
- report situazione tesseramenti
- ricerca globale in tutti i report su tutte le colonne visualizzate
- ricerca globale anche nelle viste `Dati inseriti`
- ricevute di pagamento stampabili con export PDF, preparazione email e messaggio WhatsApp
- pagamento multiplo di piu quote mensili con ricevuta unica e dettaglio delle mensilita saldate
- proposta automatica del residuo nei pagamenti, modificabile per acconti o saldi parziali
- dashboard con selezione dell'anno di lavoro
- dashboard con riepilogo associati cliccabile e grafici di incassi/scadenze
- viste separate `Inserimento` e `Dati inseriti` accessibili dalle stesse maschere

## File principali

- [database/gestione_associazione.sqlite](C:/Users/mterr/Documents/Codex/2026-04-24/mi-crei-un-database-gi-pronto/database/gestione_associazione.sqlite)
- [database/schema_associazione.sql](C:/Users/mterr/Documents/Codex/2026-04-24/mi-crei-un-database-gi-pronto/database/schema_associazione.sql)
- [database/query_utili.sql](C:/Users/mterr/Documents/Codex/2026-04-24/mi-crei-un-database-gi-pronto/database/query_utili.sql)
- [app.py](C:/Users/mterr/Documents/Codex/2026-04-24/mi-crei-un-database-gi-pronto/app.py)
- [static/style.css](C:/Users/mterr/Documents/Codex/2026-04-24/mi-crei-un-database-gi-pronto/static/style.css)
- [scripts/avvia_gestionale.ps1](C:/Users/mterr/Documents/Codex/2026-04-24/mi-crei-un-database-gi-pronto/scripts/avvia_gestionale.ps1)
- [scripts/crea_database.py](C:/Users/mterr/Documents/Codex/2026-04-24/mi-crei-un-database-gi-pronto/scripts/crea_database.py)

## Struttura

Le aree coperte sono queste:

1. `associati`
   Anagrafica dei soci con dati di contatto e stato.
2. `tesseramenti_annuali` e `pagamenti_tesseramenti`
   Gestione quota associativa annuale e relativi incassi.
3. `tipologie_corsi` e `corsi`
   Creazione di varie tipologie di corsi e catalogo corsi.
4. `iscrizioni_corsi`, `pagamenti_iscrizioni_corsi`, `rate_corsi_mensili`, `pagamenti_rate_corsi`
   Iscrizioni ai corsi e annotazione dei pagamenti mensili, anche multipli e parziali con ricevuta unica.
5. `campi_estivi`, `iscrizioni_campi_estivi`, `pagamenti_campi_estivi`
   Campo estivo con iscrizione diretta del partecipante e quota unica di partecipazione.
6. `eventi`, `iscrizioni_eventi`, `pagamenti_eventi`
   Eventi multipli con registrazione partecipanti e quota una tantum.

## Viste gia pronte

Sono incluse viste utili per lavorare subito:

- `v_tesseramenti_saldo`
- `v_iscrizioni_corsi_saldo`
- `v_rate_corsi_saldo`
- `v_campi_estivi_saldo`
- `v_eventi_saldo`
- `v_scadenze_da_incassare`
- `v_riepilogo_associati`
- `v_incassi_totali`

Queste viste calcolano automaticamente:

- importo dovuto
- importo pagato
- saldo residuo
- stato pagamento (`Pagato`, `Parziale`, `Da pagare`)

Per `campo estivo` ed `eventi` il modello e pensato per quote una tantum, quindi e possibile registrare un solo pagamento per ogni iscrizione.

## Come usarlo

Puoi aprire il file SQLite con strumenti come:

- DB Browser for SQLite
- DBeaver
- SQLiteStudio

## Maschere e report web

E inclusa anche una piccola applicazione web locale con maschere di inserimento e report.

Per avviarla in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\avvia_gestionale.ps1
```

Poi apri il browser su:

```text
http://127.0.0.1:8000
```

Al primo accesso il gestionale chiede di creare l'`utente amministratore`.
Da quel momento:

- l'amministratore puo creare e disattivare utenti standard
- l'amministratore puo reimpostare la password degli utenti standard
- ogni utente puo cambiare la propria password dalla maschera `Profilo accesso`
- l'utente standard puo recuperare la password dalla pagina iniziale usando `username` ed `email di recupero`

Nota: per sicurezza le password non sono mai visualizzabili in chiaro; possono solo essere cambiate o reimpostate.

### Maschere disponibili

- associati
- tesseramenti annuali e pagamenti
- tipologie corsi
- corsi
- iscrizioni ai corsi
- generazione massiva quote mensili per mese
- pagamenti quote mensili corsi filtrati per associato, anche multipli e parziali con ricevuta unica
- campo estivo con iscritti e pagamento quota una tantum
- eventi con partecipanti e pagamento quota una tantum

Ogni area include anche tabelle operative con azioni `Modifica`, `Elimina` e, per i pagamenti, accesso diretto alla `Ricevuta`.
Ogni area include anche due viste:

- `Vista inserimento` per registrare nuovi dati
- `Apri dati ...` per consultare e gestire i dati gia inseriti

### Report disponibili

- riepilogo associati
- situazione tesseramenti
- scadenze da incassare
- quote mensili corsi
- report campo estivo
- report eventi
- report partecipanti con selezione corso, campo estivo o evento
- incassi totali con filtro per intervallo date

Ogni report include:

- pulsante `Esporta Excel` per scaricare il report in formato `.xlsx`
- pulsante `Esporta PDF` per scaricare il report in formato `.pdf`
- pulsante `Stampa report` per la stampa diretta dal browser
- campo `Cerca` in alto a destra per filtrare rapidamente ogni colonna visibile, anche negli export Excel/PDF

Filtri aggiunti:

- report incassi: intervallo, area, riferimento
- report eventi: filtro evento
- report campo estivo: filtro campo estivo
- report quote mensili corsi: filtro corso e intervallo competenza da data a data
- report scadenze, incassi e tesseramenti: supporto al dettaglio del singolo associato

Le ricevute di pagamento includono:

- stampa diretta
- export PDF
- pulsante `Prepara email con PDF`
- pulsante `Invia WhatsApp con PDF` sul numero di cellulare presente nell'anagrafica

Nota operativa:

- se il browser supporta la condivisione file, il PDF viene proposto direttamente in condivisione
- in alternativa il gestionale apre il PDF e prepara email o WhatsApp, cosi puo essere allegato manualmente

## Flusso operativo consigliato

1. Inserisci l'associato in `associati`.
2. Registra il tesseramento annuale in `tesseramenti_annuali`.
3. Inserisci eventuali pagamenti in `pagamenti_tesseramenti`.
4. Crea le tipologie in `tipologie_corsi`.
5. Crea i corsi in `corsi`.
6. Registra le iscrizioni ai corsi in `iscrizioni_corsi`.
7. Crea le rate mensili in `rate_corsi_mensili`.
8. Registra i pagamenti mensili in `pagamenti_rate_corsi`, anche per piu mensilita insieme.
9. Crea il campo estivo, registra gli iscritti e il pagamento della quota unica.
10. Crea gli eventi, registra gli iscritti e il pagamento della quota una tantum.

## Esempi rapidi

Inserire un associato:

```sql
INSERT INTO associati (
    codice_associato, nome, cognome, data_nascita, email, telefono
) VALUES (
    'SOC-001', 'Mario', 'Rossi', '2012-05-14', 'mario@example.com', '3331234567'
);
```

Registrare un tesseramento annuale:

```sql
INSERT INTO tesseramenti_annuali (
    associato_id, anno_sociale, importo_dovuto, data_scadenza
) VALUES (
    1, 2026, 50.00, '2026-09-30'
);
```

Registrare un pagamento del tesseramento:

```sql
INSERT INTO pagamenti_tesseramenti (
    tesseramento_id, data_pagamento, importo, metodo_pagamento_id
) VALUES (
    1, '2026-09-10', 50.00, 2
);
```

Creare una tipologia di corso:

```sql
INSERT INTO tipologie_corsi (
    codice_tipologia, nome, descrizione
) VALUES (
    'DANZA', 'Danza', 'Corsi di danza per varie fasce di eta'
);
```

Creare un corso collegato a una tipologia:

```sql
INSERT INTO corsi (
    codice_corso, nome, tipologia_corso_id, quota_iscrizione_standard, quota_mensile_standard, giorno_settimana, orario
) VALUES (
    'C-DANZA-01',
    'Danza Bambini',
    (SELECT id FROM tipologie_corsi WHERE codice_tipologia = 'DANZA'),
    20.00,
    35.00,
    'Lunedi',
    '17:00-18:00'
);
```

Registrare un iscritto a un corso:

```sql
INSERT INTO iscrizioni_corsi (
    associato_id, corso_id, data_inizio, quota_iscrizione, quota_mensile
) VALUES (
    1,
    (SELECT id FROM corsi WHERE codice_corso = 'C-DANZA-01'),
    '2026-09-01',
    20.00,
    35.00
);
```

Creare una rata mensile di un corso:

```sql
INSERT INTO rate_corsi_mensili (
    iscrizione_corso_id, anno, mese, importo_dovuto, data_scadenza
) VALUES (
    1, 2026, 10, 35.00, '2026-10-05'
);
```

Registrare il pagamento mensile di un corso:

```sql
INSERT INTO pagamenti_rate_corsi (
    rata_corso_id, data_pagamento, importo, metodo_pagamento_id
) VALUES (
    1, '2026-10-03', 35.00, 2
);
```

Creare un campo estivo con quota di partecipazione standard:

```sql
INSERT INTO campi_estivi (
    codice_campo, nome, anno, data_inizio, data_fine, sede, quota_partecipazione_standard
) VALUES (
    'CE-2026', 'Campo Estivo 2026', 2026, '2026-06-15', '2026-07-31', 'Centro Sportivo', 120.00
);
```

Registrare un iscritto al campo estivo:

```sql
INSERT INTO iscrizioni_campi_estivi (
    associato_id, campo_estivo_id, quota_partecipazione
) VALUES (
    1, 1, 120.00
);
```

Registrare il pagamento della quota una tantum del campo estivo:

```sql
INSERT INTO pagamenti_campi_estivi (
    iscrizione_campo_id, data_pagamento, importo, metodo_pagamento_id
) VALUES (
    1, '2026-06-01', 120.00, 2
);
```

Creare un evento con quota di partecipazione una tantum:

```sql
INSERT INTO eventi (
    codice_evento, nome, tipologia, data_evento, luogo, quota_partecipazione_standard
) VALUES (
    'EVT-001', 'Saggio di Fine Anno', 'Spettacolo', '2026-12-20', 'Teatro Comunale', 15.00
);
```

Registrare un partecipante a un evento:

```sql
INSERT INTO iscrizioni_eventi (
    associato_id, evento_id, quota_partecipazione
) VALUES (
    1, 1, 15.00
);
```

Registrare il pagamento una tantum dell'evento:

```sql
INSERT INTO pagamenti_eventi (
    iscrizione_evento_id, data_pagamento, importo, metodo_pagamento_id
) VALUES (
    1, '2026-12-01', 15.00, 3
);
```

Controllare gli insoluti:

```sql
SELECT *
FROM v_scadenze_da_incassare
ORDER BY scadenza, associato;
```

## Nota

Ho scelto SQLite perche e leggero, semplice da portare e subito utilizzabile anche senza server. Se vuoi, nel passo successivo posso trasformarlo anche in:

- file Excel con fogli gia pronti
- applicazione web locale con maschere di inserimento
- versione MySQL/PostgreSQL
- modello con campi aggiuntivi su misura per la tua associazione
