-- Elenco associati con saldo totale residuo.
SELECT *
FROM v_riepilogo_associati
ORDER BY associato;

-- Scadenze e insoluti da incassare.
SELECT *
FROM v_scadenze_da_incassare
ORDER BY scadenza, associato, area;

-- Tesseramenti dell'anno corrente.
SELECT *
FROM v_tesseramenti_saldo
WHERE anno_sociale = CAST(strftime('%Y', 'now') AS INTEGER)
ORDER BY associato;

-- Rate corsi non ancora saldate.
SELECT *
FROM v_rate_corsi_saldo
WHERE saldo_residuo > 0
ORDER BY anno, mese, associato, corso;

-- Elenco corsi per tipologia.
SELECT
    tc.codice_tipologia,
    tc.nome AS tipologia_corso,
    c.codice_corso,
    c.nome AS corso,
    c.quota_iscrizione_standard,
    c.quota_mensile_standard,
    c.giorno_settimana,
    c.orario
FROM corsi c
LEFT JOIN tipologie_corsi tc ON tc.id = c.tipologia_corso_id
ORDER BY tipologia_corso, corso;

-- Incassi registrati in un intervallo.
SELECT *
FROM v_incassi_totali
WHERE data_pagamento BETWEEN '2026-01-01' AND '2026-12-31'
ORDER BY data_pagamento, area, associato;

-- Iscritti al campo estivo con saldo quota una tantum.
SELECT
    codice_campo,
    campo_estivo,
    data_inizio,
    data_fine,
    codice_associato,
    associato,
    stato_iscrizione,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_campi_estivi_saldo
WHERE codice_campo = 'CE-2026'
ORDER BY associato;

-- Partecipanti iscritti a un evento con quota una tantum.
SELECT
    codice_evento,
    evento,
    data_evento,
    codice_associato,
    associato,
    stato_iscrizione,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_eventi_saldo
WHERE codice_evento = 'EVT-001'
ORDER BY associato;

-- Partecipanti iscritti a un evento specifico.
SELECT
    e.nome AS evento,
    e.data_evento,
    a.codice_associato,
    a.cognome,
    a.nome,
    ie.stato_iscrizione,
    ie.quota_partecipazione
FROM iscrizioni_eventi ie
JOIN eventi e ON e.id = ie.evento_id
JOIN associati a ON a.id = ie.associato_id
WHERE e.codice_evento = 'EVT-001'
ORDER BY a.cognome, a.nome;
