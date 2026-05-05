PRAGMA foreign_keys = ON;

BEGIN TRANSACTION;

CREATE TABLE IF NOT EXISTS associati (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_progressivo INTEGER UNIQUE,
    codice_associato TEXT NOT NULL UNIQUE,
    nome TEXT NOT NULL,
    cognome TEXT NOT NULL,
    codice_fiscale TEXT UNIQUE,
    data_nascita TEXT,
    sesso TEXT NOT NULL DEFAULT 'M' CHECK (sesso IN ('M', 'F')),
    comune_nascita TEXT,
    provincia_nascita TEXT,
    carica TEXT NOT NULL DEFAULT 'Associato' CHECK (carica IN ('Associato', 'Presidente', 'Vice Presidente', 'Segretario', 'Tesoriere', 'Consigliere', 'Consigliere spirituale')),
    email TEXT,
    telefono TEXT,
    indirizzo TEXT,
    cap TEXT,
    citta TEXT,
    provincia TEXT,
    impiego TEXT,
    data_prima_iscrizione TEXT NOT NULL DEFAULT (date('now')),
    stato_associato TEXT NOT NULL DEFAULT 'Attivo' CHECK (stato_associato IN ('Attivo', 'Sospeso', 'Dimesso')),
    liberatoria_video TEXT NOT NULL DEFAULT 'Si' CHECK (liberatoria_video IN ('Si', 'No')),
    patologie TEXT,
    genitore_tutore_cognome TEXT,
    genitore_tutore_nome TEXT,
    genitore_tutore_cellulare TEXT,
    genitore_tutore_email TEXT,
    genitore_tutore_impiego TEXT,
    genitore_tutore_tipo_documento TEXT NOT NULL DEFAULT 'Carta d''identità' CHECK (genitore_tutore_tipo_documento IN ('Carta d''identità', 'Patente di guida')),
    genitore_tutore_numero_documento TEXT,
    prelievo_altro_genitore_nome TEXT,
    prelievo_altro_genitore_cognome TEXT,
    prelievo_altro_genitore_cellulare TEXT,
    prelievo_altro_genitore_impiego TEXT,
    prelievo_altro_genitore_tipo_documento TEXT NOT NULL DEFAULT 'Carta d''identità' CHECK (prelievo_altro_genitore_tipo_documento IN ('Carta d''identità', 'Patente di guida')),
    prelievo_altro_genitore_numero_documento TEXT,
    prelievo_altra_persona_nome TEXT,
    prelievo_altra_persona_cognome TEXT,
    prelievo_altra_persona_cellulare TEXT,
    prelievo_altra_persona_tipo_documento TEXT NOT NULL DEFAULT 'Carta d''identità' CHECK (prelievo_altra_persona_tipo_documento IN ('Carta d''identità', 'Patente di guida')),
    prelievo_altra_persona_numero_documento TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS metodi_pagamento (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    attivo INTEGER NOT NULL DEFAULT 1 CHECK (attivo IN (0, 1))
);

CREATE TABLE IF NOT EXISTS utenti_accesso (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    password_salt TEXT NOT NULL,
    password_iterations INTEGER NOT NULL DEFAULT 210000,
    email_recupero TEXT,
    is_admin INTEGER NOT NULL DEFAULT 0 CHECK (is_admin IN (0, 1)),
    attivo INTEGER NOT NULL DEFAULT 1 CHECK (attivo IN (0, 1)),
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    ultimo_accesso TEXT
);

CREATE TABLE IF NOT EXISTS sessioni_accesso (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_token TEXT NOT NULL UNIQUE,
    utente_id INTEGER NOT NULL,
    creata_il TEXT NOT NULL DEFAULT (datetime('now')),
    scade_il TEXT NOT NULL,
    FOREIGN KEY (utente_id) REFERENCES utenti_accesso (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS impostazioni_app (
    chiave TEXT PRIMARY KEY,
    valore TEXT,
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS registro_attivita (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data_ora TEXT NOT NULL DEFAULT (datetime('now')),
    username TEXT,
    associato_id INTEGER,
    associato_codice TEXT,
    associato_nominativo TEXT,
    nome_pc TEXT,
    ip_client TEXT,
    metodo_http TEXT,
    percorso TEXT,
    categoria TEXT,
    descrizione_attivita TEXT NOT NULL,
    dettaglio TEXT,
    esito TEXT,
    anno_lavoro INTEGER,
    user_agent TEXT
);

CREATE TABLE IF NOT EXISTS sequenze_progressive (
    chiave TEXT PRIMARY KEY,
    ultimo_valore INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_utenti_accesso_admin_attivo
ON utenti_accesso (is_admin, attivo, username);

CREATE INDEX IF NOT EXISTS idx_sessioni_accesso_utente
ON sessioni_accesso (utente_id, scade_il);

CREATE INDEX IF NOT EXISTS idx_registro_attivita_data
ON registro_attivita (data_ora DESC, id DESC);

CREATE TABLE IF NOT EXISTS quote_predefinite (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    area TEXT NOT NULL CHECK (area IN ('tesseramenti', 'campi-estivi', 'oratorio')),
    descrizione TEXT NOT NULL,
    importo NUMERIC NOT NULL CHECK (importo >= 0),
    attiva INTEGER NOT NULL DEFAULT 1 CHECK (attiva IN (0, 1)),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
);

INSERT OR IGNORE INTO metodi_pagamento (id, nome) VALUES
    (1, 'Contanti'),
    (2, 'Bonifico'),
    (3, 'POS'),
    (4, 'Carta'),
    (5, 'Satispay'),
    (6, 'Altro');

CREATE TABLE IF NOT EXISTS tipologie_corsi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codice_tipologia TEXT NOT NULL UNIQUE,
    nome TEXT NOT NULL UNIQUE,
    descrizione TEXT,
    attiva INTEGER NOT NULL DEFAULT 1 CHECK (attiva IN (0, 1)),
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
);

INSERT OR IGNORE INTO tipologie_corsi (id, codice_tipologia, nome, descrizione) VALUES
    (1, 'SPORT', 'Sportivo', 'Corsi di attivita sportiva'),
    (2, 'ART', 'Artistico', 'Corsi di danza, teatro o arti espressive'),
    (3, 'MUS', 'Musicale', 'Corsi di musica e canto'),
    (4, 'FORM', 'Formativo', 'Corsi educativi, culturali o laboratoriali');

CREATE TABLE IF NOT EXISTS tesseramenti_annuali (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    associato_id INTEGER NOT NULL,
    anno_sociale INTEGER NOT NULL,
    numero_progressivo_anno INTEGER,
    codice_tesseramento TEXT UNIQUE,
    data_tesseramento TEXT NOT NULL DEFAULT (date('now')),
    importo_dovuto NUMERIC NOT NULL CHECK (importo_dovuto >= 0),
    data_scadenza TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (associato_id, anno_sociale),
    FOREIGN KEY (associato_id) REFERENCES associati (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pagamenti_tesseramenti (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tesseramento_id INTEGER NOT NULL,
    data_pagamento TEXT NOT NULL,
    importo NUMERIC NOT NULL CHECK (importo > 0),
    metodo_pagamento_id INTEGER,
    riferimento TEXT,
    gruppo_ricevuta TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (tesseramento_id) REFERENCES tesseramenti_annuali (id) ON DELETE CASCADE,
    FOREIGN KEY (metodo_pagamento_id) REFERENCES metodi_pagamento (id)
);

CREATE TABLE IF NOT EXISTS corsi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_progressivo INTEGER UNIQUE,
    codice_corso TEXT NOT NULL UNIQUE,
    nome TEXT NOT NULL,
    tipologia_corso_id INTEGER,
    descrizione TEXT,
    quota_iscrizione_standard NUMERIC NOT NULL DEFAULT 0 CHECK (quota_iscrizione_standard >= 0),
    quota_mensile_standard NUMERIC NOT NULL DEFAULT 0 CHECK (quota_mensile_standard >= 0),
    data_inizio TEXT,
    data_fine TEXT,
    sede TEXT,
    giorno_settimana TEXT,
    orario TEXT,
    attivo INTEGER NOT NULL DEFAULT 1 CHECK (attivo IN (0, 1)),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (tipologia_corso_id) REFERENCES tipologie_corsi (id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS iscrizioni_corsi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    associato_id INTEGER NOT NULL,
    corso_id INTEGER NOT NULL,
    data_iscrizione TEXT NOT NULL DEFAULT (date('now')),
    data_inizio TEXT,
    data_fine TEXT,
    quota_iscrizione NUMERIC NOT NULL DEFAULT 0 CHECK (quota_iscrizione >= 0),
    quota_mensile NUMERIC NOT NULL CHECK (quota_mensile >= 0),
    stato_iscrizione TEXT NOT NULL DEFAULT 'Attiva' CHECK (stato_iscrizione IN ('Attiva', 'Sospesa', 'Chiusa')),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (associato_id) REFERENCES associati (id) ON DELETE CASCADE,
    FOREIGN KEY (corso_id) REFERENCES corsi (id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS pagamenti_iscrizioni_corsi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    iscrizione_corso_id INTEGER NOT NULL,
    data_pagamento TEXT NOT NULL,
    importo NUMERIC NOT NULL CHECK (importo > 0),
    metodo_pagamento_id INTEGER,
    riferimento TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (iscrizione_corso_id) REFERENCES iscrizioni_corsi (id) ON DELETE CASCADE,
    FOREIGN KEY (metodo_pagamento_id) REFERENCES metodi_pagamento (id)
);

CREATE TABLE IF NOT EXISTS rate_corsi_mensili (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    iscrizione_corso_id INTEGER NOT NULL,
    anno INTEGER NOT NULL,
    mese INTEGER NOT NULL CHECK (mese BETWEEN 1 AND 12),
    importo_dovuto NUMERIC NOT NULL CHECK (importo_dovuto >= 0),
    data_scadenza TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (iscrizione_corso_id, anno, mese),
    FOREIGN KEY (iscrizione_corso_id) REFERENCES iscrizioni_corsi (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pagamenti_rate_corsi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rata_corso_id INTEGER NOT NULL,
    data_pagamento TEXT NOT NULL,
    importo NUMERIC NOT NULL CHECK (importo > 0),
    metodo_pagamento_id INTEGER,
    riferimento TEXT,
    gruppo_ricevuta TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (rata_corso_id) REFERENCES rate_corsi_mensili (id) ON DELETE CASCADE,
    FOREIGN KEY (metodo_pagamento_id) REFERENCES metodi_pagamento (id)
);

CREATE TABLE IF NOT EXISTS campi_estivi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_progressivo INTEGER UNIQUE,
    codice_campo TEXT NOT NULL UNIQUE,
    nome TEXT NOT NULL,
    anno INTEGER NOT NULL,
    data_inizio TEXT NOT NULL,
    data_fine TEXT NOT NULL,
    sede TEXT,
    quota_partecipazione_standard NUMERIC NOT NULL DEFAULT 0 CHECK (quota_partecipazione_standard >= 0),
    posti_massimi INTEGER,
    descrizione TEXT,
    attivo INTEGER NOT NULL DEFAULT 1 CHECK (attivo IN (0, 1)),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS iscrizioni_campi_estivi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    associato_id INTEGER NOT NULL,
    campo_estivo_id INTEGER NOT NULL,
    data_iscrizione TEXT NOT NULL DEFAULT (date('now')),
    quota_partecipazione NUMERIC NOT NULL CHECK (quota_partecipazione >= 0),
    stato_iscrizione TEXT NOT NULL DEFAULT 'Iscritto' CHECK (stato_iscrizione IN ('Iscritto', 'Lista attesa', 'Annullato')),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (associato_id, campo_estivo_id),
    FOREIGN KEY (associato_id) REFERENCES associati (id) ON DELETE CASCADE,
    FOREIGN KEY (campo_estivo_id) REFERENCES campi_estivi (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pagamenti_campi_estivi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    iscrizione_campo_id INTEGER NOT NULL,
    data_pagamento TEXT NOT NULL,
    importo NUMERIC NOT NULL CHECK (importo > 0),
    metodo_pagamento_id INTEGER,
    riferimento TEXT,
    gruppo_ricevuta TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (iscrizione_campo_id),
    FOREIGN KEY (iscrizione_campo_id) REFERENCES iscrizioni_campi_estivi (id) ON DELETE CASCADE,
    FOREIGN KEY (metodo_pagamento_id) REFERENCES metodi_pagamento (id)
);

CREATE TABLE IF NOT EXISTS oratorio (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_progressivo INTEGER UNIQUE,
    codice_oratorio TEXT NOT NULL UNIQUE,
    nome TEXT NOT NULL,
    anno INTEGER NOT NULL,
    data_inizio TEXT NOT NULL,
    data_fine TEXT NOT NULL,
    sede TEXT,
    quota_partecipazione_standard NUMERIC NOT NULL DEFAULT 0 CHECK (quota_partecipazione_standard >= 0),
    posti_massimi INTEGER,
    descrizione TEXT,
    attivo INTEGER NOT NULL DEFAULT 1 CHECK (attivo IN (0, 1)),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS iscrizioni_oratorio (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    associato_id INTEGER NOT NULL,
    oratorio_id INTEGER NOT NULL,
    data_iscrizione TEXT NOT NULL DEFAULT (date('now')),
    quota_partecipazione NUMERIC NOT NULL CHECK (quota_partecipazione >= 0),
    stato_iscrizione TEXT NOT NULL DEFAULT 'Iscritto' CHECK (stato_iscrizione IN ('Iscritto', 'Lista attesa', 'Annullato')),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (associato_id, oratorio_id),
    FOREIGN KEY (associato_id) REFERENCES associati (id) ON DELETE CASCADE,
    FOREIGN KEY (oratorio_id) REFERENCES oratorio (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pagamenti_oratorio (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    iscrizione_oratorio_id INTEGER NOT NULL,
    data_pagamento TEXT NOT NULL,
    importo NUMERIC NOT NULL CHECK (importo > 0),
    metodo_pagamento_id INTEGER,
    riferimento TEXT,
    gruppo_ricevuta TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (iscrizione_oratorio_id),
    FOREIGN KEY (iscrizione_oratorio_id) REFERENCES iscrizioni_oratorio (id) ON DELETE CASCADE,
    FOREIGN KEY (metodo_pagamento_id) REFERENCES metodi_pagamento (id)
);

CREATE TABLE IF NOT EXISTS eventi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_progressivo INTEGER UNIQUE,
    codice_evento TEXT NOT NULL UNIQUE,
    nome TEXT NOT NULL,
    tipologia TEXT,
    data_evento TEXT NOT NULL,
    luogo TEXT,
    quota_partecipazione_standard NUMERIC NOT NULL DEFAULT 0 CHECK (quota_partecipazione_standard >= 0),
    posti_massimi INTEGER,
    descrizione TEXT,
    attivo INTEGER NOT NULL DEFAULT 1 CHECK (attivo IN (0, 1)),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS iscrizioni_eventi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    associato_id INTEGER NOT NULL,
    evento_id INTEGER NOT NULL,
    data_iscrizione TEXT NOT NULL DEFAULT (date('now')),
    quota_partecipazione NUMERIC NOT NULL CHECK (quota_partecipazione >= 0),
    stato_iscrizione TEXT NOT NULL DEFAULT 'Iscritto' CHECK (stato_iscrizione IN ('Iscritto', 'Confermato', 'Annullato')),
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    aggiornato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (associato_id, evento_id),
    FOREIGN KEY (associato_id) REFERENCES associati (id) ON DELETE CASCADE,
    FOREIGN KEY (evento_id) REFERENCES eventi (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS pagamenti_eventi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    iscrizione_evento_id INTEGER NOT NULL,
    data_pagamento TEXT NOT NULL,
    importo NUMERIC NOT NULL CHECK (importo > 0),
    metodo_pagamento_id INTEGER,
    riferimento TEXT,
    gruppo_ricevuta TEXT,
    note TEXT,
    creato_il TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (iscrizione_evento_id),
    FOREIGN KEY (iscrizione_evento_id) REFERENCES iscrizioni_eventi (id) ON DELETE CASCADE,
    FOREIGN KEY (metodo_pagamento_id) REFERENCES metodi_pagamento (id)
);

CREATE INDEX IF NOT EXISTS idx_tesseramenti_associato ON tesseramenti_annuali (associato_id, anno_sociale);
CREATE UNIQUE INDEX IF NOT EXISTS idx_tesseramenti_codice_tesseramento ON tesseramenti_annuali (codice_tesseramento);
CREATE UNIQUE INDEX IF NOT EXISTS idx_tesseramenti_anno_progressivo ON tesseramenti_annuali (anno_sociale, numero_progressivo_anno);
CREATE UNIQUE INDEX IF NOT EXISTS idx_associati_numero_progressivo ON associati (numero_progressivo);
CREATE INDEX IF NOT EXISTS idx_pagamenti_tesseramenti_tesseramento ON pagamenti_tesseramenti (tesseramento_id, data_pagamento);
CREATE INDEX IF NOT EXISTS idx_pagamenti_tesseramenti_gruppo ON pagamenti_tesseramenti (gruppo_ricevuta);
CREATE UNIQUE INDEX IF NOT EXISTS idx_corsi_numero_progressivo ON corsi (numero_progressivo);
CREATE INDEX IF NOT EXISTS idx_corsi_tipologia ON corsi (tipologia_corso_id, nome);
CREATE INDEX IF NOT EXISTS idx_iscrizioni_corsi_associato ON iscrizioni_corsi (associato_id, corso_id);
CREATE INDEX IF NOT EXISTS idx_pagamenti_iscrizioni_corsi_iscrizione ON pagamenti_iscrizioni_corsi (iscrizione_corso_id, data_pagamento);
CREATE INDEX IF NOT EXISTS idx_rate_corsi_iscrizione ON rate_corsi_mensili (iscrizione_corso_id, anno, mese);
CREATE INDEX IF NOT EXISTS idx_pagamenti_rate_corsi_rata ON pagamenti_rate_corsi (rata_corso_id, data_pagamento);
CREATE INDEX IF NOT EXISTS idx_pagamenti_rate_corsi_gruppo ON pagamenti_rate_corsi (gruppo_ricevuta);
CREATE INDEX IF NOT EXISTS idx_quote_predefinite_area ON quote_predefinite (area, attiva, descrizione);
CREATE UNIQUE INDEX IF NOT EXISTS idx_campi_estivi_numero_progressivo ON campi_estivi (numero_progressivo);
CREATE INDEX IF NOT EXISTS idx_iscrizioni_campi_associato ON iscrizioni_campi_estivi (associato_id, campo_estivo_id);
CREATE INDEX IF NOT EXISTS idx_pagamenti_campi_iscrizione ON pagamenti_campi_estivi (iscrizione_campo_id, data_pagamento);
CREATE INDEX IF NOT EXISTS idx_pagamenti_campi_estivi_gruppo ON pagamenti_campi_estivi (gruppo_ricevuta);
CREATE UNIQUE INDEX IF NOT EXISTS idx_eventi_numero_progressivo ON eventi (numero_progressivo);
CREATE INDEX IF NOT EXISTS idx_iscrizioni_eventi_associato ON iscrizioni_eventi (associato_id, evento_id);
CREATE INDEX IF NOT EXISTS idx_pagamenti_eventi_iscrizione ON pagamenti_eventi (iscrizione_evento_id, data_pagamento);
CREATE INDEX IF NOT EXISTS idx_pagamenti_eventi_gruppo ON pagamenti_eventi (gruppo_ricevuta);

CREATE TRIGGER IF NOT EXISTS trg_associati_aggiornato
AFTER UPDATE ON associati
FOR EACH ROW
BEGIN
    UPDATE associati
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_tesseramenti_aggiornato
AFTER UPDATE ON tesseramenti_annuali
FOR EACH ROW
BEGIN
    UPDATE tesseramenti_annuali
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_tipologie_corsi_aggiornato
AFTER UPDATE ON tipologie_corsi
FOR EACH ROW
BEGIN
    UPDATE tipologie_corsi
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_corsi_aggiornato
AFTER UPDATE ON corsi
FOR EACH ROW
BEGIN
    UPDATE corsi
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_iscrizioni_corsi_aggiornato
AFTER UPDATE ON iscrizioni_corsi
FOR EACH ROW
BEGIN
    UPDATE iscrizioni_corsi
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_rate_corsi_aggiornato
AFTER UPDATE ON rate_corsi_mensili
FOR EACH ROW
BEGIN
    UPDATE rate_corsi_mensili
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_campi_estivi_aggiornato
AFTER UPDATE ON campi_estivi
FOR EACH ROW
BEGIN
    UPDATE campi_estivi
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_iscrizioni_campi_aggiornato
AFTER UPDATE ON iscrizioni_campi_estivi
FOR EACH ROW
BEGIN
    UPDATE iscrizioni_campi_estivi
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_eventi_aggiornato
AFTER UPDATE ON eventi
FOR EACH ROW
BEGIN
    UPDATE eventi
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_iscrizioni_eventi_aggiornato
AFTER UPDATE ON iscrizioni_eventi
FOR EACH ROW
BEGIN
    UPDATE iscrizioni_eventi
    SET aggiornato_il = datetime('now')
    WHERE id = NEW.id;
END;

CREATE VIEW IF NOT EXISTS v_tesseramenti_saldo AS
SELECT
    t.id,
    t.anno_sociale,
    t.codice_tesseramento,
    a.id AS associato_id,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    t.data_tesseramento,
    t.data_scadenza,
    t.importo_dovuto,
    COALESCE(SUM(pt.importo), 0) AS importo_pagato,
    t.importo_dovuto - COALESCE(SUM(pt.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(pt.importo), 0) >= t.importo_dovuto THEN 'Pagato'
        WHEN COALESCE(SUM(pt.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM tesseramenti_annuali t
JOIN associati a ON a.id = t.associato_id
LEFT JOIN pagamenti_tesseramenti pt ON pt.tesseramento_id = t.id
GROUP BY t.id;

CREATE VIEW IF NOT EXISTS v_iscrizioni_corsi_saldo AS
SELECT
    ic.id,
    a.id AS associato_id,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    tc.codice_tipologia,
    COALESCE(tc.nome, 'Senza tipologia') AS tipologia_corso,
    c.codice_corso,
    c.nome AS corso,
    ic.data_iscrizione,
    ic.data_inizio,
    ic.data_fine,
    ic.stato_iscrizione,
    ic.quota_iscrizione,
    COALESCE(SUM(pic.importo), 0) AS importo_pagato_iscrizione,
    ic.quota_iscrizione - COALESCE(SUM(pic.importo), 0) AS saldo_iscrizione,
    CASE
        WHEN COALESCE(SUM(pic.importo), 0) >= ic.quota_iscrizione THEN 'Pagato'
        WHEN COALESCE(SUM(pic.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento_iscrizione,
    ic.quota_mensile
FROM iscrizioni_corsi ic
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN tipologie_corsi tc ON tc.id = c.tipologia_corso_id
LEFT JOIN pagamenti_iscrizioni_corsi pic ON pic.iscrizione_corso_id = ic.id
GROUP BY ic.id;

CREATE VIEW IF NOT EXISTS v_rate_corsi_saldo AS
SELECT
    r.id,
    a.id AS associato_id,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    tc.codice_tipologia,
    COALESCE(tc.nome, 'Senza tipologia') AS tipologia_corso,
    c.codice_corso,
    c.nome AS corso,
    r.anno,
    r.mese,
    printf('%04d-%02d', r.anno, r.mese) AS competenza,
    r.data_scadenza,
    r.importo_dovuto,
    COALESCE(SUM(prc.importo), 0) AS importo_pagato,
    r.importo_dovuto - COALESCE(SUM(prc.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(prc.importo), 0) >= r.importo_dovuto THEN 'Pagato'
        WHEN COALESCE(SUM(prc.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM rate_corsi_mensili r
JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN tipologie_corsi tc ON tc.id = c.tipologia_corso_id
LEFT JOIN pagamenti_rate_corsi prc ON prc.rata_corso_id = r.id
GROUP BY r.id;

CREATE VIEW IF NOT EXISTS v_campi_estivi_saldo AS
SELECT
    ice.id,
    a.id AS associato_id,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    ce.codice_campo,
    ce.nome AS campo_estivo,
    ce.anno,
    ce.data_inizio,
    ce.data_fine,
    ice.data_iscrizione,
    ice.stato_iscrizione,
    ice.quota_partecipazione AS importo_dovuto,
    COALESCE(SUM(pce.importo), 0) AS importo_pagato,
    ice.quota_partecipazione - COALESCE(SUM(pce.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(pce.importo), 0) >= ice.quota_partecipazione THEN 'Pagato'
        WHEN COALESCE(SUM(pce.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM iscrizioni_campi_estivi ice
JOIN associati a ON a.id = ice.associato_id
JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
LEFT JOIN pagamenti_campi_estivi pce ON pce.iscrizione_campo_id = ice.id
GROUP BY ice.id;

CREATE VIEW IF NOT EXISTS v_oratorio_saldo AS
SELECT
    io.id,
    a.id AS associato_id,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    o.codice_oratorio,
    o.nome AS oratorio,
    o.anno,
    o.data_inizio,
    o.data_fine,
    io.data_iscrizione,
    io.stato_iscrizione,
    io.quota_partecipazione AS importo_dovuto,
    COALESCE(SUM(po.importo), 0) AS importo_pagato,
    io.quota_partecipazione - COALESCE(SUM(po.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(po.importo), 0) >= io.quota_partecipazione THEN 'Pagato'
        WHEN COALESCE(SUM(po.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM iscrizioni_oratorio io
JOIN associati a ON a.id = io.associato_id
JOIN oratorio o ON o.id = io.oratorio_id
LEFT JOIN pagamenti_oratorio po ON po.iscrizione_oratorio_id = io.id
GROUP BY io.id;

CREATE VIEW IF NOT EXISTS v_eventi_saldo AS
SELECT
    ie.id,
    a.id AS associato_id,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    e.codice_evento,
    e.nome AS evento,
    e.tipologia,
    e.data_evento,
    ie.data_iscrizione,
    ie.stato_iscrizione,
    ie.quota_partecipazione AS importo_dovuto,
    COALESCE(SUM(pe.importo), 0) AS importo_pagato,
    ie.quota_partecipazione - COALESCE(SUM(pe.importo), 0) AS saldo_residuo,
    CASE
        WHEN COALESCE(SUM(pe.importo), 0) >= ie.quota_partecipazione THEN 'Pagato'
        WHEN COALESCE(SUM(pe.importo), 0) > 0 THEN 'Parziale'
        ELSE 'Da pagare'
    END AS stato_pagamento
FROM iscrizioni_eventi ie
JOIN associati a ON a.id = ie.associato_id
JOIN eventi e ON e.id = ie.evento_id
LEFT JOIN pagamenti_eventi pe ON pe.iscrizione_evento_id = ie.id
GROUP BY ie.id;

CREATE VIEW IF NOT EXISTS v_scadenze_da_incassare AS
SELECT
    associato_id,
    codice_associato,
    associato,
    'Tesseramento annuale' AS area,
    'Anno ' || anno_sociale AS riferimento,
    data_scadenza AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_tesseramenti_saldo
WHERE saldo_residuo > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    'Corso - iscrizione' AS area,
    corso AS riferimento,
    data_iscrizione AS scadenza,
    quota_iscrizione AS importo_dovuto,
    importo_pagato_iscrizione AS importo_pagato,
    saldo_iscrizione AS saldo_residuo,
    stato_pagamento_iscrizione AS stato_pagamento
FROM v_iscrizioni_corsi_saldo
WHERE saldo_iscrizione > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    'Corso - quota mensile' AS area,
    corso || ' ' || competenza AS riferimento,
    data_scadenza AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_rate_corsi_saldo
WHERE saldo_residuo > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    'Campo estivo' AS area,
    campo_estivo AS riferimento,
    data_inizio AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_campi_estivi_saldo
WHERE saldo_residuo > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    'Oratorio' AS area,
    oratorio AS riferimento,
    data_inizio AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_oratorio_saldo
WHERE saldo_residuo > 0

UNION ALL

SELECT
    associato_id,
    codice_associato,
    associato,
    'Evento' AS area,
    evento AS riferimento,
    data_evento AS scadenza,
    importo_dovuto,
    importo_pagato,
    saldo_residuo,
    stato_pagamento
FROM v_eventi_saldo
WHERE saldo_residuo > 0;

CREATE VIEW IF NOT EXISTS v_riepilogo_associati AS
WITH movimenti AS (
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_tesseramenti_saldo
    UNION ALL
    SELECT associato_id, quota_iscrizione AS importo_dovuto, importo_pagato_iscrizione AS importo_pagato
    FROM v_iscrizioni_corsi_saldo
    UNION ALL
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_rate_corsi_saldo
    UNION ALL
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_campi_estivi_saldo
    UNION ALL
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_oratorio_saldo
    UNION ALL
    SELECT associato_id, importo_dovuto, importo_pagato
    FROM v_eventi_saldo
)
SELECT
    a.id AS associato_id,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    a.stato_associato,
    COALESCE(SUM(m.importo_dovuto), 0) AS totale_dovuto,
    COALESCE(SUM(m.importo_pagato), 0) AS totale_pagato,
    COALESCE(SUM(m.importo_dovuto), 0) - COALESCE(SUM(m.importo_pagato), 0) AS saldo_residuo
FROM associati a
LEFT JOIN movimenti m ON m.associato_id = a.id
GROUP BY a.id;

CREATE VIEW IF NOT EXISTS v_incassi_totali AS
SELECT
    'Tesseramento annuale' AS area,
    pt.data_pagamento,
    pt.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    'Anno ' || t.anno_sociale AS riferimento
FROM pagamenti_tesseramenti pt
JOIN tesseramenti_annuali t ON t.id = pt.tesseramento_id
JOIN associati a ON a.id = t.associato_id
LEFT JOIN metodi_pagamento mp ON mp.id = pt.metodo_pagamento_id

UNION ALL

SELECT
    'Corso - iscrizione' AS area,
    pic.data_pagamento,
    pic.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    c.nome AS riferimento
FROM pagamenti_iscrizioni_corsi pic
JOIN iscrizioni_corsi ic ON ic.id = pic.iscrizione_corso_id
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN metodi_pagamento mp ON mp.id = pic.metodo_pagamento_id

UNION ALL

SELECT
    'Corso - quota mensile' AS area,
    prc.data_pagamento,
    prc.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    c.nome || ' ' || printf('%04d-%02d', r.anno, r.mese) AS riferimento
FROM pagamenti_rate_corsi prc
JOIN rate_corsi_mensili r ON r.id = prc.rata_corso_id
JOIN iscrizioni_corsi ic ON ic.id = r.iscrizione_corso_id
JOIN associati a ON a.id = ic.associato_id
JOIN corsi c ON c.id = ic.corso_id
LEFT JOIN metodi_pagamento mp ON mp.id = prc.metodo_pagamento_id

UNION ALL

SELECT
    'Campo estivo' AS area,
    pce.data_pagamento,
    pce.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    ce.nome AS riferimento
FROM pagamenti_campi_estivi pce
JOIN iscrizioni_campi_estivi ice ON ice.id = pce.iscrizione_campo_id
JOIN associati a ON a.id = ice.associato_id
JOIN campi_estivi ce ON ce.id = ice.campo_estivo_id
LEFT JOIN metodi_pagamento mp ON mp.id = pce.metodo_pagamento_id

UNION ALL

SELECT
    'Oratorio' AS area,
    po.data_pagamento,
    po.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    o.nome AS riferimento
FROM pagamenti_oratorio po
JOIN iscrizioni_oratorio io ON io.id = po.iscrizione_oratorio_id
JOIN associati a ON a.id = io.associato_id
JOIN oratorio o ON o.id = io.oratorio_id
LEFT JOIN metodi_pagamento mp ON mp.id = po.metodo_pagamento_id

UNION ALL

SELECT
    'Evento' AS area,
    pe.data_pagamento,
    pe.importo,
    mp.nome AS metodo_pagamento,
    a.codice_associato,
    a.cognome || ' ' || a.nome AS associato,
    e.nome AS riferimento
FROM pagamenti_eventi pe
JOIN iscrizioni_eventi ie ON ie.id = pe.iscrizione_evento_id
JOIN associati a ON a.id = ie.associato_id
JOIN eventi e ON e.id = ie.evento_id
LEFT JOIN metodi_pagamento mp ON mp.id = pe.metodo_pagamento_id;

COMMIT;
