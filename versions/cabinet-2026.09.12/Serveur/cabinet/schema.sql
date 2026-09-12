CREATE TABLE IF NOT EXISTS schema_version(version integer PRIMARY KEY);
INSERT INTO schema_version VALUES (1) ON CONFLICT DO NOTHING;
CREATE TABLE IF NOT EXISTS comptes(
    identifiant text PRIMARY KEY,
    token_sha256 text NOT NULL UNIQUE,
    roles jsonb NOT NULL,
    actif boolean NOT NULL DEFAULT true
);
CREATE TABLE IF NOT EXISTS ressources(
    genre text NOT NULL,
    id text NOT NULL,
    revision integer NOT NULL DEFAULT 1 CHECK (revision > 0),
    donnees jsonb NOT NULL,
    PRIMARY KEY(genre,id)
);
CREATE INDEX IF NOT EXISTS ressources_donnees ON ressources USING gin(donnees);
CREATE TABLE IF NOT EXISTS consultations(
    id text PRIMARY KEY,
    rdv_id text UNIQUE NOT NULL,
    patient_id text NOT NULL,
    etat text NOT NULL CHECK(etat IN ('arrive','encours','publie','traite','annule')),
    proprietaire text,
    donnees jsonb NOT NULL,
    modifie_le timestamptz NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS publications(
    id text PRIMARY KEY,
    consultation_id text NOT NULL REFERENCES consultations(id),
    etat text NOT NULL CHECK(etat IN ('a_traiter','traite','remplacee')),
    donnees jsonb NOT NULL,
    cree_le timestamptz NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS seances(
    id text PRIMARY KEY REFERENCES consultations(id),
    patient_id text NOT NULL,
    empreinte text NOT NULL,
    lignes jsonb NOT NULL,
    imprimee boolean NOT NULL DEFAULT false
);
CREATE TABLE IF NOT EXISTS commandes(
    compte text NOT NULL REFERENCES comptes(identifiant),
    id text NOT NULL,
    empreinte text NOT NULL,
    resultat jsonb NOT NULL,
    cree_le timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY(compte,id)
);
CREATE TABLE IF NOT EXISTS evenements(
    numero bigserial PRIMARY KEY,
    compte text NOT NULL,
    operation text NOT NULL,
    objet text NOT NULL,
    date_evenement timestamptz NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS imports(
    empreinte text PRIMARY KEY,
    rapport jsonb NOT NULL,
    importe_le timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ressources_rdv_date ON ressources ((donnees->>'Date')) WHERE genre='RDV';
CREATE INDEX IF NOT EXISTS consultations_file ON consultations(etat,modifie_le);
CREATE INDEX IF NOT EXISTS consultations_proprietaire ON consultations(proprietaire,etat);
CREATE INDEX IF NOT EXISTS publications_file ON publications(etat,cree_le);
