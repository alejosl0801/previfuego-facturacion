-- ══════════════════════════════════════════════════════════════════════════════
-- PREVIFUEGO — Supabase Schema
-- Ejecuta este SQL en el SQL Editor de tu proyecto Supabase (supabase.com)
-- ══════════════════════════════════════════════════════════════════════════════

-- Tabla principal: un registro por local (210 locales)
CREATE TABLE IF NOT EXISTS locales (
  codigo         TEXT PRIMARY KEY,          -- K002, M058, BS004, KN001, etc. (normalizado)
  marca          TEXT,                      -- KFC, MENESTRAS DEL NEGRO, TROPIBURGER, etc.
  nombre_local   TEXT,                      -- K002 - GUAYAQUIL - CC SAN MARINO
  mes_servicio   TEXT,                      -- ENERO, FEBRERO, … DICIEMBRE
  n_extintores   INTEGER DEFAULT 0,
  total_mantt    NUMERIC(10,2) DEFAULT 0,
  total_recarga  NUMERIC(10,2) DEFAULT 0,
  cobro_anual    NUMERIC(10,2) DEFAULT 0,
  anio_ult_recarga INTEGER,
  anio_recarga     INTEGER,
  created_at     TIMESTAMPTZ DEFAULT NOW(),
  updated_at     TIMESTAMPTZ DEFAULT NOW()
);

-- Detalle: un registro por extintor (968+ extintores)
CREATE TABLE IF NOT EXISTS extintores (
  id              SERIAL PRIMARY KEY,
  codigo_local    TEXT NOT NULL REFERENCES locales(codigo) ON DELETE CASCADE,
  marca           TEXT,
  nombre_local    TEXT,                     -- texto original del Excel
  mes_servicio    TEXT,
  ubicacion       TEXT,
  tipo            TEXT,                     -- PQS, CO2, TIPO K
  capacidad       TEXT,                     -- 5 LBS, 10 LBS, 2.5 GLS, etc.
  costo_mantt     NUMERIC(10,2) DEFAULT 0,
  precio_recarga  NUMERIC(10,2) DEFAULT 0,
  anio_ult_recarga INTEGER,
  anio_recarga     INTEGER
);

CREATE INDEX IF NOT EXISTS idx_extintores_codigo ON extintores(codigo_local);
CREATE INDEX IF NOT EXISTS idx_extintores_mes    ON extintores(mes_servicio);
CREATE INDEX IF NOT EXISTS idx_locales_mes       ON locales(mes_servicio);

-- ── Row Level Security ───────────────────────────────────────────────────────
-- anon key = solo lectura; service_role bypasses RLS automáticamente
ALTER TABLE locales    ENABLE ROW LEVEL SECURITY;
ALTER TABLE extintores ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Lectura pública locales"    ON locales    FOR SELECT USING (true);
CREATE POLICY "Lectura pública extintores" ON extintores FOR SELECT USING (true);

-- ── Trigger para updated_at ──────────────────────────────────────────────────
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER set_updated_at
  BEFORE UPDATE ON locales
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
