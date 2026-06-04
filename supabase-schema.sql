-- ─────────────────────────────────────────────────────────────────────────
-- FCP Intel — Supabase Schema
-- Run this in your Supabase project: SQL Editor → New Query → Run
-- ─────────────────────────────────────────────────────────────────────────

create table if not exists fcp_enrichments (
  -- Identity
  code              text primary key,
  name              text not null,
  city              text,

  -- Financial
  ca_montant        text,
  tranche_ca        text,
  ca_export         text,
  part_export       text,

  -- Workforce
  effectif_tranche  text,

  -- Management
  dirigeant         text,
  email_dirigeant   text,
  mobile_dirigeant  text,
  site_web          text,

  -- Sectors
  sectors_checked       text[]  default '{}',
  filieres_confirmees   text[]  default '{}',
  filieres_rejetees     text[]  default '{}',
  filieres_ajoutees     text[]  default '{}',
  filiere_principale    text,
  secteur_reel          text,

  -- Classification
  type_entreprise   text,

  -- Metadata
  confiance         float   default 0,
  raison            text,
  sources           text[]  default '{}',
  processed_at      timestamptz,
  updated_at        timestamptz default now()
);

-- Auto-update updated_at on any row update
create or replace function update_updated_at_column()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists fcp_enrichments_updated_at on fcp_enrichments;
create trigger fcp_enrichments_updated_at
  before update on fcp_enrichments
  for each row execute function update_updated_at_column();

-- Indexes for common query patterns
create index if not exists idx_fcp_filiere     on fcp_enrichments(filiere_principale);
create index if not exists idx_fcp_type        on fcp_enrichments(type_entreprise);
create index if not exists idx_fcp_updated     on fcp_enrichments(updated_at desc);
create index if not exists idx_fcp_confiance   on fcp_enrichments(confiance desc);

-- Row Level Security (allows anon key access — adjust if you need auth)
alter table fcp_enrichments enable row level security;

create policy "anon_read_write" on fcp_enrichments
  for all
  to anon
  using (true)
  with check (true);
