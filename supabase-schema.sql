-- Methodencurriculum-App: Datenbank-Struktur für Supabase
-- Im Supabase-Dashboard unter "SQL Editor" -> "New query" einfügen und ausführen.

create extension if not exists pgcrypto;

-- Fächer
create table faecher (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  kuerzel text not null unique,
  quelle text
);

-- Lehrkräfte
create table lehrer (
  id uuid primary key default gen_random_uuid(),
  name text,
  kuerzel text not null unique,
  email text,
  quelle text
);

-- Klassen
create table klassen (
  id uuid primary key default gen_random_uuid(),
  jahrgang int not null,
  buchstabe text not null,
  lehrer1_id uuid references lehrer(id) on delete set null,
  lehrer2_id uuid references lehrer(id) on delete set null,
  unique (jahrgang, buchstabe)
);

-- Lerngruppen (Fach + Lehrer + eine oder mehrere Klassen)
create table lerngruppen (
  id uuid primary key default gen_random_uuid(),
  fach_id uuid not null references faecher(id) on delete cascade,
  bezeichnung text,
  jahrgang int,
  lehrer_id uuid references lehrer(id) on delete set null,
  klassen_ids uuid[] not null default '{}',
  quelle text
);

-- Methoden
create table methoden (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  beschreibung text default '',
  jahrgaenge int[] default '{}',
  faecher_ids uuid[] default '{}',
  halbjahr int,
  materialien jsonb default '[]',
  links jsonb default '[]'
);

-- Planungen (Zuordnung Methode <-> Lerngruppe/Klasse/Quartal + Durchführungsstatus)
create table planungen (
  id uuid primary key default gen_random_uuid(),
  methode_id uuid not null references methoden(id) on delete cascade,
  lerngruppe_id uuid not null references lerngruppen(id) on delete cascade,
  klasse_id uuid references klassen(id) on delete cascade,
  quartal int not null,
  status text not null default 'ausstehend',
  datum date,
  notiz text default ''
);

-- Row Level Security aktivieren (Pflicht bei Supabase, sonst ist eine Tabelle für den
-- öffentlichen anon-Schlüssel standardmäßig unsichtbar)
alter table faecher enable row level security;
alter table lehrer enable row level security;
alter table klassen enable row level security;
alter table lerngruppen enable row level security;
alter table methoden enable row level security;
alter table planungen enable row level security;

-- VORLÄUFIGE, OFFENE Regeln: erlauben erstmal alles, für jeden mit dem (öffentlichen)
-- anon-Schlüssel - siehe Hinweis im Chat, das ist nur ein erster Arbeitsstand!
create policy "vorlaeufig_alles_erlaubt" on faecher for all using (true) with check (true);
create policy "vorlaeufig_alles_erlaubt" on lehrer for all using (true) with check (true);
create policy "vorlaeufig_alles_erlaubt" on klassen for all using (true) with check (true);
create policy "vorlaeufig_alles_erlaubt" on lerngruppen for all using (true) with check (true);
create policy "vorlaeufig_alles_erlaubt" on methoden for all using (true) with check (true);
create policy "vorlaeufig_alles_erlaubt" on planungen for all using (true) with check (true);
