import { supabase } from "./supabaseClient";

// ---------------------------------------------------------------------------
// Umwandlung zwischen Datenbank-Feldnamen (snake_case, wie in supabase-schema.sql)
// und den im Programm verwendeten Feldnamen (camelCase). Fächer und Lehrer brauchen
// keine Umwandlung, dort stimmen die Feldnamen bereits überein.
// ---------------------------------------------------------------------------
const klasseAusDb = (r) => ({ id: r.id, jahrgang: r.jahrgang, buchstabe: r.buchstabe, lehrer1: r.lehrer1_id, lehrer2: r.lehrer2_id, vorgaenger: null });
const klasseZuDb = (k) => ({ jahrgang: k.jahrgang, buchstabe: k.buchstabe, lehrer1_id: k.lehrer1 || null, lehrer2_id: k.lehrer2 || null });

const lerngruppeAusDb = (r) => ({ id: r.id, fachId: r.fach_id, bezeichnung: r.bezeichnung, jahrgang: r.jahrgang, lehrerId: r.lehrer_id, klassenIds: r.klassen_ids || [], quelle: r.quelle });
const lerngruppeZuDb = (g) => ({ fach_id: g.fachId, bezeichnung: g.bezeichnung, jahrgang: g.jahrgang, lehrer_id: g.lehrerId || null, klassen_ids: g.klassenIds || [], quelle: g.quelle || null });

const methodeAusDb = (r) => ({ id: r.id, name: r.name, beschreibung: r.beschreibung || "", jahrgaenge: r.jahrgaenge || [], faecherIds: r.faecher_ids || [], halbjahr: r.halbjahr, materialien: r.materialien || [], links: r.links || [] });
const methodeZuDb = (m) => ({ name: m.name, beschreibung: m.beschreibung || "", jahrgaenge: m.jahrgaenge || [], faecher_ids: m.faecherIds || [], halbjahr: m.halbjahr, materialien: m.materialien || [], links: m.links || [] });

const planungAusDb = (r) => ({ id: r.id, methodeId: r.methode_id, lerngruppeId: r.lerngruppe_id, klasseId: r.klasse_id, quartal: r.quartal, status: r.status, datum: r.datum, notiz: r.notiz || "" });
const planungZuDb = (p) => ({ methode_id: p.methodeId, lerngruppe_id: p.lerngruppeId, klasse_id: p.klasseId || null, quartal: p.quartal, status: p.status, datum: p.datum || null, notiz: p.notiz || "" });

// ---------------------------------------------------------------------------
// Laden
// ---------------------------------------------------------------------------
export async function ladeAlleDaten() {
  const [faecher, lehrer, klassen, lerngruppen, methoden, planungen] = await Promise.all([
    supabase.from("faecher").select("*").order("name"),
    supabase.from("lehrer").select("*").order("kuerzel"),
    supabase.from("klassen").select("*"),
    supabase.from("lerngruppen").select("*"),
    supabase.from("methoden").select("*").order("name"),
    supabase.from("planungen").select("*"),
  ]);
  const fehlerhaft = [faecher, lehrer, klassen, lerngruppen, methoden, planungen].find((r) => r.error);
  if (fehlerhaft) throw fehlerhaft.error;
  return {
    faecher: faecher.data,
    lehrer: lehrer.data,
    klassen: klassen.data.map(klasseAusDb),
    lerngruppen: lerngruppen.data.map(lerngruppeAusDb),
    methoden: methoden.data.map(methodeAusDb),
    planungen: planungen.data.map(planungAusDb),
  };
}

// ---------------------------------------------------------------------------
// Fächer
// ---------------------------------------------------------------------------
export async function fachErstellen(fach) {
  const { data, error } = await supabase.from("faecher").insert({ name: fach.name, kuerzel: fach.kuerzel, quelle: fach.quelle || null }).select().single();
  if (error) throw error;
  return data;
}
export async function fachAktualisieren(id, patch) {
  const dbPatch = {};
  if ("name" in patch) dbPatch.name = patch.name;
  if ("kuerzel" in patch) dbPatch.kuerzel = patch.kuerzel;
  const { error } = await supabase.from("faecher").update(dbPatch).eq("id", id);
  if (error) throw error;
}
export async function fachLoeschen(id) {
  const { error } = await supabase.from("faecher").delete().eq("id", id);
  if (error) throw error;
}

// ---------------------------------------------------------------------------
// Lehrer
// ---------------------------------------------------------------------------
export async function lehrerErstellen(lehrerObj) {
  const { data, error } = await supabase
    .from("lehrer")
    .insert({ name: lehrerObj.name || "", kuerzel: lehrerObj.kuerzel, email: lehrerObj.email || "", quelle: lehrerObj.quelle || null })
    .select()
    .single();
  if (error) throw error;
  return data;
}
export async function lehrerAktualisieren(id, patch) {
  const dbPatch = {};
  if ("name" in patch) dbPatch.name = patch.name;
  if ("kuerzel" in patch) dbPatch.kuerzel = patch.kuerzel;
  if ("email" in patch) dbPatch.email = patch.email;
  const { error } = await supabase.from("lehrer").update(dbPatch).eq("id", id);
  if (error) throw error;
}
export async function lehrerLoeschen(id) {
  const { error } = await supabase.from("lehrer").delete().eq("id", id);
  if (error) throw error;
}

// ---------------------------------------------------------------------------
// Klassen
// ---------------------------------------------------------------------------
export async function klasseErstellen(klasse) {
  const { data, error } = await supabase.from("klassen").insert(klasseZuDb(klasse)).select().single();
  if (error) throw error;
  return klasseAusDb(data);
}
export async function klasseAktualisieren(id, patch) {
  const dbPatch = {};
  if ("lehrer1" in patch) dbPatch.lehrer1_id = patch.lehrer1 || null;
  if ("lehrer2" in patch) dbPatch.lehrer2_id = patch.lehrer2 || null;
  if ("jahrgang" in patch) dbPatch.jahrgang = patch.jahrgang;
  if ("buchstabe" in patch) dbPatch.buchstabe = patch.buchstabe;
  const { error } = await supabase.from("klassen").update(dbPatch).eq("id", id);
  if (error) throw error;
}

// ---------------------------------------------------------------------------
// Lerngruppen
// ---------------------------------------------------------------------------
export async function lerngruppeErstellen(gruppe) {
  const { data, error } = await supabase.from("lerngruppen").insert(lerngruppeZuDb(gruppe)).select().single();
  if (error) throw error;
  return lerngruppeAusDb(data);
}
export async function lerngruppeAktualisieren(id, patch) {
  const dbPatch = {};
  if ("fachId" in patch) dbPatch.fach_id = patch.fachId;
  if ("bezeichnung" in patch) dbPatch.bezeichnung = patch.bezeichnung;
  if ("jahrgang" in patch) dbPatch.jahrgang = patch.jahrgang;
  if ("lehrerId" in patch) dbPatch.lehrer_id = patch.lehrerId;
  if ("klassenIds" in patch) dbPatch.klassen_ids = patch.klassenIds;
  const { error } = await supabase.from("lerngruppen").update(dbPatch).eq("id", id);
  if (error) throw error;
}
export async function lerngruppeLoeschen(id) {
  const { error } = await supabase.from("lerngruppen").delete().eq("id", id);
  if (error) throw error;
}

// ---------------------------------------------------------------------------
// Methoden
// ---------------------------------------------------------------------------
export async function methodeErstellen(methode) {
  const { data, error } = await supabase.from("methoden").insert(methodeZuDb(methode)).select().single();
  if (error) throw error;
  return methodeAusDb(data);
}
export async function methodeAktualisieren(id, patch) {
  const dbPatch = {};
  if ("name" in patch) dbPatch.name = patch.name;
  if ("beschreibung" in patch) dbPatch.beschreibung = patch.beschreibung;
  if ("jahrgaenge" in patch) dbPatch.jahrgaenge = patch.jahrgaenge;
  if ("faecherIds" in patch) dbPatch.faecher_ids = patch.faecherIds;
  if ("halbjahr" in patch) dbPatch.halbjahr = patch.halbjahr;
  if ("materialien" in patch) dbPatch.materialien = patch.materialien;
  if ("links" in patch) dbPatch.links = patch.links;
  const { error } = await supabase.from("methoden").update(dbPatch).eq("id", id);
  if (error) throw error;
}
export async function methodeLoeschen(id) {
  const { error } = await supabase.from("methoden").delete().eq("id", id);
  if (error) throw error;
}

// ---------------------------------------------------------------------------
// Planungen
// ---------------------------------------------------------------------------
export async function planungErstellen(planung) {
  const { data, error } = await supabase.from("planungen").insert(planungZuDb(planung)).select().single();
  if (error) throw error;
  return planungAusDb(data);
}
export async function planungAktualisieren(id, patch) {
  const dbPatch = {};
  if ("lerngruppeId" in patch) dbPatch.lerngruppe_id = patch.lerngruppeId;
  if ("quartal" in patch) dbPatch.quartal = patch.quartal;
  if ("status" in patch) dbPatch.status = patch.status;
  if ("datum" in patch) dbPatch.datum = patch.datum;
  if ("notiz" in patch) dbPatch.notiz = patch.notiz;
  const { error } = await supabase.from("planungen").update(dbPatch).eq("id", id);
  if (error) throw error;
}
export async function planungLoeschen(id) {
  const { error } = await supabase.from("planungen").delete().eq("id", id);
  if (error) throw error;
}

// ---------------------------------------------------------------------------
// Einmalige Migration: kompletten aktuellen (bisher nur im Browser gehaltenen) Stand
// nach Supabase hochladen. Reihenfolge wichtig wegen Fremdschlüsseln. Alte, im Browser
// erzeugte IDs werden dabei auf die von Supabase neu vergebenen IDs übersetzt.
// ---------------------------------------------------------------------------
export async function allesNachSupabaseHochladen({ faecher, lehrer, klassen, lerngruppen, methoden, planungen }, fortschritt) {
  const melde = (text) => fortschritt && fortschritt(text);
  const fachIdKarte = {};
  melde(`Fächer (0/${faecher.length})…`);
  for (const [i, f] of faecher.entries()) {
    const neu = await fachErstellen(f);
    fachIdKarte[f.id] = neu.id;
    melde(`Fächer (${i + 1}/${faecher.length})…`);
  }
  const lehrerIdKarte = {};
  melde(`Lehrkräfte (0/${lehrer.length})…`);
  for (const [i, l] of lehrer.entries()) {
    const neu = await lehrerErstellen(l);
    lehrerIdKarte[l.id] = neu.id;
    melde(`Lehrkräfte (${i + 1}/${lehrer.length})…`);
  }
  const klassenIdKarte = {};
  melde(`Klassen (0/${klassen.length})…`);
  for (const [i, k] of klassen.entries()) {
    const neu = await klasseErstellen({
      ...k,
      lehrer1: k.lehrer1 ? lehrerIdKarte[k.lehrer1] : null,
      lehrer2: k.lehrer2 ? lehrerIdKarte[k.lehrer2] : null,
    });
    klassenIdKarte[k.id] = neu.id;
    melde(`Klassen (${i + 1}/${klassen.length})…`);
  }
  const lerngruppenIdKarte = {};
  melde(`Lerngruppen (0/${lerngruppen.length})…`);
  for (const [i, g] of lerngruppen.entries()) {
    const neu = await lerngruppeErstellen({
      ...g,
      fachId: fachIdKarte[g.fachId],
      lehrerId: g.lehrerId ? lehrerIdKarte[g.lehrerId] : null,
      klassenIds: (g.klassenIds || []).map((kid) => klassenIdKarte[kid]).filter(Boolean),
    });
    lerngruppenIdKarte[g.id] = neu.id;
    melde(`Lerngruppen (${i + 1}/${lerngruppen.length})…`);
  }
  const methodenIdKarte = {};
  melde(`Methoden (0/${methoden.length})…`);
  for (const [i, m] of methoden.entries()) {
    const neu = await methodeErstellen({ ...m, faecherIds: (m.faecherIds || []).map((fid) => fachIdKarte[fid]).filter(Boolean) });
    methodenIdKarte[m.id] = neu.id;
    melde(`Methoden (${i + 1}/${methoden.length})…`);
  }
  melde(`Planungen (0/${planungen.length})…`);
  let uebersprungen = 0;
  for (const [i, p] of planungen.entries()) {
    if (!methodenIdKarte[p.methodeId] || !lerngruppenIdKarte[p.lerngruppeId]) {
      uebersprungen++;
      continue;
    }
    await planungErstellen({
      ...p,
      methodeId: methodenIdKarte[p.methodeId],
      lerngruppeId: lerngruppenIdKarte[p.lerngruppeId],
      klasseId: p.klasseId ? klassenIdKarte[p.klasseId] : null,
    });
    melde(`Planungen (${i + 1}/${planungen.length})…`);
  }
  return { uebersprungen };
}
