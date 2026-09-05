import { createClient } from "@supabase/supabase-js";

// Projekt-Adresse und der öffentliche "Publishable Key" - beide sind bewusst öffentlich
// (der Key läuft im Browser mit, das ist bei Supabase so vorgesehen). Die eigentliche
// Absicherung passiert über die Row-Level-Security-Regeln in der Datenbank, nicht durch
// Geheimhaltung dieses Schlüssels. Den "Secret Key" (Gegenstück zum alten service_role-
// Schlüssel) NIE hier eintragen - der würde alle Zugriffsregeln umgehen.
const SUPABASE_URL = "https://aoeopmgsauarboqddrlg.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_BsT9Jmcaz8IoI0dGTD8YjA_Si47upVu";

export const supabase = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);
