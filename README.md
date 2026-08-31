# Methodencurriculum-App (Prototyp)

Web-App zur Planung und Dokumentation des schulischen Methodencurriculums – Zuordnung
von Methoden zu Fach/Klasse/Quartal, Durchführungsstatus, sowie Verwaltung von
Fächern, Lehrkräften, Klassen, Lerngruppen und einem Untis-CSV-Import.

Dieses Projekt ist ein **Prototyp** zum Ausprobieren, kein fertiges Produkt.

## Lokal starten

Voraussetzung: [Node.js](https://nodejs.org) (LTS-Version).

```bash
npm install
npm run dev
```

Die Konsole zeigt danach eine lokale Adresse (z. B. `http://localhost:5173`) – im
Browser öffnen.

## Auf GitHub veröffentlichen, damit Kolleg:innen es ausprobieren können

Damit die App unter einer eigenen Adresse erreichbar ist, ohne dass jemand Node.js
installieren muss:

1. Auf [github.com](https://github.com) ein neues, leeres Repository anlegen (z. B.
   `methodencurriculum-app`). **Kein** README/`.gitignore` beim Anlegen mit erzeugen
   lassen, das ist hier schon dabei.
2. Im Projektordner, einmalig:
   ```bash
   git init
   git add .
   git commit -m "Erster Stand"
   git branch -M main
   git remote add origin https://github.com/<dein-nutzername>/methodencurriculum-app.git
   git push -u origin main
   ```
3. Im GitHub-Repository: **Settings → Pages → Source** auf **"GitHub Actions"** stellen.
4. Nach dem Push läuft automatisch eine Aktion (sichtbar unter dem Reiter **Actions**),
   die die App baut und veröffentlicht. Nach ein bis zwei Minuten ist sie erreichbar
   unter:
   ```
   https://<dein-nutzername>.github.io/methodencurriculum-app/
   ```

Diesen Link können Kolleg:innen direkt öffnen – auch auf dem iPad, ganz ohne
Installation.

Spätere Änderungen: einfach `git add . && git commit -m "..." && git push` – die Seite
aktualisiert sich automatisch.

## Nutzung auf dem iPad

- Der Link lässt sich in Safari über **Teilen → Zum Home-Bildschirm** wie eine App
  ablegen (eigenes Icon, öffnet ohne Safari-Rahmen).
- Auf schmalen Ansichten (z. B. iPad im Hochformat) klappt die Seitenleiste ein; sie
  lässt sich über das Symbol ☰ oben links öffnen.
- Folgendes ist technisch grundsätzlich touch-tauglich, aber noch nicht auf einem
  echten Gerät geprüft – bitte gezielt testen:
  - Drag & Drop beim Zuordnen von Methoden in der Zeitleiste
  - Die Bild-Größenänderung im Methodenbeschreibung-Editor (Ecke ziehen)
  - Das Herunterladen der CSV-Exporte (iOS zeigt dafür meist einen
    "In Dateien sichern"-Dialog statt eines klassischen Downloads)

## Bekannte Einschränkungen

- **Daten sind nicht dauerhaft gespeichert.** Alles läuft im Arbeitsspeicher des
  Browsers – ein Neuladen der Seite setzt auf den Ausgangszustand zurück (bzw. auf
  den unten beschriebenen eingebetteten Testdatensatz). Das gilt auch für
  hochgeladene Materialien – für sehr große Dateien nicht geeignet.
- **Word-Import/-Export für Methodenbeschreibungen** – im Bearbeiten-Fenster einer
  Methodenbeschreibung: "Word importieren" liest ein .docx (inkl. Bilder) ein,
  "Als Word exportieren" erzeugt aus dem aktuellen Inhalt eine echte .docx-Datei.
  Der Editor selbst basiert auf [Tiptap](https://tiptap.dev) (kostenloser,
  quelloffener Kern) – Bilder lassen sich per Ecke ziehen in der Größe ändern und
  im Text verschieben.
- **Bearbeitung bestehender Methoden**: Name, Jahrgangsstufen, empfohlene Fächer und
  Halbjahr lassen sich in der Methoden-Verwaltung direkt ändern; Materialien
  (Arbeitsblätter etc.) lassen sich hochladen und wieder entfernen.
- **Eingebetteter Testdatensatz:** `src/App.jsx` enthält aktuell testweise einen
  kompletten Untis-CSV-Export als Text im Code (`STANDARD_UNTIS_EXPORT`), der beim
  Start automatisch geladen wird. Das ist bewusst nur eine Übergangslösung für die
  Entwicklung – vor einer Weitergabe an eine größere Gruppe oder einem produktiven
  Einsatz sollte das entfernt werden (Suche im Code nach `STANDARD_UNTIS_EXPORT`),
  damit nicht dauerhaft eine echte Schuldatei im Quellcode mitgeführt wird.
