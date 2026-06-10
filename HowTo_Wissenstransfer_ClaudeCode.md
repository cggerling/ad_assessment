# How-To: Aktuellen Wissensstand des AD-Assessment-Scripts in Claude Code übertragen

Dieses How-To beschreibt, wie du das bereinigte Skript `Analyse_V4_6.ps1` und den dokumentierten
Status (`PROJECT_CONTEXT.md`) so in Claude Code überführst, dass die **Weiterentwicklung dort
iterativ** stattfinden kann — **ohne** dass das ausgelieferte/versionierte Projekt einen Hinweis
auf das verwendete KI-Werkzeug enthält.

> **Wichtig zur Trennung:** Die Vorgabe „kein Hinweis auf das KI-Werkzeug" gilt für das
> **Skript-Artefakt und die darin/daran gebundenen Befehle**. Dieses How-To ist ein separates
> Anleitungsdokument für dich und benennt das Werkzeug daher zwangsläufig. Der unten beschriebene
> Aufbau hält das **versionierte Repository selbst werkzeug-neutral**.

---

## 0. Voraussetzungen

- Ein lokaler Projektordner / Git-Repository für das Skript.
- Claude Code lokal installiert.
  *(Hinweis: Die genaue, aktuell gültige Installationsmethode bitte der offiziellen Doku entnehmen
  — siehe „Quellen / Verweise". Ich habe den exakten Installationsbefehl in dieser Sitzung nicht
  erneut verifiziert und gebe ihn daher nicht als Faktum aus.)*

---

## 1. Projektordner aufsetzen

Lege einen Ordner an und kopiere die beiden Dateien hinein:

```
ad-assessment/
├── Analyse_V4_6.ps1        # bereinigtes Skript (werkzeug-neutral)
└── PROJECT_CONTEXT.md      # dokumentierter Status / Projektwissen (werkzeug-neutral)
```

Optional als Git-Repo initialisieren:

```bash
cd ad-assessment
git init
git add Analyse_V4_6.ps1 PROJECT_CONTEXT.md
git commit -m "AD-Assessment Script: bereinigt + Projektkontext dokumentiert"
```

Beide Dateien sind bewusst neutral benannt und enthalten **keinen** Werkzeug-Hinweis — sie können
unbedenklich committet/ausgeliefert werden.

---

## 2. Wissenstransfer — empfohlener, repo-neutraler Weg (Variante A)

Claude Code lädt zu Sitzungsbeginn automatisch eine Kontextdatei. Die Standard-/Team-Variante
hieße `CLAUDE.md` und würde mit ins Repo wandern — also einen werkzeug-bezogenen Dateinamen
ausliefern. Um das zu vermeiden, nutzt du eine **lokale, von Git ausgeschlossene** Kontextdatei,
die nur auf dein neutrales `PROJECT_CONTEXT.md` verweist:

1. Lege im Projektordner eine **lokale** Kontextdatei an (wird nur auf deiner Maschine geladen,
   nicht versioniert). Inhalt: **eine Zeile**, die das Projektwissen importiert:

   ```text
   @PROJECT_CONTEXT.md
   ```

   Dateiname: `CLAUDE.local.md` (lokale, persönliche Kontextdatei).

2. Schließe diese Datei in `.gitignore` aus, damit nichts Werkzeug-Benanntes ausgeliefert wird:

   ```gitignore
   # Lokale Werkzeug-Kontextdatei nicht versionieren
   CLAUDE.local.md
   ```

Ergebnis: Das **gesamte Projektwissen** steht in `PROJECT_CONTEXT.md` (neutral, versioniert) und
wird zu jeder Sitzung über die git-ignorierte Importdatei automatisch geladen. Das versionierte
Repository bleibt frei von Werkzeug-Hinweisen.

*(Belegt durch die offizielle Doku: Projekt-Kontextdateien werden zu Sitzungsbeginn vollständig
geladen, können per `@pfad`-Syntax weitere Dateien importieren, und die lokale Variante ist für
`.gitignore` vorgesehen.)*

---

## 3. Schneller Alternativweg (Variante B)

Wenn der werkzeug-neutrale Repo-Aufbau für eine reine Arbeitskopie (die nicht ausgeliefert wird)
nicht nötig ist:

- Im Projektordner eine Sitzung starten und `/init` ausführen. Damit analysiert das Werkzeug die
  Codebasis und erzeugt automatisch eine Kontextdatei mit erkannten Konventionen/Befehlen. Du kannst
  anschließend deinen Text aus `PROJECT_CONTEXT.md` einarbeiten.

> Für eine **auszuliefernde** Codebasis bleibt **Variante A** die saubere Wahl, weil `/init` eine
> werkzeug-benannte Datei direkt im Repo anlegt.

---

## 4. Laufender Wissenszuwachs (während der Weiterentwicklung)

- **Manuelle Aktualisierung:** Erkenntnisse, die dauerhaft gelten (Build-/Testbefehle,
  Konventionen, Architekturentscheidungen), in `PROJECT_CONTEXT.md` ergänzen. Da `CLAUDE.local.md`
  nur `@PROJECT_CONTEXT.md` importiert, wirkt jede Ergänzung automatisch in der nächsten Sitzung.
- **Automatische Notizen (Auto-Memory):** Das Werkzeug kann selbst Notizen über das Projekt
  ablegen (entdeckte Befehle, Muster, Vorlieben). Diese liegen außerhalb des Repos in deinem
  Benutzerprofil und werden nicht ausgeliefert. Mit `/memory` siehst und bearbeitest du, was
  gespeichert wurde.
- **Ad-hoc merken:** Eine Anweisung wie „merke dir: CRLF-Zeilenenden und UTF-8 beibehalten" wird
  in die automatische Notiz übernommen; „nimm das in PROJECT_CONTEXT.md auf" schreibt es in deine
  versionierte Kontextdatei.

---

## 5. Sauber halten — Checkliste vor Auslieferung

- [ ] `Analyse_V4_6.ps1` enthält keinen Werkzeug-/Personen-/Firmenhinweis. *(Aktuell erfüllt.)*
- [ ] `PROJECT_CONTEXT.md` ist werkzeug-neutral. *(Aktuell erfüllt.)*
- [ ] `CLAUDE.local.md` steht in `.gitignore` und ist **nicht** committet.
- [ ] Auto-Memory-Ablage liegt außerhalb des Repos (Benutzerprofil) — nicht mit ausliefern.
- [ ] Vor Übergabe: `git log`/`git diff` und einen Suchlauf nach Werkzeug-/Personen-Begriffen
      über das gesamte Repo laufen lassen.

---

## 6. Empfohlener Einstieg in die Weiterentwicklung

Sinnvolle erste Schritte in der Sitzung (iterativ, je ein Thema):

1. „Lies `PROJECT_CONTEXT.md` und fasse Aufbau und offene Modernisierungs-Themen zusammen."
2. Ein Thema aus Abschnitt 6 der `PROJECT_CONTEXT.md` auswählen (z. B. Fehlerhandling pro
   Prüfblock) und **klein** umsetzen — read-only-Charakter und Report-Layout dabei wahren.
3. Vor größeren Refactorings ein Sicherheitsnetz schaffen (z. B. Tests für die
   Formatierungs-/Hilfsfunktionen), dann schrittweise modernisieren.
4. Nach jeder gesicherten Erkenntnis `PROJECT_CONTEXT.md` aktualisieren.

---

## Quellen / Verweise

- Offizielle Doku „How Claude remembers your project" (Projekt-Kontextdateien, `@pfad`-Import,
  lokale/`.gitignore`-Variante, `/init`, Auto-Memory, `/memory`).
- Offizielle Claude-Code-Produktdokumentation (Übersicht/Installation) — für die aktuell gültige
  Installationsmethode.

*(Bewusst keine Inline-Links im Fließtext; Verweise gebündelt in diesem Abschnitt.)*
