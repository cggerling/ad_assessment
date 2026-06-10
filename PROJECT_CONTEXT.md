# PROJECT_CONTEXT — AD-Assessment Script

> Zentrale Projekt-/Kontextdatei für die Weiterentwicklung des Skripts `Analyse_V4_6.ps1`.
> Inhalt: Was das Skript tut, wie es aufgebaut ist, welche Konventionen gelten und welche
> Modernisierungsthemen offen sind. Diese Datei ist bewusst werkzeug-neutral gehalten.

---

## 1. Zweck & Charakter

- Das Skript erzeugt einen **fest formatierten Text-Report** über eine On-Premises
  **Active-Directory-Umgebung** (Assessment / Bestandsaufnahme).
- Es arbeitet **ausschließlich lesend** — es führen keine Funktionen schreibende
  AD-Operationen aus. *(Verifiziert: keine `Set-AD*`/`New-AD*`/`Remove-AD*`-Aufrufe gefunden;
  einzige schreibende Aktion ist das Anlegen von Ausgabeverzeichnis/-datei.)*
- Aktuelle Version laut Kopf: **4.6**.

## 2. Tech-Stack & Abhängigkeiten *(verifiziert anhand der genutzten Cmdlets)*

| Abhängigkeit | Wofür | Belegt durch |
|---|---|---|
| PowerShell-Modul **ActiveDirectory** | Domain/Forest/Objekte | `Get-ADDomain`, `Get-ADForest`, `Get-ADComputer`, `Get-ADUser`, `Get-ADObject`, `Get-ADGroup(Member)`, `Get-ADTrust`, `Get-ADServiceAccount`, `Get-AD*PasswordPolicy`, `Get-ADRootDSE`, `Get-ADOptionalFeature`, … |
| PowerShell-Modul **GroupPolicy** | GPO-Auswertung | `Get-GPO`, `Get-GPOReport` |
| PowerShell-Modul **DnsServer** | DNS-Prüfungen | `Get-DnsServerZone`, `Get-DnsServerZoneAging`, `Get-DnsServerScavenging`, `Get-DnsServerForwarder`, `Get-DnsServerResourceRecord` |
| **WinRM / `Invoke-Command`** | Remote-Prüfungen auf Domain Controllern (Dienste, Rollen, Features, LDAPS, NTLM, SMB1, BitLocker, ExecutionPolicy) | zahlreiche `Invoke-Command -ComputerName … -ScriptBlock { … }` |

> **Ausführungskontext (Einschätzung, nicht hart im Skript erzwungen):** Ausführung auf einem
> Domain Controller oder einem Management-Host mit installierten RSAT-Modulen, im Kontext eines
> Kontos mit Domänen-Leserechten. Für einzelne DC-Detailprüfungen sind erhöhte Rechte
> (Domain-Admin-äquivalent) praktisch erforderlich. *Das Skript prüft beim Start:
> `#Requires -Version 5.1` plus Modul-Vorabprüfung (ActiveDirectory = Pflicht/Abbruch;
> GroupPolicy/DnsServer = Warnung, betroffener Bereich wird deaktiviert).*

## 3. Architektur / Aufbau

Das Skript hat **65 Funktionen** und folgt grob drei Schichten:

1. **Kopf / Konfiguration (oben im Skript):**
   - Metadaten: `$version`, `$company`, `$madeby`, `$maintitel`, `$B_Datei`.
   - **Schalter-Variablen** je Prüfbereich (`$domoco`, `$schema`, `$censto`, `$domdcs`,
     `$loggin`, `$adtchk`, `$dnschk`, `$SysRep`, `$admusr`, `$usrchk`, `$syschk`, `$manacc`,
     `$dDPchk`, `$allgru`, `$allgpo`, `$OrgUni`, `$caschk`, `$DomCon`, …). Werte: `0` = aus,
     `1` = an, teils `2` = erweitert.
   - **Ausgabe-/Pfad-Variablen:** `$verz` (Ausgabeverzeichnis, jetzt `c:\AD-Assessment`),
     `$sysverz = $verz\$system`, Dateiname `$B_Datei_<Datum>.txt`, Ausgabebreite `$sb`.
   - Ausgabesteuerung: `$A_Dat` (Datei) / `$A_Con` (Konsole).

2. **Formatierungs-/Layout-Funktionen** (erzeugen den ASCII-Report):
   `Header`, `Bottom`, `Vollzeile`, `Leerzeile`, `Trennzeile`, `tablinie`, `Bereich`,
   `Bereichstitel`, `Subtitel`, `2werte`, `new_2werte`, `neu_tab_max6w_fb`, `neu_text`.
   *(Hinweis: `2werte` beginnt mit einer Ziffer — in PowerShell zulässig, aber unüblich.)*

3. **Prüf-Funktionen je Themenbereich**, u. a.:
   `dom_allgemein`, `centralstore`, `sec_templates`, `controller`, `Auditcheck`, `trusts`,
   `aging`, `dfsr`, `Get-AllDomainControllers`, `controller_check`, `Admins`,
   `lokale_AdmGru`, `dom_AdmGri`, `AdmCount`, `builtin_usr`, `User_chk`, `inaktive_User`,
   `gesperrte_User`, `ou_users`, `sys_konten`, `clt_chk`, `srv_chk`, `oth_chk`, `KDSR`,
   `MSA`, `gMSA`, `ad_gruppen`, `GPO_all`, `ddomainpol`, `fGPO`, `spezial_user`, `OUS`,
   `dacls`, `ca_root`, `ca_sub`, `ca_templates`, `dcdienste`, `dcprog`, `dchot`, `dcroles`,
   `dcfeature`, `dc_ldaps`, `NTLM`, `dc_SMB1`, `Power`, `OF_Bitlocker`, `AD_Controller`.

4. **Hauptablauf (Ende des Skripts):** eine Folge von
   `if ($schalter -ge 1) { Pruefbereich "<Titel>" { <Funktionen> } }`-Blöcken, abgeschlossen durch
   `bottom`. `Pruefbereich` kapselt jeden Bereich in `try/catch`: Ein Fehler wird im Report
   vermerkt und rot auf der Konsole gemeldet, danach läuft der nächste Bereich weiter.

## 4. Ausgabe

- **Ziel:** `c:\AD-Assessment\<COMPUTERNAME>\<B_Datei>_<dd-MM-yyyy HH.mm.ss>.txt`
- **Format:** feste Spaltenbreite (`$sb`, 70–90), ASCII-Rahmen, Konsolen-Farbausgabe via `Write-Host`.
- Datei wird via `Out-File … -Encoding ascii` angelegt; die Report-Zeilen werden in einem
  `StringBuilder`-Puffer gesammelt (`Ausgabe`) und je Bereich in einem Rutsch geschrieben
  (`Puffer_leeren`, aufgerufen im `finally` von `Pruefbereich` sowie nach Header/Bottom).

## 5. Konventionen, die erhalten bleiben sollen

- **Sprache:** Kommentare und Report-Texte sind **deutsch** — beibehalten.
- **Dateikodierung:** UTF-8, **CRLF**-Zeilenenden — beibehalten.
- **Read-only-Charakter:** Es dürfen **keine schreibenden AD-Operationen** ergänzt werden.
- **Schalter-Logik:** Bestehende Schalter-Variablen-Semantik (0/1/2) nicht brechen.
- **Report-Layout:** Spaltenausrichtung/Rahmen ist Teil des Outputs; Änderungen daran sind
  bewusst und getestet vorzunehmen.

## 6. Offene Modernisierungs-Themen

> **Klar getrennt:** „Beobachtet" = im Code verifiziert. „Empfehlung" = fachliche Einschätzung,
> noch nicht entschieden/umgesetzt — vor Umsetzung bewerten.

**Beobachtet (verifizierbar im Code):**
- ~~Keine `#Requires`-Direktiven~~ → *erledigt (PR „Fundament"): `#Requires -Version 5.1` + Modul-Vorabprüfung.*
- ~~Ausgabe via zeilenweisem `Add-Content` (viele Einzel-I/O-Operationen)~~ → *erledigt
  (PR „Performance"): gepufferte Ausgabe, ein Schreibvorgang je Bereich (~1.000× schneller).*
- ~~Keine zentrale Fehlerbehandlung (`try/catch`) erkennbar~~ → *erledigt (PR „Fehlerbehandlung"):
  `Pruefbereich`-Wrapper pro Bereich; `$DCs`-Ermittlung beim Start abgesichert.*
- Remoting durchgängig über `Invoke-Command` (WinRM-Abhängigkeit, keine Session-Wiederverwendung an allen Stellen).
- Funktionsname `2werte` beginnt mit Ziffer (legal, aber unüblich/fehleranfällig beim Aufruf).
- Auskommentierte Bereiche vorhanden (z. B. `dacls`/`$aclchk`, `ca_templates`).

**Empfehlung (Einschätzung — zu bewerten/iterieren):**
- ~~`#Requires -Version` und Modul-Vorabprüfung ergänzen~~ → *erledigt (PR „Fundament").*
- ~~Strukturiertes Fehlerhandling pro Prüfblock~~ → *erledigt (PR „Fehlerbehandlung").*
- Optionalen **strukturierten Export** (CSV/JSON/HTML) zusätzlich zum Text-Report prüfen —
  erleichtert spätere Auswertung im Kontext der geplanten AD-Ablösung. HTML-Report im Stil
  von water.css gewünscht.
- ~~Performance: I/O bündeln (StringBuilder)~~ → *erledigt (PR „Performance").*
- Parametrisierung (z. B. Ausgabepfad, Zielbereiche) statt fester Variablen im Kopf erwägen.
- ~~Pester-Tests für die Formatierungs-/Hilfsfunktionen~~ → *erledigt (PR „Fundament"):
  `Tests/Analyse_V4_6.Format.Tests.ps1` (Pester 5, 25 Tests). Ausführen mit
  `Invoke-Pester -Path .\Tests`; läuft unter PowerShell 5.1 und 7.*

## 7. Aktueller Stand (Changelog)

**PR „Performance" (Juni 2026):**
- Datei-I/O gebündelt: neue Funktionen `Ausgabe` (sammelt Report-Zeilen in einem
  `StringBuilder`) und `Puffer_leeren` (schreibt den Puffer in einem Rutsch in die Datei).
  Alle 21 `Write-Output … | Add-Content`-Stellen der Formatierungsfunktionen umgestellt.
- Geschrieben wird je Bereich (`finally` in `Pruefbereich`), nach dem Header und am Skriptende —
  bei einem Abbruch geht maximal der aktuelle Bereich verloren.
- Messung: 2.000 Zeilen alt ~88 s, neu ~0,08 s (lokal; Faktor ~1.100).
- Testsuite auf 34 Tests erweitert (Puffer-Semantik, Flush im Fehlerfall, statischer Check:
  genau eine `Add-Content`-Stelle im Skript).

**PR „Fehlerbehandlung" (Juni 2026):**
- Neue Funktion `Pruefbereich ($titel, $aktion)`: kapselt jeden Report-Bereich in `try/catch`.
  Fehler → Vermerk im Report (`FEHLER - Bereich nur unvollstaendig geprueft` + Meldung +
  Skriptzeile) und rote Konsolenmeldung; der Lauf wird mit dem nächsten Bereich fortgesetzt.
- Alle 22 Blöcke des Hauptablaufs auf `Pruefbereich` umgestellt (Layout unverändert,
  per Test abgesichert).
- `$DCs`-Ermittlung beim Skriptstart abgesichert: schlägt sie fehl → klare Meldung + Abbruch.
- Header-Bugfixes: Dateiausgabe des Headers funktioniert jetzt auch bei `$A_Con = 0`
  (vorher leere Zeilen); `$firma`-Werte ≠ 16 Zeichen sprengen nicht mehr den Rahmen
  (zu lange Werte werden mit `~` gekürzt).
- Testsuite auf 30 Tests erweitert (Pruefbereich, Header-Fixes, statischer Check:
  kein ungeschützter `Bereich`-Aufruf im Hauptablauf).

**PR „Fundament" (Juni 2026):**
- `#Requires -Version 5.1` am Skriptanfang.
- Modul-Vorabprüfung vor Anlage der Ausgabedatei: `ActiveDirectory` fehlt → klare Meldung +
  Abbruch (`exit 1`); `GroupPolicy`/`DnsServer` fehlen → Warnung + betroffener Schalter
  (`$allgpo` bzw. `$dnschk`) wird auf 0 gesetzt, der Lauf geht weiter.
- Pester-5-Tests für die 13 Formatierungs-/Layout-Funktionen (`Tests/Analyse_V4_6.Format.Tests.ps1`).
  Die Tests extrahieren die Funktionen per AST aus der Skriptdatei (das Skript wird dabei nicht
  ausgeführt) und prüfen das Datei-Layout zeichengenau; zusätzlich statische Prüfungen
  (parsebar, `#Requires` vorhanden, keine schreibenden AD-Cmdlets).

**Vorarbeit (Bereinigungs-Durchlauf):**
- Firmen-/Personenbezüge entfernt: `$company` → `"AD-Assessment"`, `$madeby` → `"AD-Assessment Tool"`
  (vormals Firmen- bzw. Personenkürzel).
- Ausgabeverzeichnis `$verz` von `c:\p-sec` auf **`c:\AD-Assessment`** umgestellt.
- Werkzeug-neutraler Comment-Based-Help-Block (`.SYNOPSIS/.DESCRIPTION/.OUTPUTS/.NOTES`) im Kopf ergänzt.
- Keine Logik-/Funktionsänderungen — Bereinigung war rein chirurgisch.

---

## Quellen / Verweise

- PowerShell Comment-Based Help (`about_Comment_Based_Help`) — Konzept des `.SYNOPSIS/.DESCRIPTION`-Blocks.
- Microsoft-Module: `ActiveDirectory`, `GroupPolicy`, `DnsServer` (RSAT) — Cmdlet-Referenzen.

*(Bewusst keine Inline-Links im Fließtext; Verweise gebündelt in diesem Abschnitt.)*
