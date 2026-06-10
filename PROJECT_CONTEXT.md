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
> (Domain-Admin-äquivalent) praktisch erforderlich. *Es gibt aktuell keine `#Requires`-Zeile —
> Abhängigkeiten werden zur Laufzeit nicht hart geprüft.*

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

4. **Hauptablauf (Ende des Skripts):** eine Folge von `if ($schalter -ge 1) { Bereich …; <Funktion> }`-
   Blöcken, die die Prüf-Funktionen abhängig von den Schaltern aufrufen, abgeschlossen durch `bottom`.

## 4. Ausgabe

- **Ziel:** `c:\AD-Assessment\<COMPUTERNAME>\<B_Datei>_<dd-MM-yyyy HH.mm.ss>.txt`
- **Format:** feste Spaltenbreite (`$sb`, 70–90), ASCII-Rahmen, Konsolen-Farbausgabe via `Write-Host`.
- Datei wird via `Out-File … -Encoding ascii` angelegt und zeilenweise mit `Add-Content` befüllt.

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
- Keine `#Requires`-Direktiven → fehlende Module/Rechte fallen erst zur Laufzeit auf.
- Ausgabe via `Out-File -Encoding ascii` + zeilenweisem `Add-Content` (viele Einzel-I/O-Operationen).
- Keine zentrale Fehlerbehandlung (`try/catch`) erkennbar; Remote-Fehler je DC können den Lauf stören.
- Remoting durchgängig über `Invoke-Command` (WinRM-Abhängigkeit, keine Session-Wiederverwendung an allen Stellen).
- Funktionsname `2werte` beginnt mit Ziffer (legal, aber unüblich/fehleranfällig beim Aufruf).
- Auskommentierte Bereiche vorhanden (z. B. `dacls`/`$aclchk`, `ca_templates`).

**Empfehlung (Einschätzung — zu bewerten/iterieren):**
- `#Requires -Version` und Modul-Vorabprüfung ergänzen (Abhängigkeiten früh & klar melden).
- Strukturiertes Fehlerhandling pro Prüfblock (`try/catch`, Fortsetzen statt Abbruch).
- Optionalen **strukturierten Export** (CSV/JSON/HTML) zusätzlich zum Text-Report prüfen —
  erleichtert spätere Auswertung im Kontext der geplanten AD-Ablösung.
- Performance: I/O bündeln (StringBuilder statt vieler `Add-Content`-Aufrufe) prüfen.
- Parametrisierung (z. B. Ausgabepfad, Zielbereiche) statt fester Variablen im Kopf erwägen.
- Pester-Tests für die Formatierungs-/Hilfsfunktionen als Sicherheitsnetz vor größeren Refactorings.

## 7. Aktueller Stand (Changelog dieses Durchlaufs)

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
