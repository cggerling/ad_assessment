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
- Aktuelle Version laut Kopf: **5.0** (Ausbau zum Security-Assessment, laufend; Dateiname
  bleibt vorerst `Analyse_V4_6.ps1` aus Kontinuitätsgründen).

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
   - **`param()`-Block** (optional beim Aufruf): `-Verzeichnis`, `-Breite` (70–90),
     `-KeineKonsole`, `-KeineDatei`, `-Bereiche` (Hashtable mit Schalter-Overrides,
     z. B. `@{ dnschk = 0 }`). Ohne Parameter gelten die Standardwerte aus dem Kopf.
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
- **Zusatz-Ausgaben** (abschaltbar via `$A_Htm`/`$A_Jsn` bzw. `-KeinHTML`/`-KeinJSON`):
  Die Formatierungsfunktionen erfassen jede Ausgabe zusätzlich als strukturiertes Ereignis
  (`Merken` → `$R_Daten`). Daraus entstehen am Skriptende ein **HTML-Report** (`.html`,
  water.css-inspiriertes eingebettetes CSS, offline-fähig, hell/dunkel automatisch,
  Befund-Farben als CSS-Klassen ok/warn/err) und ein **JSON-Export** (`.json`) mit
  denselben Basisnamen wie der Text-Report.

## 5. Konventionen, die erhalten bleiben sollen

- **Sprache:** Kommentare und Report-Texte sind **deutsch** — beibehalten.
- **Dateikodierung:** UTF-8 **mit BOM**, **CRLF**-Zeilenenden — beibehalten. Der BOM ist
  wichtig, damit Windows PowerShell 5.1 die Umlaute in den Katalogtexten korrekt liest
  (BOM-lose UTF-8-Skripte interpretiert 5.1 als ANSI). Bei jeder Skript-Bearbeitung den
  BOM erhalten bzw. wiederherstellen.
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
- ~~Optionalen strukturierten Export (JSON/HTML) zusätzlich zum Text-Report~~ → *erledigt
  (PR „Export"): HTML-Report im water.css-Stil + JSON-Export. CSV bewusst ausgelassen —
  bei Bedarf aus dem JSON ableitbar.*
- ~~Performance: I/O bündeln (StringBuilder)~~ → *erledigt (PR „Performance").*
- ~~Parametrisierung (Ausgabepfad, Zielbereiche)~~ → *erledigt (PR „Parametrisierung").*
- ~~Pester-Tests für die Formatierungs-/Hilfsfunktionen~~ → *erledigt (PR „Fundament"):
  `Tests/Analyse_V4_6.Format.Tests.ps1` (Pester 5, 25 Tests). Ausführen mit
  `Invoke-Pester -Path .\Tests`; läuft unter PowerShell 5.1 und 7.*

## 7. Aktueller Stand (Changelog)

**Neue Datei `AD-Analyse-V5.ps1` als Arbeitsbasis (Juni 2026):**
- Byte-genaue Kopie des aktuellen `Analyse_V4_6.ps1` (UTF-8 mit BOM) unter neuem Namen; nur die
  vier `.EXAMPLE`-Selbstaufrufe auf `.\AD-Analyse-V5.ps1` umgestellt. Soll künftig die Grundlage
  für weitere Anpassungen sein. `Analyse_V4_6.ps1` bleibt vorerst bestehen.
- **Offen/zu entscheiden:** Die Pester-Tests (`Tests/Analyse_V4_6.Format.Tests.ps1`) zeigen weiter
  auf die alte Datei. Für Weiterentwicklung an `AD-Analyse-V5.ps1` sollten die Tests darauf
  umgezogen (oder dupliziert) und die alte Datei ggf. später entfernt werden.

**Inhalt „Texte etwas weniger technisch" (Juni 2026):**
- Auf Wunsch (Management/Audit-Lesbarkeit) die Begründungstexte (Zweck/Hintergrund) aller ~45
  Katalogeinträge entschärft: mit der Wirkung beginnen, tiefe Identifier entfernt (GUIDs, Hex-UAC-
  Flags, OIDs, Hashcat-Modi, API-/Interna-Namen wie MS-DRSR/SDProp/kdssvc, Event-IDs, EDITF/
  msPKI-Flags, CVE-Nummern in der Prosa), kürzer gefasst. Erkennbare Anker (DCSync, AdminSDHolder,
  Kerberoasting, Golden Ticket, ESC1-8, PrinterBug) bleiben. **Titel/Schwere/Beispiele/Quellen
  unverändert.**
- Tonalität vorab an dcsync/adminsdholder kalibriert und freigegeben; per verifizierter Voll-
  Fragment-Ersetzung ausgerollt (Restjargon-Scan: 0 in Zweck/Hintergrund). 93 Tests grün PS 7 + 5.1.

**Bugfix „Delta-Bereich doppelter Doku-Block" (Juni 2026):**
- Im Delta-Bereich teilten sich `Pruefbereich` und die einzelne `Unterpruefung` die CheckId
  `'delta'`, wodurch Begründung (Doku), Exec-Summary-Eintrag und HTML-Anker `chk-delta`
  **doppelt** erschienen (am DC-Report 18.06. aufgefallen). Fix: die Unterprüfung nutzt jetzt
  CheckId `$null` (Doku kommt allein vom Bereich).
- Neuer struktureller Regressionstest: **keine CheckId** darf gleichzeitig von einem
  `Pruefbereich` und einer `Unterpruefung` verwendet werden (per AST geprüft). Suite **92 → 93**,
  grün PS 7 + 5.1.
- DC-Lauf 18.06. bestätigte Paket F (Delta 70/70, 0 neu/behoben) und den Umlaut-Feinschliff
  (Text-Report UTF-8+BOM, kein Mojibake, Labels korrekt). `dcdienste`-Fix dort nicht ausgelöst
  (Detailmodus `$DomCon=1`); zum Testen `-Bereiche @{ DomCon = 2 }`.

**Umlaut-Feinschliff Teil 2 „Alt-Labels de-transliteriert" (Juni 2026):**
- Die verbliebenen ~42 ASCII-transliterierten **Anzeige-Labels und Abschnittstitel** (z. B.
  „DC-Härtung (vertieft)", „Kerberos - Angriffsflächen", „GPP-Passwörter", „Veränderungen seit
  letztem Lauf", „Verfügbarkeit AD-Dienste", „über Port … möglich", „Uneingeschränkte
  Delegation", „läuft (PrinterBug/PetitPotam)", „für breite Gruppen", „gefährlich") auf echte
  Umlaute umgestellt. Damit nutzen Text-Report, HTML und JSON jetzt durchgängig korrekte Umlaute.
- Methode: per AST alle deutschen Transliterations-Strings extrahiert (CSS, Datumsformate,
  englische Begriffe, bereits korrekte Texte ausgeschlossen), dann verifizierte Voll-Fragment-
  Ersetzung (Umlaute aus Zeichencodes erzeugt, kein blindes ae/oe/ue/ss-Replace). **Code-
  Identifier unangetastet** (CheckIds wie `dc_haertung`, Funktionsnamen `Pruefbereich`/
  `Unterpruefung`/`Entschluessle-GPP`, Switches, AD-Filter, Registry-Pfade). Kontrolle danach:
  0 verbliebene transliterierte Labels, Parse OK.
- Tests unverändert grün (**92**, PS 7 + 5.1).
- Damit ist der Umlaut-Feinschliff abgeschlossen. Offen bleibt nur noch die optionale
  inhaltliche Tiefen-Justierung der Texte („technischer/weniger technisch"), am besten am echten
  Report bewertet.

**Umlaut-Feinschliff Teil 1 „Encoding-Fundament" (Juni 2026):**
- Text-Report jetzt **UTF-8 mit BOM** statt ASCII: Datei wird per
  `[IO.File]::WriteAllText($path,'',UTF8(BOM))` angelegt, der gepufferte Inhalt per
  `[IO.File]::AppendAllText($path,…,UTF8(ohne BOM))` angehängt (genau ein BOM am Dateianfang,
  verifiziert auch nach mehreren Flushes). Ersetzt `Out-File -Encoding ascii` + `Add-Content`.
- `neu_text` transliteriert **nicht mehr** (ü→ue etc. entfernt); Umlaute bleiben im Text-Report
  erhalten. HTML/JSON bekamen über `Merken` ohnehin schon das Original. Damit zeigen u.a. die
  Fehlermeldungen („Bereich nur unvollständig geprüft", „Teilprüfung übersprungen") jetzt echte
  Umlaute.
- Tests: `Get-ReportZeilen` liest UTF-8; betroffene Erwartungen auf Umlaute umgestellt; neuer
  statischer Guard (kein `-Encoding ascii`/keine `neu_text`-Transliteration mehr, ein
  `AppendAllText`). Suite **91 → 92**, grün PS 7 + 5.1. Round-trip separat unter PS 5.1 geprüft.
- Teil 2 (die ~60 ASCII-transliterierten Alt-Labels/Bereichstitel de-transliterieren) folgt als
  eigener PR.

**Bugfix „dcdienste: Enter-PSSession → Invoke-Command" (Juni 2026):**
- `dcdienste` (Dienst-Inventar je DC, DC-Detailmodus `DomCon=2`) nutzte `Enter-PSSession`/
  `Exit-PSSession` — im nicht-interaktiven Skript wirkungslos, die `Get-Service`/`Get-WmiObject`
  liefen lokal statt auf dem Ziel-DC. Jetzt sammelt die Funktion alle Werte read-only via
  `Invoke-Command -ComputerName $dcho` (Zählungen auf Live-Objekten, Status/Starttyp als String
  zurückgegeben → robust gegen Deserialisierung); Ausgabe bleibt lokal. Counts zudem
  `@(...).Count`-sicher (Einzeltreffer-Problem).
- **Read-only bestätigt:** vollständiges Audit ergab keine schreibenden AD-/GPO-/Registry-/
  Dienst-/ADSI-Operationen; einzige Schreibziele sind der Ergebnisordner + Reportdateien sowie
  (transient, selbstgelöscht) zwei `secedit /export`-Temp-Dateien je DC in dessen %TEMP%.
- Neuer statischer Regressionstest: Skript enthält kein `Enter-/Exit-PSSession` mehr;
  `dcdienste` nutzt `Invoke-Command -ComputerName $dcho`. Testsuite **90 → 91**, grün PS 7 + 5.1.

**v5.0 PR „Paket F – Delta-Modus / Zeitvergleich" (Juni 2026):**
- Neuer Parameter `-Vergleich <Pfad>` (+ Schalter `$deltchk`, Default 1) und Bereich
  „Veränderungen seit letztem Lauf (Delta)". Läuft **nur**, wenn `-Vergleich` gesetzt ist,
  ganz am Ende nach allen anderen Bereichen.
- Zwei reine Datenfunktionen (keine AD-Abfragen): `Extrahiere-Befunde` bildet aus einer
  Ereignisliste die Menge der rot/gelb markierten Befunde (eigene Delta-Artefakte werden
  herausgefiltert); `chk_delta` lädt einen früheren JSON-Export, vergleicht ihn mit dem
  aktuellen Lauf (`$R_Daten`) und listet **Neu hinzugekommen** (rot) und **Behoben** (grün)
  plus Zähler „unverändert".
- Robustheit: fehlende/ungültige Vergleichsdatei wird freundlich gemeldet; Hinweis, wenn der
  aktuelle Lauf ohne Export (kein HTML/JSON) läuft.
- Ein neuer Katalogeintrag `delta` (Schwere Info) mit Hintergrund + verifizierten Links
  (MS Best Practices for Securing AD, MS Security Baselines).
- Testsuite auf **90 Tests** (+9): Katalog/Schalter/Parameter plus **funktionale** Tests der
  Delta-Logik (rein lokal, ohne AD: Befund-Extraktion, Neu/Behoben-Erkennung gegen ein
  Temp-JSON, Artefakt-Filter, fehlende Datei). Grün unter PS 7 und PS 5.1.
- **Letztes geplantes Sicherheitspaket.** Danach folgt der große Umlaut-/Content-Feinschliff.

**v5.0 PR „Paket E – DC-Härtung vertieft" (Juni 2026):**
- Neuer Schalter `$dchaert` (Default 1) und Bereich „DC-Härtung (vertieft)".
- Vier read-only Checks je DC: **LDAP-Signing & Channel Binding** (Registry NTDS\Parameters:
  LDAPServerIntegrity, LdapEnforceChannelBinding), **SMB-Signing erforderlich** (LanManServer:
  RequireSecuritySignature), **Print Spooler** (Get-Service, PrinterBug/PetitPotam), **anonyme
  LDAP-Binds** (dSHeuristics 7. Zeichen = 2). Iteriert über `$DCs`, Remote-Registry via
  Invoke-Command/Get-Service.
- Fünf neue Katalogeinträge mit Hintergrund + verifizierten Links (MS LDAP-Signing, MS
  SMB-Signing, MITRE T1187 Forced Authentication, MS Security Baselines).
- Testsuite auf 81 Tests. **Registry-/Dienst-Abfragen nicht lokal testbar -> DC-Test.**

**v5.0 PR „Paket D – GPO/SYSVOL-Geheimnisse" (Juni 2026):**
- Neuer Schalter `$sysvchk` (Default 1) und Bereich „GPO & SYSVOL – Geheimnisse".
- Drei read-only Checks: **GPP-cpassword** (durchsucht SYSVOL-XMLs, entschlüsselt cpassword
  mit dem öffentlichen AES-Schlüssel MS14-025 = Kritisch), **SYSVOL-Skripte** (heuristische
  Mustersuche nach Klartext-Credentials in .bat/.cmd/.ps1/.vbs/.kix), **GPO-Bearbeitungsrechte**
  (GPO-ACLs mit Schreibrecht für breite Gruppen, reuse `Ist-NiedrigPriv`).
- Helfer `Entschluessle-GPP` (AES-256-CBC, Zero-IV, PKCS7, UTF-16); per Round-Trip-Test
  **lokal verifiziert** (Krypto ist ohne AD prüfbar).
- Vier neue Katalogeinträge (gpp_cpassword=Kritisch) mit Hintergrund + verifizierten Links
  (MITRE T1552.006/.001, T1484.001, MS14-025/CVE-2014-1812).
- Testsuite auf 78 Tests. **Datei-/SYSVOL-/ACL-Abfragen nicht lokal testbar -> DC-Test.**

**v5.0 PR „Paket C Nachschliff: konkretere ESC-Befunde" (Juni 2026, nach DC-Lauf):**
- ESC-Treffer jetzt klar beschriftet: „Vorlage: <Name>" plus die handlungsrelevante Info,
  **wer** die Vorlage anfordern darf (ESC1/ESC2/ESC3) bzw. welcher Prinzipal welches
  Schreibrecht hat (ESC4). Neuer Helfer `Get-NiedrigPrivEnroller` (ersetzt
  `Hat-NiedrigPrivEnroll`, liefert die Prinzipal-Namen statt nur bool).
- DC-Lauf (Lab it-pirate) fand korrekt: 1x ESC1 (UserAzureStrong), 11x ESC4,
  ESC8 (Web Enrollment installiert) — die Checks arbeiten also wie gewünscht.

**v5.0 PR „Paket C – AD CS / ESC1–8" (Juni 2026):**
- Neuer Schalter `$adcschk` (Default 1, in Whitelist) und Bereich „AD CS – Zertifikatsdienste (ESC)".
- Read-only über AD-Objekte unter `CN=Public Key Services` (Configuration-NC) + `certutil`:
  Bestand (CAs/Vorlagen), **ESC1** (ENROLLEE_SUPPLIES_SUBJECT + Auth-EKU + kein Approval +
  Enroll für breite Gruppen = Kritisch), **ESC2/ESC3** (Any-Purpose-/Enrollment-Agent-EKU),
  **ESC4** (Vorlagen-ACL mit Schreibrecht für breite Gruppen), **ESC6** (CA-Flag
  EDITF_ATTRIBUTESUBJECTALTNAME2 via certutil), **ESC8** (Web-Enrollment-Rolle via WinRM).
- Helfer `Get-ADCSObjekte`, `Ist-NiedrigPriv` (SID-basiert, sprachunabhängig),
  `Hat-NiedrigPrivEnroll` (Enrollment-Right-GUID 0e10c968-…).
- Sechs neue Katalogeinträge (esc1=Kritisch) mit Hintergrund + verifizierten Links
  (SpecterOps „Certified Pre-Owned", Microsoft AD CS overview).
- Testsuite auf 73 Tests. Generischer Unterpruefung-CheckId-Test toleriert jetzt `$null`
  (Inventory-Teilprüfung ohne Doku). **AD-/certutil-/WinRM-Abfragen nicht lokal testbar ->
  Funktionstest am DC.**

**v5.0 PR „Paket B Nachschliff + Umlaut-Quickwins" (Juni 2026, nach DC-Lauf):**
- AdminSDHolder-Check verfeinert: nur noch übernahme-relevante Rechte (GenericAll,
  GenericWrite, WriteDacl, WriteOwner) gelten als Befund; reines (oft attributgebundenes)
  WriteProperty/ReadProperty ist häufig ein legitimes Default-ACE (Cert Publishers, TS
  License Servers, Azure AD Connect) und wird nicht mehr fälschlich rot markiert.
- Umlaut-Quickwins in HTML-only/neu_text-Strings (Zusammenfassungs-Hinweis,
  FEHLER-/Teilprüfung-Meldungen) — echte Umlaute im HTML, Text-Report transliteriert wie gehabt.
- **Offen (eigener PR vorgesehen):** Die Beschriftungen der ~60 Alt-Check-Funktionen sind
  noch transliteriert (geteilter ASCII-Text-Report). Voller Umlaut-Durchgang = Text-Report
  auf UTF-8 (BOM) umstellen + neu_text-Transliteration entfernen + Labels de-transliterieren.

**v5.0 PR „Paket B – Privilegien & ACLs" (Juni 2026):**
- Neuer Schalter `$privchk` (Default 1, in Whitelist) und Bereich „Privilegien & ACLs".
- Fünf read-only Checks: DCSync-Rechte (Replikations-ACEs am Domänenobjekt, GUIDs
  1131f6aa/1131f6ad), gefährliche Builtin-/Operatoren-Gruppen (Well-Known-SIDs S-1-5-32-548..551,
  Schema/Enterprise Admins, DnsAdmins), AdminSDHolder-ACL (Nicht-Standard-Schreibrechte),
  Protected Users (Nutzung, SID -525), Pre-Windows 2000 Compatible Access (Everyone/Anonymous,
  SID S-1-5-32-554).
- Sechs neue Katalogeinträge (privilegien, dcsync=Kritisch, operatoren, adminsdholder,
  protected_users, prewin2000) mit Hintergrund + live-verifizierten Links (MITRE T1003.006,
  MS Protected Users/Protected Accounts/Security Groups, Best practices for securing AD).
- Testsuite auf 69 Tests. **AD-Abfragen (Get-Acl AD:\, Get-ADGroupMember) nicht lokal
  testbar -> Funktionstest am DC.**

**v5.0 PR „Doku-Tiefe für die 22 Inventar-Checks" (Juni 2026):**
- Alle 22 Inventar-/Bestands-Checks auf Paket-A-Niveau gehoben: je ein Feld `Hintergrund`
  (Technik/Protokoll/Schwachstelle) und `Quellen` als Liste **live-verifizierter** Links.
- Damit haben jetzt **alle 28 Katalogeinträge** Hintergrund + Link-Quellen (45 HTTPS-Links).
- Jeder Link per WebFetch/WebSearch geprüft (erreichbar + Inhalt passend), u. a.: Microsoft
  Learn (Functional Levels, Central Store, Audit Policy, SYSVOL FRS→DFSR, Protected Accounts/
  AdminSDHolder, Security Groups, DNS Scavenging, gMSA, FGPP, Min Password Length, Security
  Baselines/SCT, Lifecycle FAQ, UserAccountControl), MITRE ATT&CK T1482/T1484, SpecterOps
  „Certified Pre-Owned".
- Testsuite auf 65 Tests: ein Test verlangt jetzt für **jeden** Katalogeintrag Hintergrund +
  HTTPS-Quell-Links mit Titel.

**Bugfix „srv_chk Zählung" (Juni 2026):**
- Server Check brach im echten DC-Lauf ab: `op_Subtraction` auf `ADPropertyValueCollection`.
  Ursache: `(Get-ADComputer … | Where-Object …).count` liefert bei **genau einem** Treffer
  keine Zahl, sondern eine leere `ADPropertyValueCollection` (AD-Objekte geben bei
  unbekanntem Member kein `$null`). Die anschließende Subtraktion/Arithmetik scheitert.
- Fix: alle Zählungen in `srv_chk` auf `@(…).Count` (erzwingt Array → echter Integer),
  Zähl-Abfragen auf `-Properties OperatingSystem` (statt `*`), die beiden Subtraktionen
  mit `[int]`-Cast. Regressionstest (case-sensitiv: kein `).count` in `srv_chk`).
- **Bekanntes latentes Thema:** dasselbe `(...).count`-Muster existiert noch in `clt_chk`,
  `sys_konten` und einigen `Get-Service`-Zählungen. Diese liefen im Test durch (0 oder ≥2
  Treffer), können bei genau 1 Treffer aber gleich fehlschlagen → bei Gelegenheit gleich
  mitziehen.

**v5.0 PR „Doku-Tiefe & verifizierte Quellen" (Juni 2026):**
- Neues Doku-Feld **`Hintergrund`** (Technik/Protokoll/Schwachstelle), im HTML als
  „Technischer Hintergrund" gerendert; optional (nur wo gepflegt).
- **`Quellen`** kann jetzt eine Liste `@{ Titel; Url }` sein → klickbare Links im HTML
  (`<a target=_blank rel=noopener noreferrer>`); String-Quellen weiterhin unterstützt
  (Rückwärtskompatibilität für die 22 Inventar-Checks).
- Die 6 Kerberos-Einträge (Paket A) sind voll ausgebaut: deutlich längerer Zweck,
  technischer Hintergrund (KDC/TGT/TGS, RC4-HMAC, UAC-Flags mit Hex, S4U/RBCD, noPac)
  und je 2–3 **live-verifizierte** Quell-Links. Jeder Link wurde per WebFetch geprüft
  (erreichbar + zitierter Inhalt tatsächlich enthalten): MITRE ATT&CK T1558/.003/.004,
  Microsoft Learn (SPN, UserAccountControl-Flags, KCD, Encryption Types, MAQ-Attribut),
  adsecurity.org (Metcalf), Shenanigans Labs (Shamir, RBCD), NVD CVE-2021-42278.
- Verifikation der LDAP-Filter durch die Microsoft-UAC-Quelle bestätigt: 0x80000
  (TRUSTED_FOR_DELEGATION), 0x200000 (USE_DES_KEY_ONLY), 0x400000 (DONT_REQ_PREAUTH).
- Testsuite auf 63 Tests (Hintergrund-Pflicht + HTTPS-Links je Kerberos-Eintrag,
  HTML rendert „Technischer Hintergrund" + `<a href>`).
- Offen: Die 22 Inventar-Checks behalten vorerst String-Quellen; gleiche Aufwertung
  (Hintergrund + verifizierte Links) folgt schrittweise.

**v5.0 Nachtrag „Umlaute & Encoding" (in Paket-A-PR enthalten):**
- Skript-Datei auf **UTF-8 mit BOM** umgestellt (PS 5.1 liest die Umlaute sonst falsch).
- Alle 28 Katalogtexte (Zweck/Beispiel/Empfehlung/Titel) nutzen jetzt echte Umlaute statt
  Transliteration. Wichtig: Nur Katalogtexte (HTML/JSON) — Bereichstitel und 2werte-Labels
  bleiben ASCII, da sie auch in den ASCII-Text-Report fließen.
- HTML-Überschriften nutzen den Katalog-Titel (DTitel, mit Umlauten) statt des ASCII-Titels,
  wenn eine Doku folgt.
- `Esc` escapet nur noch `< > & "`; Umlaute bleiben als echtes UTF-8 (keine `&#228;`-Entities).
- HTML-Report wird mit BOM geschrieben (robuste Umlaute auch außerhalb des Browsers).

**v5.0 PR „Paket A – Kerberos" (Juni 2026):**
- Neuer Schalter `$kerbchk` (Default 1, in Override-Whitelist) und neuer Bereich
  „Kerberos – Angriffsflaechen".
- Neuer Helfer `Unterpruefung ($titel,$checkid,$aktion)`: Teilprüfung innerhalb eines
  Bereichs mit eigener Begründung (`Doku`) und eigenem try/catch — ein fehlschlagender
  Sub-Check überspringt nur sich selbst. HTML badged jetzt auch Unter-Überschriften (h3)
  per Look-ahead und setzt Anker `#chk-<id>`.
- Fünf read-only Checks: Kerberoasting (SPN-Konten), AS-REP-Roasting (DONT_REQ_PREAUTH),
  Delegation (Unconstrained/Constrained/RBCD, DCs ausgenommen), schwache Verschlüsselung
  (UseDESKeyOnly), MachineAccountQuota. Alle über bitweise LDAP-Filter
  (`1.2.840.113556.1.4.803`).
- Sechs neue Katalogeinträge (kerberos, kerberoasting, asrep, delegation, kerb_enc,
  machine_quota) mit Quellen (MITRE ATT&CK, Microsoft, SpecterOps, CVE-2021-42278/42287).
- Testsuite auf 59 Tests erweitert. **Wichtig:** Die AD-Abfragen selbst sind hier nicht
  live testbar (kein AD-Modul) — Funktionstest erfolgt auf dem Test-DC.

**v5.0 PR „Doku-Framework" (Juni 2026):**
- Version auf 5.0 (Dateiname bleibt `Analyse_V4_6.ps1`).
- Zentraler `$CheckKatalog` (22 Einträge): je Prüfbereich Titel, **Schwere**
  (Info/Niedrig/Mittel/Hoch/Kritisch) und vier Begründungsfelder
  (Zweck, Beispiel, Empfehlung, Quellen) — ASCII-transliteriert wegen PS-5.1-Encoding.
- `Doku <id>` emittiert den Katalogeintrag als `Doku`-Ereignis (über `Merken`);
  `Pruefbereich` nimmt jetzt optional `-CheckId` und ruft `Doku` direkt nach `Bereich`.
  Alle 22 Hauptablauf-Blöcke sind damit verdrahtet.
- HTML-Report: **Executive Summary** oben (Bereiche nach Einstufung, mit Sprungmarken +
  Severity-Badges), je Bereich ein **einklappbarer** Block „Hintergrund & Empfehlung"
  (`<details>`), Severity-Badge an der Bereichs-Überschrift, Anker `#chk-<id>`.
  Neue CSS-Variable `--info-bg`, Badge-/Doku-/Zusammenfassungs-Styles.
- JSON-Export enthält die `Doku`-Ereignisse mit allen Begründungsfeldern.
- Testsuite auf 53 Tests erweitert (Katalog-Vollständigkeit, gültige Schweregrade,
  jeder Bereich hat gültige `-CheckId`, Doku-/HTML-/JSON-Rendering, Versionsprüfung).
- Geplant (eigene PRs): Paket A Kerberos, B Privilegien/ACLs, C AD CS (ESC1–8),
  D GPO/SYSVOL-Geheimnisse, E DC-Härtung, F Delta-Modus.

**PR „Export" (Juni 2026):**
- Strukturierte Erfassung: Die Formatierungsfunktionen melden jede Ausgabe als Ereignis an
  `Merken` (Arten: Kopf/Bereich/Titel/Subtitel/Wert/TabZeile/Text) → `$R_Daten`.
  Erfassung nur aktiv, wenn HTML oder JSON eingeschaltet ist; `neu_text` erfasst den
  Originaltext **vor** der ASCII-Umlaut-Ersetzung.
- **HTML-Report** (`HTML_Report`): eigenständige `.html` im water.css-inspirierten Stil
  (eingebettetes CSS, offline-fähig, hell/dunkel folgt der Systemeinstellung). Konsolen-
  Befund-Farben → CSS-Klassen (`Farbklasse`: Red→err, Green→ok, Yellow→warn);
  `FEHLER`-Texte aus `Pruefbereich` werden als rote Hinweis-Box gerendert; alle Inhalte
  HTML-escaped.
- **JSON-Export** (`JSON_Export`): `.json` mit Metadaten (System, Datum, Version) und der
  vollständigen Ereignisliste — Basis für spätere Auswertungen (AD-Ablösung).
- Neue Schalter `$A_Htm`/`$A_Jsn` (Standard 1) bzw. Parameter `-KeinHTML`/`-KeinJSON`.
- Testsuite auf 43 Tests erweitert (Ereignis-Erfassung, Farb-Mapping, HTML-Struktur inkl.
  Escaping, JSON-Gültigkeit).

**PR „Parametrisierung" (Juni 2026):**
- `[CmdletBinding()] param(...)`-Block: `-Verzeichnis`, `-Breite` (ValidateRange 70–90),
  `-KeineKonsole`, `-KeineDatei`, `-Bereiche` (Hashtable-Schalter-Overrides mit Whitelist
  der 31 bekannten Schalter; unbekannte Namen → Warnung, werden ignoriert).
- Overrides greifen nur bei explizit gebundenen Parametern (`$PSBoundParameters`) —
  die Kopf-Variablen bleiben die dokumentierten Standardwerte, 0/1/2-Semantik unverändert.
- Comment-Based Help um `.PARAMETER`- und `.EXAMPLE`-Abschnitte erweitert
  (`Get-Help .\Analyse_V4_6.ps1 -Detailed` funktioniert).
- Testsuite auf 37 Tests erweitert (param-Block via AST, Help-Inhalte, Integrationstest:
  Skriptstart im Subprozess parst Parameter und bricht ohne AD-Modul kontrolliert ab —
  wird auf Systemen mit AD-Modul automatisch übersprungen).

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
