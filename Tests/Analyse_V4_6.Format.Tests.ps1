#Requires -Version 5.1
<#
.SYNOPSIS
    Pester-Tests (Pester 5) fuer die Formatierungs-/Layout-Funktionen von Analyse_V4_6.ps1.

.DESCRIPTION
    Das Analyse-Skript wird bewusst NICHT ausgefuehrt (es wuerde sofort die AD-Pruefungen
    starten). Stattdessen werden die Funktionsdefinitionen per AST aus der Skriptdatei
    extrahiert und in einer kontrollierten Umgebung getestet: feste Ausgabebreite,
    Konsolenausgabe aus, Datei-Ausgabe in eine Temp-Datei.

    Diese Tests sind das Sicherheitsnetz fuer kommende Refactorings (Fehlerbehandlung,
    I/O-Buendelung, Parametrisierung): Das Report-Layout darf sich dabei nicht aendern.

.NOTES
    Ausfuehren:  Invoke-Pester -Path .\Tests -Output Detailed
#>

Describe 'Analyse_V4_6.ps1' {

    BeforeAll {
        $skriptPfad = Join-Path (Split-Path -Parent $PSScriptRoot) 'Analyse_V4_6.ps1'

        ### Skript parsen (ohne Ausfuehrung) ########################################################
        $tokens = $null ; $parseFehler = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $skriptPfad, [ref]$tokens, [ref]$parseFehler)
        if ($parseFehler.Count -gt 0) {
            throw "Analyse_V4_6.ps1 laesst sich nicht parsen: $($parseFehler[0])"
        }

        ### Formatierungs-Funktionen extrahieren und global definieren #############################
        $zielFunktionen = @('Header','Bottom','Vollzeile','Leerzeile','Trennzeile','tablinie',
                            'Bereich','Bereichstitel','Subtitel','2werte','new_2werte',
                            'neu_tab_max6w_fb','neu_text','Pruefbereich','Unterpruefung',
                            'Ausgabe','Puffer_leeren',
                            'Merken','Doku','Farbklasse','HTML_Report','JSON_Export')
        $gefunden = $ast.FindAll({
            param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true) | Where-Object { $zielFunktionen -contains $_.Name }

        foreach ($f in $gefunden) {
            # "function X ..." -> "function global:X ...", damit die Definition
            # die Pester-Scopes ueberlebt und in den It-Bloecken aufrufbar ist.
            $definition = $f.Extent.Text -replace '^function\s+', 'function global:'
            Invoke-Expression $definition
        }

        ### Check-Katalog (Top-Level-Zuweisung) extrahieren und global bereitstellen ################
        $katAst = $ast.FindAll({
            param($a) $a -is [System.Management.Automation.Language.AssignmentStatementAst]
        }, $true) | Where-Object { $_.Left.Extent.Text -eq '$CheckKatalog' } | Select-Object -First 1
        Invoke-Expression ($katAst.Extent.Text -replace '^\$CheckKatalog', '$global:CheckKatalog')

        ### Kontrollierte Umgebung: gleiche Variablen wie im Skript-Kopf ###########################
        $global:sb           = 90
        $global:zeichen      = '*'
        $global:tabzeichen   = '-'
        $global:tabtrenner   = '|'
        $global:leer         = ' '
        $global:type         = 'Report'
        $global:maintitel    = 'AD-Analyse'
        $global:firma        = 'Vers. 4.6 Test'     # bewusst <= 16 Zeichen (Header-Feldbreite)
        $global:madeby       = 'AD-Assessment Tool'
        $global:datum        = '01-01-2026 12:00'   # fix, damit der Header deterministisch ist
        $global:system       = 'TESTSYS'
        $global:F_Rahmen     = 'DarkYellow'
        $global:F_Ue_Schrift = 'Gray'
        $global:F_Text       = 'White'
        $global:F_Fehler     = 'Red'
        $global:A_Con        = 0                    # keine Konsolenausgabe in den Tests
        $global:A_Dat        = 1                    # Datei-Ausgabe in Temp-Datei
        $global:path         = Join-Path ([IO.Path]::GetTempPath()) 'ad_assessment_format_tests.txt'
        $global:A_Puffer     = New-Object System.Text.StringBuilder   # Ausgabepuffer wie im Skript
        $global:A_Htm        = 1                    # Ereignis-Erfassung fuer HTML aktiv
        $global:A_Jsn        = 1                    # Ereignis-Erfassung fuer JSON aktiv
        $global:R_Daten      = New-Object 'System.Collections.Generic.List[object]'
        $global:version      = 'Vers. 5.0'

        function global:Get-ReportZeilen {
            # Erst den Puffer in die Datei schreiben (gepufferte Ausgabe seit Performance-PR),
            # dann lesen. Komma-Operator: verhindert, dass die Pipeline ein 1-Zeilen-Array
            # zum einzelnen String entrollt ($z[0] waere sonst ein [char]).
            Puffer_leeren
            , @(Get-Content -LiteralPath $global:path)
        }
    }

    AfterAll {
        Remove-Item -LiteralPath $global:path -Force -ErrorAction SilentlyContinue
        foreach ($n in @('Header','Bottom','Vollzeile','Leerzeile','Trennzeile','tablinie',
                         'Bereich','Bereichstitel','Subtitel','2werte','new_2werte',
                         'neu_tab_max6w_fb','neu_text','Pruefbereich','Unterpruefung',
                         'Ausgabe','Puffer_leeren',
                         'Merken','Doku','Farbklasse','HTML_Report','JSON_Export','Get-ReportZeilen')) {
            Remove-Item -LiteralPath "function:global:$n" -Force -ErrorAction SilentlyContinue
        }
        foreach ($v in @('sb','zeichen','tabzeichen','tabtrenner','leer','type','maintitel',
                         'firma','madeby','datum','system','F_Rahmen','F_Ue_Schrift','F_Text',
                         'F_Fehler','A_Con','A_Dat','path','A_Puffer','A_Htm','A_Jsn',
                         'R_Daten','version','CheckKatalog')) {
            Remove-Variable -Name $v -Scope Global -Force -ErrorAction SilentlyContinue
        }
    }

    BeforeEach {
        if (Test-Path -LiteralPath $global:path) { Clear-Content -LiteralPath $global:path }
        else { New-Item -ItemType File -Path $global:path | Out-Null }
        [void]$global:A_Puffer.Clear()
        $global:R_Daten.Clear()
    }

    Context 'Skript-Datei (statische Pruefungen)' {

        It 'laesst sich ohne Syntaxfehler parsen' {
            $parseFehler.Count | Should -Be 0
        }

        It 'beginnt mit #Requires -Version' {
            (Get-Content -LiteralPath $skriptPfad -TotalCount 1) |
                Should -Match '^#Requires -Version'
        }

        It 'fuehrt durchgaengig Version 5.0' {
            $inhalt = Get-Content -LiteralPath $skriptPfad -Raw
            $inhalt | Should -Match '\$version = "Vers\. 5\.0"'
            $inhalt | Should -Match 'Version       : 5\.0'
        }

        It 'ist als UTF-8 mit BOM gespeichert (korrekte Umlaute auch unter PS 5.1)' {
            $bytes = [System.IO.File]::ReadAllBytes($skriptPfad)
            $bytes[0] | Should -Be 0xEF
            $bytes[1] | Should -Be 0xBB
            $bytes[2] | Should -Be 0xBF
        }

        It 'enthaelt die Modul-Vorabpruefung' {
            $inhalt = Get-Content -LiteralPath $skriptPfad -Raw
            $inhalt | Should -Match 'Get-Module -ListAvailable -Name ActiveDirectory'
            $inhalt | Should -Match 'Get-Module -ListAvailable -Name GroupPolicy'
            $inhalt | Should -Match 'Get-Module -ListAvailable -Name DnsServer'
        }

        It 'Hauptablauf nutzt Pruefbereich (kein ungeschuetzter Bereich-Aufruf auf Top-Level)' {
            $bereichAufrufe = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst]
            }, $true) | Where-Object { $_.GetCommandName() -eq 'Bereich' } | Where-Object {
                $p = $_.Parent ; $inFunktion = $false
                while ($p) {
                    if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                        $inFunktion = $true ; break
                    }
                    $p = $p.Parent
                }
                -not $inFunktion
            }
            $bereichAufrufe | Should -BeNullOrEmpty
        }

        It 'hat einen param-Block mit den erwarteten Parametern' {
            $parameter = $ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }
            foreach ($p in @('Verzeichnis','Breite','KeineKonsole','KeineDatei','Bereiche')) {
                $parameter | Should -Contain $p
            }
        }

        It 'dokumentiert die Parameter in der Comment-Based Help' {
            $inhalt = Get-Content -LiteralPath $skriptPfad -Raw
            $inhalt | Should -Match '\.PARAMETER Verzeichnis'
            $inhalt | Should -Match '\.PARAMETER Bereiche'
            $inhalt | Should -Match '\.EXAMPLE'
        }

        It 'startet, parst Parameter und bricht ohne AD-Modul kontrolliert ab' {
            if (Get-Module -ListAvailable ActiveDirectory) {
                Set-ItResult -Skipped -Because 'AD-Modul vorhanden - der Lauf wuerde echte AD-Pruefungen starten'
                return
            }
            $ausgabe = & powershell.exe -NoProfile -Command "& '$skriptPfad' -KeineDatei -Breite 80 -Bereiche @{ unbekannt = 1 }" 2>&1 | Out-String
            $LASTEXITCODE | Should -Be 1
            $ausgabe | Should -Match "Unbekannter Bereichs-Schalter 'unbekannt'"
            $ausgabe | Should -Match 'ActiveDirectory'
        }

        It 'buendelt die Datei-Ausgabe (genau eine Add-Content-Stelle: Puffer_leeren)' {
            $addContent = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst]
            }, $true) | Where-Object { $_.GetCommandName() -eq 'Add-Content' }
            $addContent.Count | Should -Be 1
        }

        It 'enthaelt keine schreibenden AD-Cmdlets (Read-only-Konvention)' {
            $schreibend = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst]
            }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object {
                $_ -match '^(Set|New|Remove|Disable|Enable|Unlock|Move|Rename|Add|Grant|Revoke|Clear|Reset)-AD[A-Za-z]'
            }
            $schreibend | Should -BeNullOrEmpty
        }
    }

    Context 'Rahmen-Funktionen' {

        It 'Vollzeile: eine Zeile, volle Breite, nur Rahmenzeichen' {
            Vollzeile
            $z = Get-ReportZeilen
            $z.Count | Should -Be 1
            $z[0] | Should -BeExactly ('*' * 90)
        }

        It 'Leerzeile: Rahmen links/rechts, innen leer' {
            Leerzeile
            $z = Get-ReportZeilen
            $z.Count | Should -Be 1
            $z[0] | Should -BeExactly ('*' + (' ' * 88) + '*')
        }

        It 'Trennzeile (Standard): volle Trennlinie mit Rahmen' {
            Trennzeile
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('* ' + ('*' * 86) + ' *')
        }

        It 'Trennzeile (Sub): eingerueckte Trennlinie' {
            Trennzeile 's'
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('*  ' + ('*' * 85) + ' *')
        }

        It 'tablinie (Standard): Tabellen-Querlinie' {
            tablinie 'n'
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('* ' + ('-' * 86) + ' *')
        }

        It 'tablinie (Sub): eingerueckte Tabellen-Querlinie' {
            tablinie 's'
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('*  ' + ('-' * 85) + ' *')
        }
    }

    Context 'Titel-Funktionen' {

        It 'Bereich: drei Zeilen, Titel mittig im Sternrahmen' {
            Bereich 'Test Bereich'
            $z = Get-ReportZeilen
            $z.Count | Should -Be 3
            $z[0] | Should -BeExactly ('*' * 90)
            $z[1] | Should -BeExactly ('*** Test Bereich' + (' ' * 70) + ' ***')
            $z[2] | Should -BeExactly ('*' * 90)
        }

        It 'Bereichstitel: Titel plus Unterstreichung aus Rahmenzeichen' {
            Bereichstitel 'Mein Titel'
            $z = Get-ReportZeilen
            $z.Count | Should -Be 2
            $z[0] | Should -BeExactly ('* Mein Titel' + (' ' * 76) + ' *')
            $z[1] | Should -BeExactly ('* ' + ('*' * 10) + (' ' * 76) + ' *')
        }

        It 'Bereichstitel (Sub): eingerueckter Titel' {
            Bereichstitel 'Titel' 's'
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('*  Titel' + (' ' * 80) + ' *')
            $z[1] | Should -BeExactly ('*  ' + ('*' * 5) + (' ' * 80) + ' *')
        }

        It 'Subtitel: Standard-Unterstreichung mit "*"' {
            Subtitel 'Unterpunkt'
            $z = Get-ReportZeilen
            $z.Count | Should -Be 2
            $z[0] | Should -BeExactly ('* Unterpunkt' + (' ' * 76) + ' *')
            $z[1] | Should -BeExactly ('* ' + ('*' * 10) + (' ' * 76) + ' *')
        }

        It 'Subtitel: Einrueckung und eigenes Unterstreichungszeichen' {
            Subtitel 'Unterpunkt' 4 '='
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('* ' + (' ' * 4) + 'Unterpunkt' + (' ' * 72) + ' *')
            $z[1] | Should -BeExactly ('* ' + (' ' * 4) + ('=' * 10) + (' ' * 72) + ' *')
        }
    }

    Context 'Wert-Funktionen' {

        It '2werte: Beschriftung und Wert linksbuendig im Rahmen' {
            2werte 'Name:' 'Wert42'
            $z = Get-ReportZeilen
            $z.Count | Should -Be 1
            $z[0] | Should -BeExactly ('* Name: Wert42' + (' ' * 74) + ' *')
        }

        It '2werte (Sub): eingerueckt' {
            2werte 'Name:' 'Wert42' 's'
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('*  Name: Wert42' + (' ' * 73) + ' *')
        }

        It 'new_2werte: feste Spaltenbreite, beide Werte linksbuendig' {
            new_2werte $null ':' 20 'Schluessel' $null 'l' 'Wert' $null 'l'
            $z = Get-ReportZeilen
            $z.Count | Should -Be 1
            $z[0] | Should -BeExactly ('* ' + 'Schluessel'.PadRight(20) + ': ' + 'Wert'.PadRight(64) + ' *')
        }

        It 'new_2werte: zu langer erster Wert wird mit "~" gekuerzt' {
            new_2werte $null $null 8 'SehrLangerWert' $null 'l' 'Wert' $null 'l'
            $z = Get-ReportZeilen
            $z[0] | Should -BeExactly ('* ' + 'SehrLan~' + ' ' + 'Wert'.PadRight(77) + ' *')
        }
    }

    Context 'Tabellen-Funktion neu_tab_max6w_fb' {

        It '3 Spalten rechts: Text rechtsbuendig, Zahlen rechts in der Spalte' {
            neu_tab_max6w_fb -spa 3 -pos 'r' -sub 'n' -bre 10 -wex 'Hinweis' -we1 'Eins' -we2 'Zwei' -we3 '33'
            $z = Get-ReportZeilen
            $z.Count | Should -Be 1
            $erwartet = '* ' + 'Eins'.PadRight(10) + ' | ' + 'Zwei'.PadRight(10) + ' | ' +
                        '33'.PadLeft(10) + ' | ' + 'Hinweis'.PadRight(47) + ' *'
            $z[0] | Should -BeExactly $erwartet
        }

        It '2 Spalten links: Freitext-Feld vor den Spalten' {
            neu_tab_max6w_fb -spa 2 -pos 'l' -sub 'n' -bre 12 -wex 'Beschreibung' -we1 'WertA' -we2 'WertB'
            $z = Get-ReportZeilen
            $z.Count | Should -Be 1
            $z[0].Length | Should -Be 90
            $z[0] | Should -Match '^\* Beschreibung'
            $z[0] | Should -Match 'WertA\s+\| WertB\s+ \*$'
        }
    }

    Context 'Text-Funktion neu_text' {

        It 'Ueberschrift mit Unterstreichung, Text wird umbrochen, Zeilen exakt Reportbreite' {
            $langerText = 'Die Pruefung der Umgebung erfolgt ueber mehrere Schritte und dieser ' +
                          'Text ist absichtlich lang genug damit mindestens ein Zeilenumbruch entsteht.'
            neu_text 0 '=' 'Pruefbereich' $langerText
            $z = Get-ReportZeilen
            $z.Count | Should -BeGreaterOrEqual 4      # Ueberschrift + Unterstrich + >= 2 Textzeilen
            $z[0] | Should -BeExactly ('* ' + 'Pruefbereich'.PadRight(86) + ' *')
            $z[1] | Should -BeExactly ('* ' + ('=' * 12).PadRight(86) + ' *')
            foreach ($zeile in $z) {
                $zeile.Length | Should -Be 90
                $zeile | Should -Match '^\* .* \*$'
            }
        }

        It 'ersetzt Umlaute durch ue/ae/oe (ASCII-Report)' {
            neu_text 0 '-' 'Prüfung' 'Die Lösung wäre über kürzere Wörter möglich.'
            $inhalt = (Get-ReportZeilen) -join ' '
            $inhalt | Should -Match 'Pruefung'
            $inhalt | Should -Match 'Loesung'
            $inhalt | Should -Match 'waere'
            $inhalt | Should -Match 'moeglich'
            $inhalt | Should -Not -Match '[üäöÜÄÖ]'
        }
    }

    Context 'Fehlerbehandlung (Pruefbereich)' {

        It 'fuehrt die Aktion innerhalb des Bereichsrahmens aus' {
            Pruefbereich 'Testbereich' { Leerzeile }
            $z = Get-ReportZeilen
            $z.Count | Should -Be 4
            $z[0] | Should -BeExactly ('*' * 90)
            $z[1] | Should -Match '^\*\*\* Testbereich'
            $z[3] | Should -BeExactly ('*' + (' ' * 88) + '*')
            ($z -join ' ') | Should -Not -Match 'FEHLER'
        }

        It 'faengt Fehler ab, vermerkt sie im Report und wirft nicht weiter' {
            { Pruefbereich 'Testbereich' { throw 'Absichtlicher Testfehler' } } | Should -Not -Throw
            $zeilen = Get-ReportZeilen
            $inhalt = $zeilen -join ' '
            $inhalt | Should -Match 'FEHLER - Bereich nur unvollstaendig geprueft'
            $inhalt | Should -Match 'Absichtlicher Testfehler'
            $inhalt | Should -Match 'fortgesetzt'
            foreach ($zeile in $zeilen) { $zeile.Length | Should -Be 90 }
        }

        It 'nach einem Fehler laeuft die naechste Ausgabe normal weiter' {
            Pruefbereich 'Testbereich' { throw 'Absichtlicher Testfehler' }
            Vollzeile
            $z = Get-ReportZeilen
            $z[-1] | Should -BeExactly ('*' * 90)
        }
    }

    Context 'Strukturierte Erfassung und Exporte (HTML/JSON)' {

        It '2werte und Bereich erzeugen strukturierte Ereignisse' {
            Bereich 'Mein Bereich'
            2werte 'Schluessel:' 'Wert123' $null 'Red'
            $global:R_Daten.Count | Should -Be 2
            $global:R_Daten[0].Art | Should -Be 'Bereich'
            $global:R_Daten[0].Titel | Should -Be 'Mein Bereich'
            $global:R_Daten[1].Art | Should -Be 'Wert'
            $global:R_Daten[1].Name | Should -Be 'Schluessel:'
            $global:R_Daten[1].Wert | Should -Be 'Wert123'
            $global:R_Daten[1].Farbe | Should -Be 'Red'
        }

        It 'neu_text erfasst Original-Text mit Umlauten (vor der ASCII-Ersetzung)' {
            neu_text 0 '-' 'Prüfung' 'Die Lösung ist über kurze Wörter möglich.'
            $global:R_Daten[0].Ueberschrift | Should -Be 'Prüfung'
            $global:R_Daten[0].Text | Should -Match 'Lösung'
        }

        It 'keine Erfassung wenn HTML und JSON deaktiviert sind' {
            $global:A_Htm = 0 ; $global:A_Jsn = 0
            try { Bereich 'Test' } finally { $global:A_Htm = 1 ; $global:A_Jsn = 1 }
            $global:R_Daten.Count | Should -Be 0
        }

        It 'Farbklasse mappt Konsolenfarben auf Befund-Klassen' {
            Farbklasse 'Red'        | Should -Be 'err'
            Farbklasse 'DarkRed'    | Should -Be 'err'
            Farbklasse 'Green'      | Should -Be 'ok'
            Farbklasse 'DarkYellow' | Should -Be 'warn'
            Farbklasse 'White'      | Should -Be ''
            Farbklasse $null        | Should -Be ''
        }

        It 'HTML_Report erzeugt eigenstaendigen Report mit Struktur, Status-Klassen und Escaping' {
            Header
            Bereich 'Domain, Mode, FSMO'
            Bereichstitel 'Details'
            2werte 'Status:' 'kritisch <b>!</b>' $null 'Red'
            neu_tab_max6w_fb -spa 2 -pos 'r' -sub 'n' -bre 10 -wex 'Hinweis' -we1 'DC01' -we2 'online' -we4 'Green'
            neu_text 0 '!' 'FEHLER - Bereich nur unvollstaendig geprueft' 'Meldung: Testfehler. Der Lauf wird fortgesetzt.'
            HTML_Report
            $htmlPfad = [System.IO.Path]::ChangeExtension($global:path, 'html')
            Test-Path $htmlPfad | Should -BeTrue
            $html = Get-Content -LiteralPath $htmlPfad -Raw
            $html | Should -Match '<meta charset="utf-8">'
            $html | Should -Match '<style>'                               # CSS eingebettet
            $html | Should -Match 'prefers-color-scheme'                  # hell/dunkel automatisch
            $html | Should -Match '<h2>Domain, Mode, FSMO</h2>'
            $html | Should -Match '<h3>Details</h3>'
            $html | Should -Match '<td class="err">kritisch &lt;b&gt;!&lt;/b&gt;</td>'   # Escaping!
            $html | Should -Match '<td class="ok">online</td>'
            $html | Should -Match 'class="fehler"'
            $html | Should -Match 'TESTSYS'
            $html | Should -Not -Match '<b>!</b>'                         # nichts unescaped
            Remove-Item $htmlPfad -Force
        }

        It 'JSON_Export schreibt gueltiges JSON mit Metadaten und Ereignissen' {
            Bereich 'Testbereich'
            2werte 'Anzahl:' '42'
            JSON_Export
            $jsonPfad = [System.IO.Path]::ChangeExtension($global:path, 'json')
            Test-Path $jsonPfad | Should -BeTrue
            $daten = Get-Content -LiteralPath $jsonPfad -Raw | ConvertFrom-Json
            $daten.System | Should -Be 'TESTSYS'
            $daten.Version | Should -Be 'Vers. 5.0'
            @($daten.Ereignisse).Count | Should -Be 2
            @($daten.Ereignisse)[1].Wert | Should -Be '42'
            Remove-Item $jsonPfad -Force
        }
    }

    Context 'Doku-Framework (v5.0): Katalog, Severity, Begruendung' {

        It 'Katalog enthaelt Eintraege und alle Pflichtfelder sind gefuellt' {
            $global:CheckKatalog.Count | Should -BeGreaterOrEqual 22
            foreach ($id in $global:CheckKatalog.Keys) {
                foreach ($feld in 'Titel','Schwere','Zweck','Beispiel','Empfehlung') {
                    [string]::IsNullOrWhiteSpace($global:CheckKatalog[$id].$feld) |
                        Should -BeFalse -Because "$id.$feld darf nicht leer sein"
                }
                # Quellen ist entweder ein String oder eine Liste @{ Titel; Url }
                $q = $global:CheckKatalog[$id].Quellen
                if ($q -is [string]) {
                    [string]::IsNullOrWhiteSpace($q) | Should -BeFalse -Because "$id.Quellen leer"
                } else {
                    @($q).Count | Should -BeGreaterThan 0 -Because "$id.Quellen leer"
                    foreach ($l in @($q)) {
                        $l.Titel | Should -Not -BeNullOrEmpty -Because "$id Quelle ohne Titel"
                        $l.Url   | Should -Match '^https?://' -Because "$id Quelle ohne gueltige URL"
                    }
                }
            }
        }

        It 'Katalog nutzt nur gueltige Schweregrade' {
            $gueltig = 'Info','Niedrig','Mittel','Hoch','Kritisch'
            foreach ($id in $global:CheckKatalog.Keys) {
                $gueltig | Should -Contain $global:CheckKatalog[$id].Schwere
            }
        }

        It 'Katalog-Texte verwenden echte Umlaute (nicht transliteriert)' {
            $ae = [char]0x00E4   # ae
            $global:CheckKatalog['domain_allgemein'].Zweck | Should -Match $ae
            $global:CheckKatalog['kerberos'].Titel | Should -Match ('Angriffsfl' + $ae + 'chen')
            # Quellen duerfen nicht faelschlich umgewandelt sein
            $global:CheckKatalog['domain_allgemein'].Quellen | Should -Match 'Quellen|Microsoft'
        }

        It 'jeder Pruefbereich-Aufruf im Hauptablauf traegt eine -CheckId' {
            $aufrufe = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst] -and
                          $a.GetCommandName() -eq 'Pruefbereich'
            }, $true)
            $aufrufe.Count | Should -BeGreaterOrEqual 22
            foreach ($cmd in $aufrufe) {
                $hatCheckId = $cmd.CommandElements | Where-Object {
                    $_ -is [System.Management.Automation.Language.CommandParameterAst] -and
                    $_.ParameterName -eq 'CheckId'
                }
                $hatCheckId | Should -Not -BeNullOrEmpty -Because "Bereich '$($cmd.CommandElements[1].Extent.Text)' braucht eine CheckId"
            }
        }

        It 'jede im Hauptablauf verwendete CheckId existiert im Katalog' {
            $aufrufe = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst] -and
                          $a.GetCommandName() -eq 'Pruefbereich'
            }, $true)
            foreach ($cmd in $aufrufe) {
                $els = $cmd.CommandElements
                for ($i = 0; $i -lt $els.Count; $i++) {
                    if ($els[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and
                        $els[$i].ParameterName -eq 'CheckId') {
                        $id = if ($els[$i].Argument) { $els[$i].Argument.Value } else { $els[$i + 1].Value }
                        $global:CheckKatalog.Keys | Should -Contain $id
                    }
                }
            }
        }

        It 'Doku emittiert ein Ereignis mit allen Begruendungsfeldern' {
            Doku 'admins'
            $global:R_Daten.Count | Should -Be 1
            $d = $global:R_Daten[0]
            $d.Art | Should -Be 'Doku'
            $d.CheckId | Should -Be 'admins'
            $d.Schwere | Should -Be 'Hoch'
            $d.Zweck | Should -Not -BeNullOrEmpty
            $d.Empfehlung | Should -Not -BeNullOrEmpty
            $d.Quellen | Should -Not -BeNullOrEmpty
        }

        It 'Doku bei unbekannter ID erzeugt kein Ereignis' {
            Doku 'gibt_es_nicht'
            $global:R_Daten.Count | Should -Be 0
        }

        It 'Pruefbereich mit CheckId schreibt Bereich + Doku in die Ereignisliste' {
            Pruefbereich 'Testbereich' -CheckId 'dns' { Leerzeile }
            $arten = $global:R_Daten | ForEach-Object { $_.Art }
            $arten | Should -Contain 'Bereich'
            $arten | Should -Contain 'Doku'
            ($global:R_Daten | Where-Object { $_.Art -eq 'Doku' }).CheckId | Should -Be 'dns'
        }

        It 'HTML_Report rendert einklappbaren Doku-Block, Severity-Badge und Zusammenfassung' {
            Header
            Pruefbereich 'Administratoren und Builtin Benutzer' -CheckId 'admins' {
                2werte 'Domain Admins:' '5 Mitglieder' $null 'Red'
            }
            Pruefbereich 'DNS - Check' -CheckId 'dns' { 2werte 'Scavenging:' 'aus' $null 'Yellow' }
            HTML_Report
            $htmlPfad = [System.IO.Path]::ChangeExtension($global:path, 'html')
            $html = Get-Content -LiteralPath $htmlPfad -Raw
            $html | Should -Match '<details class="doku">'                       # einklappbar
            $html | Should -Match '<summary>Hintergrund &amp; Empfehlung</summary>'
            $html | Should -Match '<span class="lbl">Zweck:</span>'
            $html | Should -Match '<span class="lbl">Empfehlung:</span>'
            $html | Should -Match 'badge sev-hoch'                               # Severity-Badge
            $html | Should -Match 'id="chk-admins"'                             # Anker
            $html | Should -Match '<section class="zus">'                        # Zusammenfassung
            $html | Should -Match 'href="#chk-dns"'                             # Sprungmarke
            $html | Should -Match ([char]0x00E4)                                 # echte Umlaute im HTML
            Remove-Item $htmlPfad -Force -ErrorAction SilentlyContinue
        }

        It 'JSON enthaelt das Doku-Ereignis mit Begruendungsfeldern' {
            Pruefbereich 'Testbereich' -CheckId 'ca' { 2werte 'CA:' 'vorhanden' }
            JSON_Export
            $jsonPfad = [System.IO.Path]::ChangeExtension($global:path, 'json')
            $daten = Get-Content -LiteralPath $jsonPfad -Raw | ConvertFrom-Json
            $doku = @($daten.Ereignisse | Where-Object { $_.Art -eq 'Doku' })
            $doku.Count | Should -Be 1
            $doku[0].DTitel | Should -Be 'Zertifizierungsstelle(n)'
            $doku[0].Empfehlung | Should -Not -BeNullOrEmpty
            Remove-Item $jsonPfad -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'Paket A (v5.0): Kerberos-Angriffsflaechen' {

        It 'Katalog enthaelt alle Kerberos-Eintraege' {
            foreach ($id in 'kerberos','kerberoasting','asrep','delegation','kerb_enc','machine_quota') {
                $global:CheckKatalog.Keys | Should -Contain $id
            }
        }

        It 'Kerberos-Eintraege haben technischen Hintergrund und verifizierte Quell-Links' {
            foreach ($id in 'kerberos','kerberoasting','asrep','delegation','kerb_enc','machine_quota') {
                $k = $global:CheckKatalog[$id]
                [string]::IsNullOrWhiteSpace($k.Hintergrund) | Should -BeFalse -Because "$id braucht Hintergrund"
                $k.Hintergrund.Length | Should -BeGreaterThan 150 -Because "$id Hintergrund zu kurz"
                @($k.Quellen).Count | Should -BeGreaterOrEqual 2 -Because "$id braucht mehrere Quellen"
                foreach ($l in @($k.Quellen)) {
                    $l.Url | Should -Match '^https://' -Because "$id Quelle muss HTTPS-Link sein"
                    $l.Titel | Should -Not -BeNullOrEmpty
                }
            }
        }

        It 'HTML rendert Technischer Hintergrund und klickbare Quell-Links' {
            Unterpruefung 'Kerberoasting' 'kerberoasting' { 2werte 'SPN-Konten:' '1' 's' 'Red' }
            HTML_Report
            $htmlPfad = [System.IO.Path]::ChangeExtension($global:path, 'html')
            $html = Get-Content -LiteralPath $htmlPfad -Raw
            $html | Should -Match 'Technischer Hintergrund:'
            $html | Should -Match '<a href="https://attack\.mitre\.org/techniques/T1558/003/"'
            $html | Should -Match 'target="_blank" rel="noopener noreferrer"'
            Remove-Item $htmlPfad -Force -ErrorAction SilentlyContinue
        }

        It 'die Kerberos-Prueffunktionen sind im Skript definiert' {
            $funktionen = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true) | ForEach-Object { $_.Name }
            foreach ($fn in 'chk_kerberoasting','chk_asrep','chk_delegation','chk_kerb_enc','chk_machine_quota') {
                $funktionen | Should -Contain $fn
            }
        }

        It 'der Schalter kerbchk steht in der Override-Whitelist' {
            (Get-Content -LiteralPath $skriptPfad -Raw) | Should -Match "'kerbchk'"
        }

        It 'jede Unterpruefung-CheckId existiert im Katalog' {
            $aufrufe = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst] -and
                          $a.GetCommandName() -eq 'Unterpruefung'
            }, $true)
            $aufrufe.Count | Should -BeGreaterOrEqual 5
            foreach ($cmd in $aufrufe) {
                # Positional: [0]=Name der Funktion, [1]=Titel, [2]=CheckId, [3]=Scriptblock
                $idEl = $cmd.CommandElements[2]
                $global:CheckKatalog.Keys | Should -Contain $idEl.Value
            }
        }

        It 'Unterpruefung erzeugt Titel + Doku und rendert eine Unterueberschrift mit Badge' {
            Unterpruefung 'Kerberoasting (Konten mit SPN)' 'kerberoasting' {
                2werte 'Benutzerkonten mit SPN:' '3' 's' 'Red'
            }
            $arten = $global:R_Daten | ForEach-Object { $_.Art }
            $arten | Should -Contain 'Titel'
            $arten | Should -Contain 'Doku'
            ($global:R_Daten | Where-Object { $_.Art -eq 'Doku' }).CheckId | Should -Be 'kerberoasting'
            HTML_Report
            $htmlPfad = [System.IO.Path]::ChangeExtension($global:path, 'html')
            $html = Get-Content -LiteralPath $htmlPfad -Raw
            $html | Should -Match '<h3 id="chk-kerberoasting">'
            $html | Should -Match 'badge sev-hoch'
            Remove-Item $htmlPfad -Force -ErrorAction SilentlyContinue
        }

        It 'Unterpruefung faengt Fehler in der Teilpruefung ab (Bereich laeuft weiter)' {
            { Unterpruefung 'Testpruefung' 'kerberoasting' { throw 'Simulierter Abfragefehler' } } |
                Should -Not -Throw
            $txt = (Get-ReportZeilen) -join ' '
            $txt | Should -Match 'Teilpruefung uebersprungen'
            $txt | Should -Match 'Simulierter Abfragefehler'
        }
    }

    Context 'Gepufferte Datei-Ausgabe (Puffer)' {

        It 'Ausgaben landen erst mit Puffer_leeren in der Datei' {
            Vollzeile
            @(Get-Content -LiteralPath $global:path) | Should -BeNullOrEmpty
            $global:A_Puffer.Length | Should -BeGreaterThan 0
            Puffer_leeren
            @(Get-Content -LiteralPath $global:path).Count | Should -Be 1
        }

        It 'Puffer_leeren leert den Puffer (kein doppelter Inhalt)' {
            Vollzeile
            Puffer_leeren
            Puffer_leeren
            @(Get-Content -LiteralPath $global:path).Count | Should -Be 1
            $global:A_Puffer.Length | Should -Be 0
        }

        It 'Pruefbereich schreibt den Puffer auch im Fehlerfall in die Datei' {
            Pruefbereich 'Testbereich' { throw 'Absichtlicher Testfehler' }
            # Direkt lesen, ohne Get-ReportZeilen (das wuerde selbst flushen):
            $inhalt = @(Get-Content -LiteralPath $global:path) -join ' '
            $inhalt | Should -Match 'Absichtlicher Testfehler'
        }
    }

    Context 'Header und Bottom' {

        It 'Header: fuenf Zeilen, Eckdaten an den richtigen Positionen (auch ohne Konsole)' {
            # Seit dem Fehlerbehandlungs-PR ist die Dateiausgabe des Headers
            # unabhaengig von der Konsolenausgabe ($A_Con bleibt hier 0).
            Header
            $z = Get-ReportZeilen
            $z.Count | Should -Be 5
            $z[0] | Should -BeExactly ('*' * 90)
            $z[4] | Should -BeExactly ('*' * 90)
            $z[1] | Should -Match 'Report'
            $z[1] | Should -Match '01-01-2026 12:00'
            $z[2] | Should -Match 'AD-Analyse'
            $z[3] | Should -Match 'TESTSYS'
            $z[3] | Should -Match 'Vers\. 4\.6 Test'
            foreach ($zeile in $z) { $zeile.Length | Should -Be 90 }
        }

        It 'Header: zu lange Firma wird gekuerzt statt den Rahmen zu sprengen' {
            $global:firma = 'Vers. 4.6 AD-Assessment'   # 23 Zeichen - der reale Standardwert
            try { Header } finally { $global:firma = 'Vers. 4.6 Test' }
            $z = Get-ReportZeilen
            $z.Count | Should -Be 5
            foreach ($zeile in $z) { $zeile.Length | Should -Be 90 }
            $z[3] | Should -Match ([regex]::Escape('Vers. 4.6 AD-As~'))
        }

        It 'Bottom: drei Zeilen, Ersteller mittig' {
            Bottom
            $z = Get-ReportZeilen
            $z.Count | Should -Be 3
            $z[0] | Should -BeExactly ('*' * 90)
            $z[2] | Should -BeExactly ('*' * 90)
            $z[1] | Should -BeExactly (('*' * 34) + ' ' + 'AD-Assessment Tool'.PadRight(20) + ' ' + ('*' * 34))
            $z[1].Length | Should -Be 90
        }
    }
}
