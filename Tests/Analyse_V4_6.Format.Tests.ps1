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
                            'neu_tab_max6w_fb','neu_text','Pruefbereich','Ausgabe','Puffer_leeren')
        $gefunden = $ast.FindAll({
            param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true) | Where-Object { $zielFunktionen -contains $_.Name }

        foreach ($f in $gefunden) {
            # "function X ..." -> "function global:X ...", damit die Definition
            # die Pester-Scopes ueberlebt und in den It-Bloecken aufrufbar ist.
            $definition = $f.Extent.Text -replace '^function\s+', 'function global:'
            Invoke-Expression $definition
        }

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
                         'neu_tab_max6w_fb','neu_text','Pruefbereich','Ausgabe','Puffer_leeren',
                         'Get-ReportZeilen')) {
            Remove-Item -LiteralPath "function:global:$n" -Force -ErrorAction SilentlyContinue
        }
        foreach ($v in @('sb','zeichen','tabzeichen','tabtrenner','leer','type','maintitel',
                         'firma','madeby','datum','system','F_Rahmen','F_Ue_Schrift','F_Text',
                         'F_Fehler','A_Con','A_Dat','path','A_Puffer')) {
            Remove-Variable -Name $v -Scope Global -Force -ErrorAction SilentlyContinue
        }
    }

    BeforeEach {
        if (Test-Path -LiteralPath $global:path) { Clear-Content -LiteralPath $global:path }
        else { New-Item -ItemType File -Path $global:path | Out-Null }
        [void]$global:A_Puffer.Clear()
    }

    Context 'Skript-Datei (statische Pruefungen)' {

        It 'laesst sich ohne Syntaxfehler parsen' {
            $parseFehler.Count | Should -Be 0
        }

        It 'beginnt mit #Requires -Version' {
            (Get-Content -LiteralPath $skriptPfad -TotalCount 1) |
                Should -Match '^#Requires -Version'
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
