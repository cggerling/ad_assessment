#Requires -Version 5.1
<#
.SYNOPSIS
    Pester-Tests (Pester 5) fuer die Formatierungs-/Layout-Funktionen von AD-Analyse-V5.ps1.

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

Describe 'AD-Analyse-V5.ps1' {

    BeforeAll {
        $skriptPfad = Join-Path (Split-Path -Parent $PSScriptRoot) 'AD-Analyse-V5.ps1'

        ### Skript parsen (ohne Ausfuehrung) ########################################################
        $tokens = $null ; $parseFehler = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $skriptPfad, [ref]$tokens, [ref]$parseFehler)
        if ($parseFehler.Count -gt 0) {
            throw "AD-Analyse-V5.ps1 laesst sich nicht parsen: $($parseFehler[0])"
        }

        ### Formatierungs-Funktionen extrahieren und global definieren #############################
        $zielFunktionen = @('Header','Bottom','Vollzeile','Leerzeile','Trennzeile','tablinie',
                            'Bereich','Phase','Bereichstitel','Subtitel','2werte','new_2werte',
                            'neu_tab_max6w_fb','neu_text','Pruefbereich','Unterpruefung',
                            'Ausgabe','Puffer_leeren','Entschluessle-GPP',
                            'Merken','Doku','Schreibe-Fehler','Farbklasse','HTML_Report','JSON_Export',
                            'Extrahiere-Befunde','chk_delta')
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
        $global:A_Con        = 0                    # kein Voll-Report in der Konsole (Tests)
        $global:A_Prog       = 0                    # keine Fortschrittsanzeige in den Tests
        $global:Fehler_Anzahl = 0                   # Fehlerzaehler (Schreibe-Fehler)
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
            , @(Get-Content -LiteralPath $global:path -Encoding UTF8)
        }
    }

    AfterAll {
        Remove-Item -LiteralPath $global:path -Force -ErrorAction SilentlyContinue
        foreach ($n in @('Header','Bottom','Vollzeile','Leerzeile','Trennzeile','tablinie',
                         'Bereich','Phase','Bereichstitel','Subtitel','2werte','new_2werte',
                         'neu_tab_max6w_fb','neu_text','Pruefbereich','Unterpruefung',
                         'Ausgabe','Puffer_leeren','Entschluessle-GPP',
                         'Merken','Doku','Schreibe-Fehler','Farbklasse','HTML_Report','JSON_Export','Get-ReportZeilen',
                         'Extrahiere-Befunde','chk_delta')) {
            Remove-Item -LiteralPath "function:global:$n" -Force -ErrorAction SilentlyContinue
        }
        foreach ($v in @('sb','zeichen','tabzeichen','tabtrenner','leer','type','maintitel',
                         'firma','madeby','datum','system','F_Rahmen','F_Ue_Schrift','F_Text',
                         'F_Fehler','A_Con','A_Prog','Fehler_Anzahl','A_Dat','path','A_Puffer','A_Htm','A_Jsn',
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

        It 'schreibt den Text-Report als UTF-8 (kein -Encoding ascii, keine Transliteration mehr)' {
            $inhalt = Get-Content -LiteralPath $skriptPfad -Raw
            $inhalt | Should -Not -Match '-Encoding ascii'              # kein ASCII-Report mehr
            $inhalt | Should -Match 'AppendAllText\(\$path'             # Puffer wird per .NET UTF-8 angehaengt
            $inhalt | Should -Match 'WriteAllText\(\$path'              # Report mit UTF-8-BOM angelegt
            $inhalt | Should -Not -Match "-replace 'ü'"                 # neu_text transliteriert nicht mehr
        }

        It 'enthaelt die Modul-Vorabpruefung' {
            $inhalt = Get-Content -LiteralPath $skriptPfad -Raw
            $inhalt | Should -Match 'Get-Module -ListAvailable -Name ActiveDirectory'
            $inhalt | Should -Match 'Get-Module -ListAvailable -Name GroupPolicy'
            $inhalt | Should -Match 'Get-Module -ListAvailable -Name DnsServer'
        }

        It 'srv_chk zaehlt robust (@(...).Count) und castet die Subtraktion ([int])' {
            $srv = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                          $a.Name -eq 'srv_chk'
            }, $true) | Select-Object -First 1
            $srv | Should -Not -BeNullOrEmpty
            $txt = $srv.Extent.Text
            # keine fragile Kleinschreibung ").count" mehr (case-sensitiv, sonst matcht ".Count")
            $txt | Should -Not -CMatch '\)\.count'
            # die beiden Subtraktionen sind int-sicher
            $txt | Should -Match '\[int\]\$2008 - \[int\]\$2008r'
            $txt | Should -Match '\[int\]\$2012 - \[int\]\$2012r'
        }

        It 'verwendet kein Enter-PSSession (laeuft non-interaktiv; Remote-Reads via Invoke-Command)' {
            # Regressionsschutz: Enter-PSSession ist im Skript-Kontext wirkungslos und wuerde
            # Remote-Pruefungen lokal laufen lassen. dcdienste sammelt jetzt via Invoke-Command.
            $aufrufe = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst]
            }, $true) | ForEach-Object { $_.GetCommandName() }
            $aufrufe | Should -Not -Contain 'Enter-PSSession'
            $aufrufe | Should -Not -Contain 'Exit-PSSession'

            $dcd = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                          $a.Name -eq 'dcdienste'
            }, $true) | Select-Object -First 1
            $dcd | Should -Not -BeNullOrEmpty
            $dcd.Extent.Text | Should -Match 'Invoke-Command -ComputerName \$dcho'
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

        It 'buendelt die Datei-Ausgabe (ein UTF-8-Schreibvorgang je Puffer_leeren, kein Add-Content)' {
            $addContent = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst]
            }, $true) | Where-Object { $_.GetCommandName() -eq 'Add-Content' }
            $addContent.Count | Should -Be 0                       # ersetzt durch .NET-UTF-8-Append
            $txt = Get-Content -LiteralPath $skriptPfad -Raw
            ([regex]::Matches($txt, 'AppendAllText\(\$path')).Count | Should -Be 1
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

        It 'behaelt Umlaute im Text-Report (UTF-8, keine Transliteration mehr)' {
            neu_text 0 '-' 'Prüfung' 'Die Lösung wäre über kürzere Wörter möglich.'
            $inhalt = (Get-ReportZeilen) -join ' '
            $inhalt | Should -Match 'Prüfung'
            $inhalt | Should -Match 'Lösung'
            $inhalt | Should -Match 'wäre'
            $inhalt | Should -Match 'über'
            $inhalt | Should -Match 'möglich'
            $inhalt | Should -Not -Match 'ue/ae/oe|Loesung|waere|moeglich'
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

        It 'faengt Fehler ab: kurzer Hinweis im Report, Details ins Fehlerlog, laeuft weiter' {
            $logPfad = [System.IO.Path]::ChangeExtension($global:path, 'Fehler.log')
            Remove-Item $logPfad -Force -ErrorAction SilentlyContinue
            { Pruefbereich 'Testbereich' { throw 'Absichtlicher Testfehler' } } | Should -Not -Throw
            $zeilen = Get-ReportZeilen
            $inhalt = $zeilen -join ' '
            $inhalt | Should -Match 'Bereich nicht vollständig geprüft'   # kurzer Hinweis (Umlaute)
            $inhalt | Should -Match 'Fehlerlog'
            $inhalt | Should -Not -Match 'Absichtlicher Testfehler'        # Detail nur im Log
            foreach ($zeile in $zeilen) { $zeile.Length | Should -Be 90 }
            (Get-Content -LiteralPath $logPfad -Raw -Encoding UTF8) | Should -Match 'Absichtlicher Testfehler'
            Remove-Item $logPfad -Force -ErrorAction SilentlyContinue
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

        It 'jeder Katalog-Eintrag hat einen Hintergrund und HTTPS-Quell-Links mit Titel' {
            foreach ($id in $global:CheckKatalog.Keys) {
                $k = $global:CheckKatalog[$id]
                [string]::IsNullOrWhiteSpace($k.Hintergrund) | Should -BeFalse -Because "$id braucht Hintergrund"
                $k.Quellen -is [string] | Should -BeFalse -Because "$id Quellen muss eine Link-Liste sein"
                @($k.Quellen).Count | Should -BeGreaterThan 0 -Because "$id ohne Quelle"
                foreach ($l in @($k.Quellen)) {
                    $l.Titel | Should -Not -BeNullOrEmpty -Because "$id Quelle ohne Titel"
                    $l.Url   | Should -Match '^https://' -Because "$id Quelle muss HTTPS-Link sein"
                }
            }
        }

        It 'Katalog-Texte verwenden echte Umlaute (nicht transliteriert)' {
            $ae = [char]0x00E4   # ae
            $global:CheckKatalog['domain_allgemein'].Zweck | Should -Match $ae
            $global:CheckKatalog['kerberos'].Titel | Should -Match ('Angriffsfl' + $ae + 'chen')
            # Quellen sind jetzt Link-Listen: Titel des ersten Links pruefen
            @($global:CheckKatalog['domain_allgemein'].Quellen)[0].Titel | Should -Match 'Microsoft'
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
                # CheckId kann $null sein (z. B. Bestands-/Inventory-Teilpruefung ohne Doku).
                $idEl = $cmd.CommandElements[2]
                if ($idEl -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $global:CheckKatalog.Keys | Should -Contain $idEl.Value
                }
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
            $logPfad = [System.IO.Path]::ChangeExtension($global:path, 'Fehler.log')
            Remove-Item $logPfad -Force -ErrorAction SilentlyContinue
            { Unterpruefung 'Testpruefung' 'kerberoasting' { throw 'Simulierter Abfragefehler' } } |
                Should -Not -Throw
            $txt = (Get-ReportZeilen) -join ' '
            $txt | Should -Match 'Teilprüfung übersprungen'        # kurzer Hinweis (Umlaute)
            $txt | Should -Not -Match 'Simulierter Abfragefehler'  # Detail nur im Log
            (Get-Content -LiteralPath $logPfad -Raw -Encoding UTF8) | Should -Match 'Simulierter Abfragefehler'
            Remove-Item $logPfad -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'Paket B (v5.0): Privilegien & ACLs' {

        It 'Katalog enthaelt alle Paket-B-Eintraege' {
            foreach ($id in 'privilegien','dcsync','operatoren','adminsdholder','protected_users','prewin2000') {
                $global:CheckKatalog.Keys | Should -Contain $id
            }
        }

        It 'DCSync-Eintrag ist als Kritisch eingestuft' {
            $global:CheckKatalog['dcsync'].Schwere | Should -Be 'Kritisch'
        }

        It 'die Paket-B-Prueffunktionen sind im Skript definiert' {
            $funktionen = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true) | ForEach-Object { $_.Name }
            foreach ($fn in 'chk_dcsync','chk_operatoren','chk_adminsdholder','chk_protected_users','chk_prewin2000') {
                $funktionen | Should -Contain $fn
            }
        }

        It 'der Schalter privchk steht in der Override-Whitelist' {
            (Get-Content -LiteralPath $skriptPfad -Raw) | Should -Match "'privchk'"
        }
    }

    Context 'Paket C (v5.0): AD CS / ESC' {

        It 'Katalog enthaelt alle Paket-C-Eintraege' {
            foreach ($id in 'adcs','esc1','esc2_3','esc4','esc6','esc8') {
                $global:CheckKatalog.Keys | Should -Contain $id
            }
        }

        It 'ESC1-Eintrag ist als Kritisch eingestuft' {
            $global:CheckKatalog['esc1'].Schwere | Should -Be 'Kritisch'
        }

        It 'die Paket-C-Prueffunktionen sind im Skript definiert' {
            $funktionen = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true) | ForEach-Object { $_.Name }
            foreach ($fn in 'Get-ADCSObjekte','Ist-NiedrigPriv','Get-NiedrigPrivEnroller',
                            'chk_esc1','chk_esc2_3','chk_esc4','chk_esc6','chk_esc8') {
                $funktionen | Should -Contain $fn
            }
        }

        It 'der Schalter adcschk steht in der Override-Whitelist' {
            (Get-Content -LiteralPath $skriptPfad -Raw) | Should -Match "'adcschk'"
        }
    }

    Context 'Paket D (v5.0): GPO/SYSVOL-Geheimnisse' {

        It 'Katalog enthaelt alle Paket-D-Eintraege' {
            foreach ($id in 'gpo_sysvol','gpp_cpassword','sysvol_scripts','gpo_rights') {
                $global:CheckKatalog.Keys | Should -Contain $id
            }
        }

        It 'GPP-cpassword-Eintrag ist als Kritisch eingestuft' {
            $global:CheckKatalog['gpp_cpassword'].Schwere | Should -Be 'Kritisch'
        }

        It 'die Paket-D-Prueffunktionen sind im Skript definiert' {
            $funktionen = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true) | ForEach-Object { $_.Name }
            foreach ($fn in 'Entschluessle-GPP','chk_gpp_cpassword','chk_sysvol_scripts','chk_gpo_rights') {
                $funktionen | Should -Contain $fn
            }
        }

        It 'der Schalter sysvchk steht in der Override-Whitelist' {
            (Get-Content -LiteralPath $skriptPfad -Raw) | Should -Match "'sysvchk'"
        }

        It 'Entschluessle-GPP entschluesselt einen mit dem GPP-Schluessel kodierten Wert korrekt' {
            $klar = 'GeheimesP@ss1'
            $key = [byte[]](0x4e,0x99,0x06,0xe8,0xfc,0xb6,0x6c,0xc9,0xfa,0xf4,0x93,0x10,0x62,0x0f,0xfe,0xe8,
                            0xf4,0x96,0xe8,0x06,0xcc,0x05,0x79,0x90,0x20,0x9b,0x09,0xa4,0x33,0xb6,0x6c,0x1b)
            $aes = New-Object System.Security.Cryptography.AesManaged
            $aes.Key = $key ; $aes.IV = New-Object byte[] 16 ; $aes.Mode = 'CBC' ; $aes.Padding = 'PKCS7'
            $enc = $aes.CreateEncryptor()
            $pt  = [System.Text.Encoding]::Unicode.GetBytes($klar)
            $ct  = $enc.TransformFinalBlock($pt, 0, $pt.Length)
            $cpassword = [Convert]::ToBase64String($ct)
            (Entschluessle-GPP $cpassword) | Should -Be $klar
        }
    }

    Context 'Paket E (v5.0): DC-Härtung vertieft' {

        It 'Katalog enthaelt alle Paket-E-Eintraege' {
            foreach ($id in 'dc_haertung','ldap_signing','smb_signing','print_spooler','anon_ldap') {
                $global:CheckKatalog.Keys | Should -Contain $id
            }
        }

        It 'die Paket-E-Prueffunktionen sind im Skript definiert' {
            $funktionen = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true) | ForEach-Object { $_.Name }
            foreach ($fn in 'chk_ldap_signing','chk_smb_signing','chk_print_spooler','chk_anon_ldap') {
                $funktionen | Should -Contain $fn
            }
        }

        It 'der Schalter dchaert steht in der Override-Whitelist' {
            (Get-Content -LiteralPath $skriptPfad -Raw) | Should -Match "'dchaert'"
        }
    }

    Context 'Paket F (v5.0): Delta-Modus' {

        It 'Katalog enthaelt den Delta-Eintrag' {
            $global:CheckKatalog.Keys | Should -Contain 'delta'
        }

        It 'der Delta-Eintrag ist als Info eingestuft' {
            $global:CheckKatalog['delta'].Schwere | Should -Be 'Info'
        }

        It 'die Delta-Funktionen sind im Skript definiert' {
            $funktionen = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true) | ForEach-Object { $_.Name }
            foreach ($fn in 'Extrahiere-Befunde','chk_delta') {
                $funktionen | Should -Contain $fn
            }
        }

        It 'der Schalter deltchk steht in der Override-Whitelist' {
            (Get-Content -LiteralPath $skriptPfad -Raw) | Should -Match "'deltchk'"
        }

        It 'der Parameter -Vergleich ist deklariert' {
            (Get-Content -LiteralPath $skriptPfad -Raw) | Should -Match '\[string\]\$Vergleich'
        }

        It 'Extrahiere-Befunde erfasst nur rot/gelb markierte Wert-Eintraege' {
            $evs = @(
                [pscustomobject]@{ Art='Wert';    Name='R'; Wert='1'; Farbe='Red' }
                [pscustomobject]@{ Art='Wert';    Name='Y'; Wert='2'; Farbe='Yellow' }
                [pscustomobject]@{ Art='Wert';    Name='G'; Wert='3'; Farbe='Green' }
                [pscustomobject]@{ Art='Bereich'; Name='B'; Wert='';  Farbe='Red' }
            )
            $set = Extrahiere-Befunde $evs
            $set.Count | Should -Be 2
            $set.ContainsKey('R | 1') | Should -BeTrue
            $set.ContainsKey('Y | 2') | Should -BeTrue
            $set.ContainsKey('G | 3') | Should -BeFalse
        }

        It 'Extrahiere-Befunde ignoriert eigene Delta-Artefakte frueherer Laeufe' {
            $evs = @(
                [pscustomobject]@{ Art='Wert'; Name='  + Alt-Befund';      Wert=''; Farbe='Red' }
                [pscustomobject]@{ Art='Wert'; Name='Neu hinzugekommen:';  Wert='3'; Farbe='Red' }
                [pscustomobject]@{ Art='Wert'; Name='Echter Befund';       Wert='x'; Farbe='Red' }
            )
            $set = Extrahiere-Befunde $evs
            $set.Count | Should -Be 1
            $set.ContainsKey('Echter Befund | x') | Should -BeTrue
        }

        It 'chk_delta erkennt neue und behobene Befunde gegenueber einem alten JSON-Export' {
            # Aktueller Lauf: Befund A (bleibt) + Befund NEU; Gruen ist kein Befund.
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Wert'; Name='Befund A';   Wert='x';  Farbe='Red' })
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Wert'; Name='Befund NEU'; Wert='y';  Farbe='Red' })
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Wert'; Name='Gruen';      Wert='ok'; Farbe='Green' })

            # Alter Export: Befund A (bleibt) + Befund ALT (behoben).
            $tmp = Join-Path ([IO.Path]::GetTempPath()) 'ad_delta_old.json'
            $altObj = [pscustomobject]@{ Ereignisse = @(
                [pscustomobject]@{ Art='Wert'; Name='Befund A';   Wert='x'; Farbe='Red' }
                [pscustomobject]@{ Art='Wert'; Name='Befund ALT'; Wert='z'; Farbe='Yellow' }
            )}
            $altObj | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $tmp -Encoding UTF8

            try {
                chk_delta $tmp

                $neuEv = $global:R_Daten | Where-Object { $_.Name -eq 'Neu hinzugekommen:' }
                $neuEv.Wert | Should -Be '1'
                ($global:R_Daten | Where-Object { "$($_.Name)" -match '\+\s*Befund NEU' }) |
                    Should -Not -BeNullOrEmpty

                $behEv = $global:R_Daten | Where-Object { $_.Name -eq 'Behoben (nicht mehr vorhanden):' }
                $behEv.Wert | Should -Be '1'
                ($global:R_Daten | Where-Object { "$($_.Name)" -match '\-\s*Befund ALT' }) |
                    Should -Not -BeNullOrEmpty
            } finally {
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            }
        }

        It 'chk_delta meldet eine fehlende Vergleichsdatei freundlich' {
            $fehlt = Join-Path ([IO.Path]::GetTempPath()) 'gibt_es_nicht_12345.json'
            chk_delta $fehlt
            ($global:R_Daten | Where-Object {
                $_.Name -eq 'Vergleichsdatei:' -and "$($_.Wert)" -match 'nicht gefunden'
            }) | Should -Not -BeNullOrEmpty
        }

        It 'keine CheckId wird von Pruefbereich UND Unterpruefung geteilt (kein doppelter Doku-Block)' {
            # Regressionsschutz: teilen sich Bereich und Teilpruefung dieselbe CheckId, erscheint
            # die Begruendung (Doku) doppelt - in Exec-Summary und Bericht (war beim Delta-Bereich so).
            $cmds = $ast.FindAll({
                param($a) $a -is [System.Management.Automation.Language.CommandAst]
            }, $true)
            $bereichIds = @() ; $unterIds = @()
            foreach ($c in $cmds) {
                switch ($c.GetCommandName()) {
                    'Pruefbereich' {
                        for ($i = 0; $i -lt $c.CommandElements.Count - 1; $i++) {
                            $el = $c.CommandElements[$i]
                            if ($el -is [System.Management.Automation.Language.CommandParameterAst] -and
                                $el.ParameterName -eq 'CheckId') {
                                $v = $c.CommandElements[$i + 1]
                                if ($v -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                                    $bereichIds += $v.Value
                                }
                            }
                        }
                    }
                    'Unterpruefung' {
                        $v = $c.CommandElements[2]   # Unterpruefung <titel> <checkid> <block>
                        if ($v -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                            $unterIds += $v.Value
                        }
                    }
                }
            }
            $geteilt = @($bereichIds | Where-Object { $unterIds -contains $_ })
            $geteilt | Should -BeNullOrEmpty
        }
    }

    Context 'Zwei-Phasen-Struktur (Recon / Schwachstellen)' {

        It 'Hauptablauf trennt Phase 1 (Recon) und Phase 2 (Schwachstellen) sauber' {
            $inhalt = Get-Content -LiteralPath $skriptPfad -Raw
            $p1 = $inhalt.IndexOf('Phase "PHASE 1')
            $p2 = $inhalt.IndexOf('Phase "PHASE 2')
            $p1 | Should -BeGreaterThan 0
            $p2 | Should -BeGreaterThan $p1
            # Reconnaissance-Bereiche stehen zwischen Phase-1- und Phase-2-Banner
            foreach ($id in 'domain_allgemein','admins','benutzer','ad_gruppen','ca') {
                $pos = $inhalt.IndexOf("-CheckId '$id'")
                $pos | Should -BeGreaterThan $p1
                $pos | Should -BeLessThan $p2
            }
            # Schwachstellen-/Posture-Bereiche stehen nach dem Phase-2-Banner
            foreach ($id in 'ddp_password','user_vs_pw','logging','kerberos','privilegien','adcs','gpo_sysvol','dc_detail','dc_haertung') {
                $inhalt.IndexOf("-CheckId '$id'") | Should -BeGreaterThan $p2
            }
        }

        It 'Phase erzeugt ein Phase-Ereignis und einen ===-Banner im Report' {
            Phase 'TESTPHASE'
            ($global:R_Daten | Where-Object { $_.Art -eq 'Phase' -and $_.Titel -eq 'TESTPHASE' }) |
                Should -Not -BeNullOrEmpty
            $z = Get-ReportZeilen
            ($z -join "`n") | Should -Match 'TESTPHASE'
            (@($z) | Where-Object { $_ -match '^={10,}$' }).Count | Should -BeGreaterOrEqual 2
        }

        It 'HTML_Report rendert Phasen-Ueberschriften und gruppiert die Zusammenfassung' {
            $global:R_Daten.Clear()
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Phase'; Titel='PHASE 1 - RECON' })
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Bereich'; Titel='B-eins' })
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Doku'; CheckId='c1'; DTitel='Check Eins'; Schwere='Info'; Zweck='z'; Hintergrund=''; Beispiel='b'; Empfehlung='e'; Quellen='q' })
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Phase'; Titel='PHASE 2 - SCHWACH' })
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Bereich'; Titel='B-zwei' })
            [void]$global:R_Daten.Add([pscustomobject]@{ Art='Doku'; CheckId='c2'; DTitel='Check Zwei'; Schwere='Hoch'; Zweck='z'; Hintergrund=''; Beispiel='b'; Empfehlung='e'; Quellen='q' })
            HTML_Report
            $htmlPfad = [System.IO.Path]::ChangeExtension($global:path, 'html')
            $html = Get-Content -LiteralPath $htmlPfad -Raw -Encoding UTF8
            $html | Should -Match '<h2 class="phase">PHASE 1 - RECON</h2>'
            $html | Should -Match '<h2 class="phase">PHASE 2 - SCHWACH</h2>'
            $html | Should -Match '<h3>PHASE 1 - RECON</h3>'
            $html | Should -Match '<h3>PHASE 2 - SCHWACH</h3>'
            Remove-Item $htmlPfad -Force -ErrorAction SilentlyContinue
        }
    }

    Context 'Konsole-Fortschritt und Fehlerlog' {

        It 'Konsole zeigt standardmaessig nur Fortschritt (A_Con=0, A_Prog=1); -KeineKonsole schaltet beides ab' {
            $inhalt = Get-Content -LiteralPath $skriptPfad -Raw
            $inhalt | Should -Match '(?m)^\$A_Con = 0'
            $inhalt | Should -Match '(?m)^\$A_Prog = 1'
            $inhalt | Should -Match 'KeineKonsole\) \{ \$A_Con = 0 ; \$A_Prog = 0'
        }

        It 'Schreibe-Fehler protokolliert mit Zeitstempel ins separate Fehlerlog' {
            $logPfad = [System.IO.Path]::ChangeExtension($global:path, 'Fehler.log')
            Remove-Item $logPfad -Force -ErrorAction SilentlyContinue
            Schreibe-Fehler 'Bereich: Test' 'etwas ist schiefgelaufen'
            Test-Path $logPfad | Should -BeTrue
            $log = Get-Content -LiteralPath $logPfad -Raw -Encoding UTF8
            $log | Should -Match 'Fehlerprotokoll'
            $log | Should -Match '\[Bereich: Test\] etwas ist schiefgelaufen'
            $log | Should -Match '\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]'   # Zeitstempel
            Remove-Item $logPfad -Force -ErrorAction SilentlyContinue
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
            $logPfad = [System.IO.Path]::ChangeExtension($global:path, 'Fehler.log')
            Pruefbereich 'Testbereich' { throw 'Absichtlicher Testfehler' }
            # Direkt lesen, ohne Get-ReportZeilen (das wuerde selbst flushen):
            $inhalt = @(Get-Content -LiteralPath $global:path -Encoding UTF8) -join ' '
            $inhalt | Should -Match 'Fehlerlog'           # kurzer Hinweis steht im Report (Puffer geflusht)
            Remove-Item $logPfad -Force -ErrorAction SilentlyContinue
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
