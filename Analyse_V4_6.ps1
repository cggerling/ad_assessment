#Requires -Version 5.1

<#
.SYNOPSIS
    AD-Analyse / Assessment-Report fuer eine On-Premises Active-Directory-Umgebung.

.DESCRIPTION
    Read-only Sammel-Skript, das eine Active-Directory-Umgebung auswertet und einen
    fest formatierten Text-Report erzeugt. Die einzelnen Pruefbereiche lassen sich ueber
    Schalter-Variablen im Kopf des Skripts (Wert 0/1/2) aktivieren bzw. deaktivieren.

    Abgedeckte Bereiche (Auswahl): Domain/Forest/FSMO, Schema, Central Store & Security
    Templates, Domain Controller, Logging/Audit, AD-Trusts, DNS, SYSVOL/DFSR-Replikation,
    administrative Gruppen, Benutzer- und Computerkonten, MSA/gMSA, Passwort-Policies
    (Default & Fine-Grained), AD-Gruppen, GPOs, Organisationseinheiten, Zertifizierungs-
    stellen sowie DC-Detailpruefungen (Dienste, Rollen, Features, LDAPS, NTLM, SMB1,
    BitLocker, ExecutionPolicy).

.PARAMETER Verzeichnis
    Ausgabeverzeichnis fuer den Report (Standard: c:\AD-Assessment).

.PARAMETER Breite
    Breite der Report-Ausgabe in Zeichen, 70 bis 90 (Standard: 90).

.PARAMETER KeineKonsole
    Unterdrueckt die farbige Konsolenausgabe (nur Datei-Ausgabe).

.PARAMETER KeineDatei
    Unterdrueckt die Datei-Ausgabe (nur Konsolenausgabe).

.PARAMETER Bereiche
    Hashtable, um einzelne Pruefbereich-Schalter zu ueberschreiben, ohne das Skript zu
    editieren. Schluessel = Schaltername aus dem Skript-Kopf, Wert = 0/1/2 (Semantik wie
    im Kopf dokumentiert). Beispiel: -Bereiche @{ dnschk = 0; DomCon = 2 }
    Unbekannte Schalternamen werden mit Warnung ignoriert.

.PARAMETER KeinHTML
    Unterdrueckt den zusaetzlichen HTML-Report (Standard: wird erzeugt).

.PARAMETER KeinJSON
    Unterdrueckt den zusaetzlichen JSON-Export (Standard: wird erzeugt).

.PARAMETER Vergleich
    Pfad zu einem frueheren JSON-Export. Ist er angegeben, ergaenzt der Report einen
    Delta-Bereich, der neu hinzugekommene und behobene Befunde (rot/gelb markierte
    Eintraege) gegenueber jenem Lauf auflistet.

.EXAMPLE
    .\Analyse_V4_6.ps1
    Kompletter Lauf mit den im Skript-Kopf hinterlegten Standardwerten.

.EXAMPLE
    .\Analyse_V4_6.ps1 -Verzeichnis "D:\Reports" -KeineKonsole
    Report nach D:\Reports schreiben, ohne Konsolenausgabe.

.EXAMPLE
    .\Analyse_V4_6.ps1 -Bereiche @{ dnschk = 0; allgpo = 0 }
    Kompletter Lauf, aber ohne DNS- und GPO-Pruefung.

.EXAMPLE
    .\Analyse_V4_6.ps1 -Vergleich "C:\AD-Assessment\DC01\AD-Analyse-Report_alt.json"
    Kompletter Lauf mit Delta-Bereich: neue/behobene Befunde gegenueber dem alten Export.

.OUTPUTS
    Fest formatierter Text-Report unter $verz\<COMPUTERNAME>\<B_Datei>_<Datum>.txt
    (Standard-Ausgabeverzeichnis: c:\AD-Assessment). Optional zusaetzlich Konsolenausgabe.
    Zusaetzlich (abschaltbar): HTML-Report (.html, eigenstaendig/offline-faehig) und
    strukturierter JSON-Export (.json) mit denselben Basisnamen.

.NOTES
    Version       : 5.0
    Ausfuehrung    : Auf einem Domain Controller oder einem System mit installierten
                     RSAT-Modulen, im Kontext eines Kontos mit Leserechten in der Domaene
                     (Domain-Admin-aequivalent fuer einzelne Detailpruefungen empfohlen).
    Abhaengigkeiten: PowerShell-Module ActiveDirectory, GroupPolicy, DnsServer;
                     WinRM/Invoke-Command fuer Remote-Pruefungen der Domain Controller.
    Charakter      : Ausschliesslich lesend (keine schreibenden AD-Operationen).
#>

####################################################################################################
# Parameter (optional - ohne Angabe gelten die Standardwerte aus dem Skript-Kopf)                  #
####################################################################################################
[CmdletBinding()]
param (
    [string]$Verzeichnis,                            # Ausgabeverzeichnis (Standard im Kopf: $verz)
    [ValidateRange(70,90)]
    [int]$Breite,                                    # Ausgabebreite (Standard im Kopf: $sb)
    [switch]$KeineKonsole,                           # Konsolenausgabe unterdruecken ($A_Con = 0)
    [switch]$KeineDatei,                             # Datei-Ausgabe unterdruecken   ($A_Dat = 0)
    [switch]$KeinHTML,                               # HTML-Report unterdruecken     ($A_Htm = 0)
    [switch]$KeinJSON,                               # JSON-Export unterdruecken     ($A_Jsn = 0)
    [string]$Vergleich,                              # Delta: Pfad zu frueherem JSON-Export
    [hashtable]$Bereiche                             # Schalter-Overrides, z.B. @{ dnschk = 0 }
)
####################################################################################################
### AD-Analyse Script                                                                Version 5.0 ###
####################################################################################################
# Globale Definitionen:                                                                            #
##################################################                                                 #
$version = "Vers. 5.0"                           # Script Version                                  #
$company = "AD-Assessment"                       # Firma (neutraler Platzhalter)                   #
$firma = "$version " + "$company"                # Header Rechts unten                             #
$sb = 90                                         # Breite der Ausgabe (min.70 - max.90)            #
$type = "Report"                                 # Bereich (Anlage/Report/Check/etc..)             #
$maintitel = "AD-Analyse"                        # Gewuenschter Titel                              #
$B_Datei = "AD-Analyse-Report"                   # Erster Teil des Ausgabedateinamens              #
$zeichen = "*"                                   # Gewaehltes Zeichen fuer den Rahmen              #
$tabzeichen = "-"                                # Gewaehltes Zeichen fuer Tabellen Querzeilen     #
$tabtrenner = "|"                                # Gewaehltes Zeichen fuer Spaltentrenner          #
$leer = " "                                      # Gewsehltes Zeichen fuers Leerzeichen            #
$madeby = "AD-Assessment Tool"                   # Ersteller des Scripts                           #
$F_Rahmen = "DarkYellow"                         # Farbe fuer die Rahmenzeichen                    #
$F_Ue_Schrift = "Gray"                           # Farbe fuer Ueberschriften                       #
$F_Text = "White"                                # Farbe fuer normalen Text                        #
$F_Fehler = "Red"                                # Farbe fuer Fehlermeldungen                      #
$A_Dat = 1                                       # Ausgabe auch in Datei? 1=Ja 0=Nein              #
$A_Con = 1                                       # Ausgabe auch in Konsole? 1=Ja 0=Nein            #
$A_Htm = 1                                       # Zusaetzlich HTML-Report? 1=Ja 0=Nein            #
$A_Jsn = 1                                       # Zusaetzlich JSON-Export? 1=Ja 0=Nein            #
####################################################################################################
# Moegliche Farben:                                                                                #
###################                                                                                #
# Black, Blue, Cyan, Gray, Green, Magenta, Red, White, Yellow                                      #
# DarkBlue, DarkCyan, DarkGray, DarkGreen, DarkMagenta, DarkRed, DarkYellow                        #
####################################################################################################
# Check-Bereiche aus-/einschalten                                                                  #
####################################################################################################
# Domain Allgemein #                             ###################################################
$domoco = 1                                      # Domain/Mode/DC/FSMO               (0=nein,1=ja) #
$schema = 1                                      # Attribute zu Schema Erweiterungen (0=nein,1=ja) #
# Central Store & Templates                      ###################################################
$censto = 1                                      # Central Store prüfen              (0=nein,1=ja) #
$sectem = 1                                      # Sec. Templates prüfen             (0=nein,1=ja) #
# Domain Controller #                            ###################################################
$domdcs = 1                                      # Domain Controller                 (0=nein,1=ja) #
# Logging auf DCs #                              ###################################################
$loggin = 1                                      # Logging Check                     (0=nein,1=ja) #
# AD Trusts #                                    ###################################################
$adtchk = 1                                      # AD-Trusts Check                   (0=nein,1=ja) #
# DNS Check #                                    ###################################################
$dnschk = 1                                      # DNS Pruefung                      (0=nein,1=ja) #
# Sysvol Replication & AD-Health                 ###################################################
$SysRep = 1                                      # AD-Health Check                   (0=nein,1=ja) #
# Administratoren #                              ###################################################
$admusr = 1                                      # Admin Gruppen Check        (prim) (0=nein,1=ja) #
$lokadm = 1                                      # Lokale Adm Gruppen Check          (0=nein,1=ja) #
$AdmGri = 1                                      # Erweiterte Dom.Adm Gruppen Check  (0=nein,1=ja) #
$buildi = 1                                      # Builtin Benutzer            (sub) (0=nein,1=ja) #
$priusr = 1                                      # AdminCount 1                (sub) (0=nein,1=ja) #
# Benutzer Accounts #                            ###################################################
$usrchk = 1                                      # Benutzer Check             (prim) (0=nein,1=ja) #
$inachk = 1                                      # inaktive Benutzer           (sub) (0=nein,1=ja) #
$geschk = 1                                      # gesperrte Benutzer          (sub) (0=nein,1=ja) #
$falchk = 1                                      # User in OU Users            (sub) (0=nein,1=ja) #
# Computer Accounts #                            ###################################################
$syschk = 1                                      # System Check                      (0=nein,1=ja) #
$cltchk = 1                                      # Client Check     (0=nein,1=ja,2=alle+Supported) #
$srvchk = 1                                      # Server Check     (0=nein,1=ja,2=alle+Supported) #
$no_win = 1                                      # Nicht Windows Check               (0=nein,1=ja) #
# Group Managed Service Accounts #               ###################################################
$manacc = 1                                      # MSA/gMSA Check                    (0=nein,1=ja) #
# Alles zu Passwoertern #                        ###################################################
$dDPchk = 1                                      # dDP und User Check                (0=nein,1=ja) #
$fgppch = 1                                      # fGPP Check                        (0=nein,1=ja) #
$userpw = 1                                      # User vs Password Policies Check   (0=nein,1=ja) #
# AD-Gruppen und GPOs #                          ###################################################
$allgru = 1                                      # AD-Gruppen Check                  (0=nein,1=ja) #
$allgpo = 1                                      # GPO Check                         (0=nein,1=ja) #
# Organisation Units #                           ###################################################
$OrgUni = 1                                      # OU Check aus/ein          (0=nein,1=ja,2=Pfade) #
#$aclchk = 0                                     # Rechtedelegierung                 (0=nein,1=ja) #
# Zertifizierungsstellen #                       ###################################################
$caschk = 1                                      # CA Check aus/ein                  (0=nein,1=ja) #
# Check der einzelnen Domain Controller #        ###################################################
$DomCon = 1                                      # DC Check aus/ein     (0=nein,1=Teil,2=komplett) #
# Sicherheit: Kerberos-Angriffsflaechen (v5.0)   ###################################################
$kerbchk = 1                                     # Kerberos-Checks (Paket A)         (0=nein,1=ja) #
# Sicherheit: Privilegien & ACLs (v5.0)          ###################################################
$privchk = 1                                     # Privilegien-/ACL-Checks (Paket B) (0=nein,1=ja) #
# Sicherheit: AD CS / ESC (v5.0)                 ###################################################
$adcschk = 1                                     # AD-CS-/ESC-Checks (Paket C)       (0=nein,1=ja) #
# Sicherheit: GPO/SYSVOL-Geheimnisse (v5.0)      ###################################################
$sysvchk = 1                                     # GPO/SYSVOL-Checks (Paket D)       (0=nein,1=ja) #
# Sicherheit: DC-Haertung vertieft (v5.0)        ###################################################
$dchaert = 1                                     # DC-Haertung-Checks (Paket E)      (0=nein,1=ja) #
# Sicherheit: Delta-Modus / Zeitvergleich (v5.0) ###################################################
$deltchk = 1                                     # Delta-Bereich (Paket F)           (0=nein,1=ja) #
####################################################################################################
# Parameter-Overrides anwenden (nur wenn beim Aufruf angegeben):                                   #
################################################################                                   #
if ($PSBoundParameters.ContainsKey('Breite')) { $sb = $Breite }                                    #
if ($KeineKonsole) { $A_Con = 0 }                # Konsolenausgabe per Parameter aus               #
if ($KeineDatei)   { $A_Dat = 0 }                # Datei-Ausgabe per Parameter aus                 #
if ($KeinHTML)     { $A_Htm = 0 }                # HTML-Report per Parameter aus                   #
if ($KeinJSON)     { $A_Jsn = 0 }                # JSON-Export per Parameter aus                   #
if ($Bereiche) {                                 # Einzelne Schalter per Hashtable ueberschreiben  #
    $schalterListe = @('domoco','schema','censto','sectem','domdcs','loggin','adtchk','dnschk',    #
                       'SysRep','admusr','lokadm','AdmGri','buildi','priusr','usrchk','inachk',    #
                       'geschk','falchk','syschk','cltchk','srvchk','no_win','manacc','dDPchk',    #
                       'fgppch','userpw','allgru','allgpo','OrgUni','caschk','DomCon','kerbchk',   #
                       'privchk','adcschk','sysvchk','dchaert','deltchk')                          #
    foreach ($schalter in $Bereiche.Keys) {                                                        #
        if ($schalterListe -contains $schalter) {                                                  #
            Set-Variable -Name $schalter -Value ([int]$Bereiche[$schalter])                        #
        } else {                                                                                   #
            Write-Host "WARNUNG: Unbekannter Bereichs-Schalter '$schalter' wird ignoriert." -ForegroundColor $F_Fehler
        }                                                                                          #
    }                                                                                              #
}                                                                                                  #
####################################################################################################
# Globale Variablen:                                                                               #
##################################################                                                 #
[string]$system = $ENV:COMPUTERNAME              # Systemnamen auslesen und bereitstellen          #
$datum = Get-Date -Format "dd-MM-yyyy HH:mm"     # Datum auslesen und bereitstellen Funktion       #
$datu = Get-Date -Format "dd-MM-yyyy HH.mm.ss"   # Datum auslesen und bereitstellen Dateinamen     #
#$datei = "$type"+".txt"                         # Dateinamen zusammenbauen Type ohne Datum-Zeit   #
$datei = "$B_Datei"+"_"+"$datu.txt"              # Dateinamen zusammenbauen Type mit Datum-Zeit    #
$verz = "c:\AD-Assessment"                       # Wo soll die Datei abgelegt werden               #
if ($PSBoundParameters.ContainsKey('Verzeichnis')) { $verz = $Verzeichnis }                        #
$sysverz = "$verz\$system"                       # Verzeichnis fuer die Ausgabedatei               #
$pathtemp = "$sysverz\$datei"                    # Path und Datei zusammenbauen                    #
$path = $pathtemp                                # Zum vermeiden von Convertierungsfehler umleiten #
$A_Puffer = New-Object System.Text.StringBuilder # Sammel-Puffer fuer die Datei-Ausgabe            #
$R_Daten = New-Object 'System.Collections.Generic.List[object]' # Ereignisliste fuer HTML/JSON     #
####################################################################################################
# Modul-Vorabpruefung: Abhaengigkeiten frueh und klar melden (vor Anlage der Ausgabedatei)         #
##########################################################################################         #
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {                                      #
    Write-Host "FEHLER : PowerShell-Modul 'ActiveDirectory' nicht gefunden." -ForegroundColor $F_Fehler
    Write-Host "         Bitte RSAT-AD-PowerShell installieren. Das Skript wird beendet." -ForegroundColor $F_Fehler
    exit 1                                       # Ohne AD-Modul ist keine Pruefung moeglich       #
}                                                                                                  #
if ($allgpo -ge 1 -and -not (Get-Module -ListAvailable -Name GroupPolicy)) {                       #
    Write-Host "WARNUNG: PowerShell-Modul 'GroupPolicy' nicht gefunden." -ForegroundColor $F_Fehler
    Write-Host "         Der GPO-Check wird übersprungen (allgpo=0)." -ForegroundColor $F_Fehler
    $allgpo = 0                                  # Bereich deaktivieren statt Laufzeitfehler       #
}                                                                                                  #
if ($dnschk -ge 1 -and -not (Get-Module -ListAvailable -Name DnsServer)) {                         #
    Write-Host "WARNUNG: PowerShell-Modul 'DnsServer' nicht gefunden." -ForegroundColor $F_Fehler
    Write-Host "         Die DNS-Prüfung wird übersprungen (dnschk=0)." -ForegroundColor $F_Fehler
    $dnschk = 0                                  # Bereich deaktivieren statt Laufzeitfehler       #
}                                                                                                  #
####################################################################################################
# Verzeichnis und Datei fuer die Ausgabe pruefen bzw. anlegen:                                     #
##############################################################                                     #
if ($A_Dat -eq 1) {                                 # Pruefung ab eine Ausgabe in Datei gewuenscht #
    if (!(Test-Path $sysverz)) {                    # Pruefung ob Verzeichnis vorhanden            #
        New-Item -Path $sysverz -ItemType Directory # Verzeichnis wird erstellt                    #
    }                                               #                                              #
    # Leere Datei mit UTF-8-BOM anlegen, damit Umlaute korrekt sind                                #
    # (PS 5.1 liest BOM-loses UTF-8 sonst als ANSI -> Umlaut-Salat):                                #
    [System.IO.File]::WriteAllText($path, '', (New-Object System.Text.UTF8Encoding($true)))         #
}                                                                                                  #
####################################################################################################
# Gepufferte Datei-Ausgabe (Performance: ein Schreibvorgang je Bereich statt je Zeile)             #
####################################################################################################
function Ausgabe ([string]$zeile) {                                                                #
    [void]$A_Puffer.AppendLine($zeile)           # Zeile in den Puffer statt direkt in die Datei   #
}                                                                                                  #
function Puffer_leeren {                                                                           #
    if ($A_Dat -eq 1 -and $A_Puffer.Length -gt 0) {                                                #
        # UTF-8 anhaengen ohne weiteres BOM (das BOM steht bereits am Dateianfang):                 #
        [System.IO.File]::AppendAllText($path, $A_Puffer.ToString(), (New-Object System.Text.UTF8Encoding($false))) #
        [void]$A_Puffer.Clear()                  # Puffer nach dem Schreiben leeren                #
    }                                                                                              #
}                                                                                                  #
####################################################################################################
# Strukturierte Erfassung: Report-Ereignisse fuer HTML-Report und JSON-Export sammeln             #
####################################################################################################
function Merken ($art, [hashtable]$daten) {                                                        #
    # $art = Ereignistyp (Kopf/Bereich/Titel/Subtitel/Wert/TabZeile/Text/Doku)                     #
    # $daten = Inhalt des Ereignisses; wird nur gesammelt wenn HTML oder JSON aktiv ist            #
    if ($A_Htm -eq 1 -or $A_Jsn -eq 1) {                                                           #
        $daten['Art'] = $art                                                                       #
        [void]$R_Daten.Add([pscustomobject]$daten)                                                 #
    }                                                                                              #
}                                                                                                  #
####################################################################################################
# Check-Katalog: fundierte Begruendung je Pruefbereich (Zweck/Beispiel/Empfehlung/Quellen)         #
# Wird im HTML-Report als einklappbarer Hintergrund-Block und im JSON-Export ausgegeben.           #
# Schwere: Info | Niedrig | Mittel | Hoch | Kritisch                                               #
####################################################################################################
$CheckKatalog = @{
    'domain_allgemein' = @{
        Titel = 'Domain, Mode, FSMO'; Schwere = 'Info'
        Zweck = 'Erfasst Grunddaten von Domäne und Forest: Funktionsebenen, FSMO-Rollenverteilung, AD-Papierkorb und das Alter des krbtgt-Kontos. Diese Werte bilden die Basis für alle weiteren Bewertungen.'
        Beispiel = 'Eine niedrige Funktionsebene (z. B. Windows Server 2008) verhindert moderne Sicherheitsfunktionen wie AD-Papierkorb oder gMSA. Ein sehr altes krbtgt-Passwort erleichtert Golden-Ticket-Angriffe.'
        Empfehlung = 'Funktionsebenen auf den höchsten von allen DCs unterstützten Stand heben; AD-Papierkorb aktivieren; krbtgt-Passwort regelmäßig (zweifach im Abstand) zurücksetzen.'
        Hintergrund = 'Die Funktionsebene bestimmt, welche AD-Sicherheitsfunktionen überhaupt nutzbar sind. Das krbtgt-Konto ist der zentrale Generalschlüssel der Domäne: Wer es erbeutet, kann sich dauerhaft beliebige Identitäten ausstellen ("Golden Ticket") - ein altes krbtgt-Passwort erhöht dieses Risiko. Der AD-Papierkorb erlaubt es, versehentlich gelöschte Objekte wiederherzustellen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Active Directory Domain Services Functional Levels'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/active-directory-functional-levels' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'central_store' = @{
        Titel = 'Central Store & Templates'; Schwere = 'Niedrig'
        Zweck = 'Prüft, ob ein zentraler Speicher (Central Store) für GPO-Vorlagen (ADMX/ADML) im SYSVOL existiert und ob Security-Templates vorhanden sind. Der Central Store hält die GPO-Vorlagen auf allen DCs einheitlich.'
        Beispiel = 'Ohne Central Store ziehen Administratoren die ADMX-Dateien vom lokalen Rechner - je nach Patchstand fehlen dann Richtlinien-Einstellungen oder sind uneinheitlich.'
        Empfehlung = 'Central Store unter \\<Domäne>\SYSVOL\<Domäne>\Policies\PolicyDefinitions anlegen und aktuell halten.'
        Hintergrund = 'Gruppenrichtlinien-Vorlagen liegen als sprachneutrale .admx- und sprachspezifische .adml-Dateien vor. Ohne Central Store liest jede Verwaltungsstation diese aus ihrem lokalen C:\Windows\PolicyDefinitions - bei unterschiedlichen Patchständen entstehen abweichende oder fehlende Richtlinien-Definitionen ("SYSVOL bloat" entfällt mit ADMX zusätzlich). Der Central Store unter \\<Domäne>\SYSVOL\<Domäne>\Policies\PolicyDefinitions wird von den GPO-Werkzeugen bevorzugt und repliziert mit SYSVOL auf alle DCs.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Create and Manage the Central Store for Group Policy ADMX templates'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-client/group-policy/create-and-manage-central-store' }
        )
    }
    'domain_controller' = @{
        Titel = 'Domain Controller (Übersicht)'; Schwere = 'Info'
        Zweck = 'Listet die Domänencontroller des Forest mit Betriebssystem, Standort (Site) und Global-Catalog-Rolle auf. Verschafft den Überblick, welche DCs vorhanden sind.'
        Beispiel = 'Ein DC mit nicht mehr unterstütztem Betriebssystem (z. B. Windows Server 2012 R2 ohne ESU) erhält keine Sicherheitsupdates mehr und ist ein bevorzugtes Angriffsziel.'
        Empfehlung = 'DCs auf unterstützten, aktuell gepatchten Betriebssystemen betreiben; nicht mehr benötigte DCs sauber heruntergraden.'
        Hintergrund = 'Domänencontroller beantworten Authentifizierung (Kerberos/NTLM), LDAP-Anfragen und Replikation. Welche DC-Betriebssysteme zulässig sind, hängt an der Funktionsebene. Ein DC auf einem abgekündigten OS erhält keine Sicherheitsupdates mehr und ist als Tier-0-System (höchste Schutzklasse) besonders kritisch - eine Kompromittierung gefährdet die gesamte Domäne.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Active Directory Domain Services Functional Levels'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/active-directory-functional-levels' }
            @{ Titel = 'Microsoft Lifecycle - Windows Lifecycle FAQ'; Url = 'https://learn.microsoft.com/en-us/lifecycle/faq/windows' }
        )
    }
    'logging' = @{
        Titel = 'Logging auf Domain Controller(n)'; Schwere = 'Mittel'
        Zweck = 'Prüft Status des Eventlog-Dienstes und die Audit-Richtlinien der DCs. Ohne aktiviertes Auditing fehlen die Spuren, mit denen Angriffe überhaupt erkannt werden können.'
        Beispiel = 'Ist "Audit Kerberos Service Ticket Operations" deaktiviert, bleibt Kerberoasting unsichtbar. Ohne Anmelde-Auditing lassen sich Brute-Force-/Spraying-Angriffe nicht nachweisen.'
        Empfehlung = 'Advanced Audit Policy gemäß Microsoft-/CIS-Empfehlung konfigurieren (Anmeldungen, Kontenverwaltung, Verzeichnisdienstzugriff); Logs zentral in ein SIEM sammeln.'
        Hintergrund = 'Ohne aktiviertes Auditing auf den Domänencontrollern entstehen keine Protokolle - Angriffe wie Kerberoasting oder Passwort-Spraying bleiben dann unsichtbar und im Nachhinein nicht nachweisbar. Microsoft empfiehlt, Anmeldungen, Konten-/Gruppenänderungen und Verzeichniszugriffe gezielt zu protokollieren und die Ereignisse zentral (SIEM) zu sammeln, damit Auffälligkeiten erkannt werden.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - System Audit Policy recommendations'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/audit-policy-recommendations' }
        )
    }
    'trusts' = @{
        Titel = 'AD-Trusts'; Schwere = 'Mittel'
        Zweck = 'Wertet Vertrauensstellungen zu anderen Domänen/Forests aus: Richtung, Transitivität und SID-Filtering. Trusts sind potenzielle Angriffspfade über Domänengrenzen hinweg.'
        Beispiel = 'Fehlt bei einem Forest-Trust das SID-Filtering, kann ein Angreifer aus der vertrauten Domäne per SID-History Rechte in der eigenen Domäne erlangen (cross-forest privilege escalation).'
        Empfehlung = 'Nicht mehr benötigte Trusts entfernen; SID-Filtering/Quarantine für externe Trusts aktiviert lassen; selektive Authentifizierung prüfen.'
        Hintergrund = 'Eine Vertrauensstellung erlaubt Konten einer fremden Domäne, sich in der eigenen anzumelden - und ist damit ein möglicher Angriffsweg über Domänengrenzen hinweg. Eine Schutzfunktion (SID-Filtering) verhindert, dass aus der vertrauten Domäne unrechtmäßig erhöhte Rechte eingeschleust werden. Fehlt dieser Schutz bei einem externen Trust, kann ein Angreifer von dort Rechte in der eigenen Domäne erlangen.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1482 - Domain Trust Discovery'; Url = 'https://attack.mitre.org/techniques/T1482/' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'dns' = @{
        Titel = 'DNS'; Schwere = 'Niedrig'
        Zweck = 'Prüft DNS-Server-Einstellungen der DCs: Aging/Scavenging (Bereinigung veralteter Einträge), Forwarder und Zonenkonfiguration. Sauberes DNS ist Voraussetzung für Replikation und Anmeldung.'
        Beispiel = 'Ohne Scavenging sammeln sich veraltete Host-Einträge an; ein neuer Host kann die IP eines alten Eintrags erhalten und wird dadurch falsch aufgelöst.'
        Empfehlung = 'Aging/Scavenging mit sinnvollen Intervallen aktivieren; nur vertrauenswürdige Forwarder konfigurieren; veraltete Zonen bereinigen.'
        Hintergrund = 'AD-integriertes DNS speichert dynamisch registrierte Einträge mit Zeitstempel. Aging/Scavenging löscht Einträge, deren Zeitstempel älter als No-Refresh- plus Refresh-Intervall ist. Scavenging muss an drei Stellen aktiv sein (Eintrag, Zone, mindestens ein Server) und ist standardmäßig deaktiviert. Ohne Scavenging sammeln sich veraltete Einträge; eine neu vergebene IP kann auf einen alten Namen verweisen (Fehlauflösung, potenziell Spoofing).'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - DNS Aging and Scavenging in Windows Server'; Url = 'https://learn.microsoft.com/en-us/windows-server/networking/dns/aging-scavenging' }
        )
    }
    'sysvol_health' = @{
        Titel = 'Sysvol Replication & AD-Health'; Schwere = 'Mittel'
        Zweck = 'Prüft die SYSVOL-Replikation (DFS-R statt des veralteten FRS) und die AD-Replikationsgesundheit. SYSVOL trägt Gruppenrichtlinien und Anmeldeskripte - Replikationsfehler führen zu uneinheitlichen Richtlinien.'
        Beispiel = 'Repliziert SYSVOL noch über FRS, ist die Umgebung nicht migriert und auf modernen DCs nicht mehr unterstützt; GPO-Änderungen kommen evtl. nicht auf allen DCs an.'
        Empfehlung = 'Von FRS auf DFS-R migrieren (falls noch nicht geschehen); Replikationsfehler (repadmin) regelmäßig prüfen und beheben.'
        Hintergrund = 'SYSVOL ist die replizierte Freigabe mit GPO-Dateien und Anmeldeskripten. Die Replikation erfolgt über FRS (veraltet, fehleranfällig) oder DFS-R (ab Funktionsebene 2008). Windows Server 2016 ist die letzte Version mit FRS-Unterstützung; die Migration FRS->DFS-R ist eine Einbahnstraße. Replikationsfehler führen dazu, dass GPO-Änderungen nicht auf allen DCs ankommen - Richtlinien wirken dann uneinheitlich.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Migrate SYSVOL replication from FRS to DFS Replication'; Url = 'https://learn.microsoft.com/en-us/windows-server/storage/dfs-replication/migrate-sysvol-to-dfsr' }
        )
    }
    'admins' = @{
        Titel = 'Administratoren und Builtin Benutzer'; Schwere = 'Hoch'
        Zweck = 'Wertet hochprivilegierte Gruppen aus (Domänen-, Organisations- und Schema-Admins, Administratoren, Builtin-Konten und weitere als privilegiert markierte Konten). Diese Konten sind das primäre Ziel jedes Angreifers.'
        Beispiel = 'Ein vergessenes Dienstkonto in "Domain Admins" mit schwachem Passwort genügt, um die gesamte Domäne zu übernehmen. Je mehr Mitglieder, desto größer die Angriffsfläche.'
        Empfehlung = 'Mitgliederzahl privilegierter Gruppen minimieren (Tier-0-Modell); Enterprise/Schema Admins im Normalbetrieb leer halten; dedizierte Admin-Konten, kein Tagesgeschäft mit Admin-Rechten.'
        Hintergrund = 'Die höchstprivilegierten Gruppen (u. a. Domänen-, Organisations- und Schema-Admins, Administratoren, die Operatoren-Gruppen) verfügen über die mächtigsten Rechte der Domäne und sind das bevorzugte Ziel von Angreifern. Je mehr Mitglieder sie haben, desto größer die Angriffsfläche. Best Practice ist, diese Gruppen so klein wie möglich zu halten und privilegierte Konten strikt von der täglichen Arbeit zu trennen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Appendix C: Protected Accounts and Groups in Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'benutzer' = @{
        Titel = 'Benutzer und Benutzer Accounts'; Schwere = 'Mittel'
        Zweck = 'Untersucht Benutzerkonten auf inaktive, gesperrte und falsch platzierte Konten (Standard-OU "Users"). Verwaiste Konten sind unbeaufsichtigte Einfallstore.'
        Beispiel = 'Ein seit zwei Jahren inaktives, aber aktiviertes Konto eines ausgeschiedenen Mitarbeiters lässt sich übernehmen, ohne dass es auffällt.'
        Empfehlung = 'Inaktive Konten deaktivieren und nach Frist löschen; Konten in passende OUs strukturieren; Joiner-/Mover-/Leaver-Prozess etablieren.'
        Hintergrund = 'Verwaiste, inaktive oder falsch platzierte Benutzerkonten vergrößern die Angriffsfläche, ohne Nutzen zu stiften. Inaktive, aber aktivierte Konten (lastLogonTimestamp lange her) sind unbeaufsichtigte Anmeldeziele - besonders, wenn ihr Passwort nie geändert wurde. Ein sauberer Joiner-/Mover-/Leaver-Prozess und regelmäßige Rezertifizierung halten den Bestand aktuell.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'computerkonten' = @{
        Titel = 'Computerkonten'; Schwere = 'Info'
        Zweck = 'Erfasst die Computerkonten der Domäne. Liefert den Bestand und hilft, verwaiste oder veraltete Maschinenkonten zu erkennen.'
        Beispiel = 'Ein Computerkonto ohne jüngste Anmeldung deutet auf ein ausgemustertes Gerät hin, dessen Konto noch missbraucht werden könnte.'
        Empfehlung = 'Veraltete Computerkonten regelmäßig identifizieren und entfernen.'
        Hintergrund = 'Computerkonten authentifizieren Maschinen gegen AD; ihr Passwort wird normalerweise alle 30 Tage automatisch erneuert. Konten ohne jüngste Anmeldung deuten auf ausgemusterte Geräte hin und sollten entfernt werden. Selbst angelegte Computerkonten sind zudem Baustein von RBCD- und noPac-Angriffen (siehe Computerkonten-Kontingent/MachineAccountQuota im Kerberos-Bereich).'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'clients' = @{
        Titel = 'Client Check'; Schwere = 'Niedrig'
        Zweck = 'Listet Client-Betriebssysteme und prüft, ob sie noch vom Hersteller unterstützt werden. Nicht unterstützte Systeme erhalten keine Sicherheitsupdates.'
        Beispiel = 'Ein verbliebenes Windows 7 oder ausgelaufenes Windows 10 ist über bekannte, ungepatchte Lücken angreifbar und kann als Sprungbrett dienen.'
        Empfehlung = 'Nicht unterstützte Clients ausmustern oder isolieren; Patch-Management sicherstellen.'
        Hintergrund = 'Microsoft-Betriebssysteme folgen einem festen Lebenszyklus; nach dem End-of-Support gibt es keine Sicherheitsupdates mehr (außer kostenpflichtige Extended Security Updates). Windows 10 erreicht z. B. am 14.10.2025 das Support-Ende. Nicht mehr unterstützte Clients sind über bekannte, ungepatchte Lücken angreifbar und eignen sich als Einstiegs- und Sprungbrett-Systeme im Netz.'
        Quellen = @(
            @{ Titel = 'Microsoft Lifecycle - Windows Lifecycle FAQ'; Url = 'https://learn.microsoft.com/en-us/lifecycle/faq/windows' }
        )
    }
    'server' = @{
        Titel = 'Server Check'; Schwere = 'Niedrig'
        Zweck = 'Listet Server-Betriebssysteme und deren Support-Status. Wie bei Clients sind nicht unterstützte Server ein erhöhtes Risiko.'
        Beispiel = 'Ein Windows Server 2008 R2 ohne ESU ist dauerhaft verwundbar; fällt er, können darauf gespeicherte Anmeldedaten den Angriff ausweiten.'
        Empfehlung = 'Server auf unterstützte Versionen heben; Altsysteme isolieren; Patch-Stand überwachen.'
        Hintergrund = 'Wie Clients folgen Server-Betriebssysteme dem Microsoft-Lebenszyklus. Ein nicht mehr unterstützter Server (z. B. Windows Server 2008 R2 ohne ESU) bleibt dauerhaft verwundbar; wird er kompromittiert, lassen sich darauf zwischengespeicherte Anmeldedaten für laterale Bewegung nutzen. Welche Server-Versionen als DC zulässig sind, hängt zusätzlich an der Funktionsebene.'
        Quellen = @(
            @{ Titel = 'Microsoft Lifecycle - Windows Lifecycle FAQ'; Url = 'https://learn.microsoft.com/en-us/lifecycle/faq/windows' }
            @{ Titel = 'Microsoft Learn - Active Directory Domain Services Functional Levels'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/active-directory-functional-levels' }
        )
    }
    'nicht_windows' = @{
        Titel = 'Nicht-Windows-Systeme'; Schwere = 'Info'
        Zweck = 'Erfasst in der Domäne registrierte Nicht-Windows-Systeme. Schafft Transparenz über heterogene Geräte (Linux, Appliances), die ebenfalls Konten besitzen.'
        Beispiel = 'Ein an die Domäne angebundenes Linux-System mit veralteter Konfiguration kann eigene Schwachstellen einbringen.'
        Empfehlung = 'Nicht-Windows-Beitritte dokumentieren und in das Patch-/Härtungskonzept einbeziehen.'
        Hintergrund = 'Auch Nicht-Windows-Systeme (Linux, Appliances, NAS) können der Domäne beitreten und besitzen dann AD-Konten. Sie unterliegen nicht dem Windows-Patch-/Härtungsprozess und bringen eigene Schwachstellen mit; Transparenz über solche Systeme ist Voraussetzung für ein vollständiges Härtungs- und Monitoring-Konzept.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'ad_gruppen' = @{
        Titel = 'AD-Gruppen'; Schwere = 'Niedrig'
        Zweck = 'Wertet AD-Gruppen aus (Anzahl, Typ, leere oder verschachtelte Gruppen). Unübersichtliche Gruppenstrukturen führen zu schleichender Rechteanhäufung.'
        Beispiel = 'Tief verschachtelte Gruppen verschleiern, wer am Ende welche Rechte hat - so entstehen ungewollte Berechtigungen.'
        Empfehlung = 'Gruppenmodell aufräumen (z. B. AGDLP), leere/verwaiste Gruppen entfernen, Verschachtelung begrenzen, regelmäßig rezertifizieren.'
        Hintergrund = 'Gruppen lassen sich ineinander verschachteln. Tiefe oder unübersichtliche Verschachtelung verschleiert, wer am Ende welche Rechte besitzt, und führt dazu, dass Konten unbemerkt immer mehr Berechtigungen ansammeln. Ein klares, flaches Gruppenmodell und regelmäßiges Aufräumen halten die Rechtevergabe nachvollziehbar.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Active Directory Security Groups (scope and nesting)'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-security-groups' }
        )
    }
    'gpos' = @{
        Titel = 'GPOs'; Schwere = 'Mittel'
        Zweck = 'Erfasst die Gruppenrichtlinienobjekte (GPOs), u. a. nicht verknüpfte oder leere GPOs und die Domänen-Standardrichtlinien. GPOs steuern zentrale Sicherheitseinstellungen.'
        Beispiel = 'Eine GPO, deren Bearbeitungsrecht an eine breite Gruppe vergeben ist, erlaubt Angreifern, über die GPO Code auf vielen Systemen auszuführen.'
        Empfehlung = 'Nicht verknüpfte/leere GPOs entfernen; GPO-Bearbeitungsrechte streng begrenzen; Änderungen versionieren und dokumentieren.'
        Hintergrund = 'GPOs steuern zentrale Sicherheits- und Systemeinstellungen und werden über SYSVOL verteilt. Wer eine verknüpfte GPO bearbeiten darf, kann auf allen davon erfassten Systemen Einstellungen oder Skripte ausrollen - daher ist die Delegation der GPO-Bearbeitung sicherheitskritisch (MITRE T1484: Domain or Tenant Policy Modification). Nicht verknüpfte oder leere GPOs sind meist Altlasten und sollten entfernt werden.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1484 - Domain or Tenant Policy Modification'; Url = 'https://attack.mitre.org/techniques/T1484/' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'ddp_password' = @{
        Titel = 'dDP Password Settings'; Schwere = 'Hoch'
        Zweck = 'Prüft die Default Domain Password Policy (Mindestlänge, Komplexität, Historie, Sperrschwelle). Diese Richtlinie bestimmt die Grundsicherheit aller Domänenpasswörter.'
        Beispiel = 'Eine Mindestlänge von 7 Zeichen ohne Sperrschwelle ermöglicht praktikable Brute-Force- und Password-Spraying-Angriffe.'
        Empfehlung = 'Mindestens 14 Zeichen, Sperrschwelle setzen (Microsoft-Baseline empfiehlt 10), gegen bekannte/kompromittierte Passwörter prüfen (z. B. Azure AD Password Protection); statt erzwungenem periodischem Wechsel auf längere Passphrasen setzen (NIST SP 800-63B).'
        Hintergrund = 'Die Default Domain Password Policy gilt domänenweit für alle Konten ohne abweichende fGPP. Mindestlänge (in der GUI max. 14), Komplexität, Historie und Sperrschwelle bestimmen den Brute-Force-/Spraying-Widerstand. Eine fehlende oder zu hohe Sperrschwelle macht Password-Spraying praktikabel; eine zu niedrige begünstigt DoS durch Aussperren. Microsofts Security-Baseline nennt als Richtwert eine Sperrschwelle von 10.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Minimum password length (Security policy setting)'; Url = 'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/security-policy-settings/minimum-password-length' }
        )
    }
    'fgpp' = @{
        Titel = 'Fine Grained Password Policies'; Schwere = 'Mittel'
        Zweck = 'Wertet Fine-Grained Password Policies (PSOs) aus, mit denen abweichende Passwortrichtlinien für einzelne Gruppen/Konten gelten - wichtig vor allem für privilegierte und Dienstkonten.'
        Beispiel = 'Fehlt eine strengere Richtlinie für Administratoren oder Dienstkonten, gilt für sie nur die oft schwächere Standardrichtlinie.'
        Empfehlung = 'Für privilegierte Konten und Dienstkonten strengere PSOs definieren (länger, häufigerer Wechsel bzw. gMSA).'
        Hintergrund = 'Fine-Grained Password Policies (Password Settings Objects, PSOs) erlauben abweichende Passwort-/Sperr-Richtlinien für einzelne globale Sicherheitsgruppen oder Benutzer - zusätzlich zur Default Domain Password Policy. Bei mehreren zutreffenden PSOs gewinnt die mit der niedrigsten Precedence. PSOs greifen nur auf Gruppen/Benutzer, nicht direkt auf OUs. Damit lassen sich für Admins und Dienstkonten strengere Vorgaben durchsetzen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Configure fine grained password policies for AD DS'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/get-started/adac/fine-grained-password-policies' }
        )
    }
    'user_vs_pw' = @{
        Titel = 'User vs Password Policies'; Schwere = 'Mittel'
        Zweck = 'Gleicht Benutzerkonten gegen passwortbezogene Risikomerkmale ab, u. a. "Passwort läuft nie ab" und "Passwort nicht erforderlich". Solche Konten unterlaufen die Passwortrichtlinie.'
        Beispiel = 'Ein Konto mit "Password never expires" und schwachem Passwort bleibt dauerhaft angreifbar; "Password not required" erlaubt im Extremfall ein leeres Passwort.'
        Empfehlung = 'Flags PASSWD_NOTREQD und DONT_EXPIRE_PASSWORD prüfen und entfernen (Ausnahmen nur für gMSA/begründete Fälle).'
        Hintergrund = 'Einzelne Konten können so eingestellt sein, dass ihr Passwort nie abläuft oder gar keines erforderlich ist (im Extremfall ein leeres Passwort). Solche Konten unterlaufen die Passwortrichtlinie und bleiben dauerhaft ein leichtes Ziel. Vertretbar ist das nur für automatisch verwaltete Dienstkonten (gMSA) oder begründete technische Ausnahmen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - UserAccountControl property flags'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/useraccountcontrol-manipulate-account-properties' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'ous' = @{
        Titel = 'Organisation Units'; Schwere = 'Info'
        Zweck = 'Stellt die OU-Struktur dar und prüft optional den Schutz vor versehentlichem Löschen. Eine saubere OU-Struktur ist Basis für gezielte GPO-Verknüpfung und Delegation.'
        Beispiel = 'Eine flache oder chaotische OU-Struktur erschwert das Tiering und führt dazu, dass GPOs zu breit greifen.'
        Empfehlung = 'OU-Struktur an Verwaltung/Tiering ausrichten; Schutz vor versehentlichem Löschen aktivieren; Delegationen dokumentieren.'
        Hintergrund = 'Organisationseinheiten strukturieren AD-Objekte für gezielte GPO-Verknüpfung und delegierte Verwaltung. Eine an Tiering/Verwaltung ausgerichtete OU-Struktur verhindert, dass GPOs zu breit greifen, und ermöglicht eine least-privilege-Delegation. Der Schutz "vor versehentlichem Löschen" (ProtectedFromAccidentalDeletion) verhindert das versehentliche Entfernen ganzer Teilbäume.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'msa' = @{
        Titel = 'Managed Service Accounts (MSA/gMSA)'; Schwere = 'Niedrig'
        Zweck = 'Prüft (group) Managed Service Accounts und den KDS-Root-Key. (g)MSA bieten automatisch verwaltete, sehr lange Passwörter und sind die sichere Alternative zu klassischen Dienstkonten.'
        Beispiel = 'Laufen Dienste noch unter klassischen Konten mit fixem Passwort und SPN, sind sie Kerberoasting-fähig - ein gMSA wäre dagegen praktisch nicht knackbar.'
        Empfehlung = 'Dienste auf gMSA umstellen; KDS-Root-Key bereitstellen; klassische Dienstkonten ablösen.'
        Hintergrund = 'Ein gMSA ist ein Dienstkonto, dessen sehr langes Passwort von den Domänencontrollern automatisch erzeugt und regelmäßig gewechselt wird - niemand muss es kennen oder pflegen. Dadurch sind solche Konten praktisch nicht angreifbar (kein Kerberoasting) und die sichere Alternative zu klassischen Dienstkonten mit festem Passwort.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Group Managed Service Accounts overview'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/group-managed-service-accounts/group-managed-service-accounts/group-managed-service-accounts-overview' }
        )
    }
    'ca' = @{
        Titel = 'Zertifizierungsstelle(n)'; Schwere = 'Mittel'
        Zweck = 'Erfasst die Zertifizierungsstellen (Root/Sub-CA) der Umgebung (AD CS). Die PKI ist sicherheitskritisch: Wer Zertifikate ausstellen kann, kann sich als beliebiger Benutzer ausgeben.'
        Beispiel = 'Eine falsch konfigurierte Vorlage kann es jedem Benutzer erlauben, ein Anmeldezertifikat für einen Administrator zu beantragen (ESC1) - die Detailprüfung erfolgt im AD-CS-Sicherheitscheck.'
        Empfehlung = 'CA-Rollen und Vorlagen regelmäßig auf Fehlkonfigurationen prüfen (ESC1-ESC8); Ausstellungsrechte streng begrenzen.'
        Hintergrund = 'Die hauseigene Zertifikatsstelle (AD CS) stellt u. a. Zertifikate aus, mit denen man sich anmelden kann. Ist eine Zertifikatvorlage falsch konfiguriert, kann ein normaler Benutzer sich darüber ein Zertifikat auf einen Administrator ausstellen und dessen Identität übernehmen. Diese Schwachstellen sind als ESC1-ESC8 bekannt; sie werden im AD-CS-Sicherheitscheck im Detail geprüft.'
        Quellen = @(
            @{ Titel = 'SpecterOps (Schroeder/Christensen) - Certified Pre-Owned: Abusing AD CS'; Url = 'https://specterops.io/blog/2021/06/17/certified-pre-owned/' }
        )
    }
    'dc_detail' = @{
        Titel = 'Domain Controller (Detailprüfung)'; Schwere = 'Mittel'
        Zweck = 'Führt pro Domänencontroller Detailprüfungen durch (Dienste, Rollen, Features, LDAPS, NTLM, SMB1, BitLocker, ExecutionPolicy). Der DC ist das Herz der Domäne; seine Härtung ist entscheidend.'
        Beispiel = 'Ist SMB1 auf einem DC noch aktiv, ist er über längst bekannte Lücken (z. B. EternalBlue) angreifbar; fehlendes LDAPS erlaubt das Mitlesen von Verzeichnisanfragen.'
        Empfehlung = 'SMB1 entfernen, LDAPS/LDAP-Signing erzwingen, NTLM einschränken, DCs nach CIS-/Microsoft-Baseline härten.'
        Hintergrund = 'Die Detailprüfung betrachtet je Domänencontroller Dienste, Rollen, Features und mehrere Härtungsindikatoren: das veraltete, angreifbare SMB1, den Schutz der Verzeichnis-Anmeldung (LDAPS/LDAP-Signing), die NTLM-Stufe, BitLocker und die PowerShell-Ausführungsrichtlinie. Microsofts Security Baselines und CIS-Benchmarks liefern dafür geprüfte Soll-Werte.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Windows security baselines guide'; Url = 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/device-management/windows-security-configuration-framework/windows-security-baselines' }
            @{ Titel = 'Microsoft Learn - Microsoft Security Compliance Toolkit'; Url = 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/device-management/windows-security-configuration-framework/security-compliance-toolkit-10' }
        )
    }
    'kerberos' = @{
        Titel = 'Kerberos - Angriffsflächen'; Schwere = 'Hoch'
        Zweck = 'Bündelt die wichtigsten Kerberos-bezogenen Angriffsflächen einer AD-Umgebung: angreifbare Dienstkonten (SPN), Konten ohne Vorauthentifizierung, missbrauchbare Delegation, schwache Verschlüsselung und das Computerkonten-Kontingent.'
        Hintergrund = 'Kerberos ist das zentrale Anmeldeverfahren in Active Directory: Nutzer erhalten vom Domänencontroller Tickets, mit denen sie sich an Diensten anmelden. Bestimmte Konto-Einstellungen (z. B. hinterlegte Dienstnamen, fehlende Vorauthentifizierung, Delegationsrechte oder veraltete Verschlüsselung) machen diese Tickets angreifbar - oft so, dass ein normaler Benutzer ohne Adminrechte an privilegierte Anmeldedaten gelangt. Die folgenden Teilprüfungen lesen genau diese Einstellungen read-only aus.'
        Beispiel = 'Diese Schwachstellen erlauben es einem normalen Domänenbenutzer häufig, ohne Adminrechte an privilegierte Anmeldedaten zu gelangen - ein klassischer Einstieg in die Domänenübernahme.'
        Empfehlung = 'Die einzelnen Teilprüfungen abarbeiten; Dienstkonten auf gMSA umstellen, AES erzwingen, Delegation minimieren und das Kontingent auf 0 setzen.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1558 - Steal or Forge Kerberos Tickets'; Url = 'https://attack.mitre.org/techniques/T1558/' }
            @{ Titel = 'Microsoft Learn - UserAccountControl property flags'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/useraccountcontrol-manipulate-account-properties' }
        )
    }
    'kerberoasting' = @{
        Titel = 'Kerberoasting (Konten mit SPN)'; Schwere = 'Hoch'
        Zweck = 'Listet aktivierte Benutzerkonten (keine Computer) mit gesetztem servicePrincipalName (SPN) auf, krbtgt ausgenommen. Treffer, die zu privilegierten Konten gehören, werden zusätzlich hervorgehoben.'
        Hintergrund = 'Bei Kerberoasting fordert ein beliebiger angemeldeter Benutzer beim Domänencontroller ein Ticket für ein Dienstkonto an. Ein Teil dieses Tickets ist mit dem Passwort des Dienstkontos verschlüsselt und lässt sich in Ruhe offline knacken - ohne weitere Spuren am Domänencontroller. Schwache Dienstkonto-Passwörter fallen so in Minuten bis Stunden; besonders kritisch, wenn das Konto privilegiert ist.'
        Beispiel = 'Jeder Domänenbenutzer fordert für ein SPN-Konto wie "svc-sql" ein Ticket an und knackt es offline. Steckt das Konto in "Domain Admins", ist die Domäne kompromittiert.'
        Empfehlung = 'Dienstkonten auf (group) Managed Service Accounts umstellen (automatische 120-Zeichen-Passwörter); sonst Passwörter >= 25 Zeichen; msDS-SupportedEncryptionTypes der Konten auf AES beschränken (entzieht RC4-Tickets die Grundlage); privilegierte Konten nie mit SPN betreiben.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1558.003 - Kerberoasting'; Url = 'https://attack.mitre.org/techniques/T1558/003/' }
            @{ Titel = 'Microsoft Learn - Service Principal Names'; Url = 'https://learn.microsoft.com/en-us/windows/win32/ad/service-principal-names' }
            @{ Titel = 'adsecurity.org (Sean Metcalf) - Cracking Kerberos TGS Tickets Using Kerberoast'; Url = 'https://adsecurity.org/?p=2293' }
        )
    }
    'asrep' = @{
        Titel = 'AS-REP Roasting (ohne Vorauthentifizierung)'; Schwere = 'Hoch'
        Zweck = 'Findet aktivierte Konten, bei denen die Kerberos-Vorauthentifizierung abgeschaltet ist ("Pre-Authentication nicht erforderlich").'
        Hintergrund = 'Normalerweise muss ein Konto bei der Anmeldung zuerst sein Passwort nachweisen (Vorauthentifizierung). Ist diese abgeschaltet, kann ein Angreifer für das Konto ohne jede eigene Anmeldung eine Antwort vom Domänencontroller anfordern, deren verschlüsselter Teil vom Benutzerpasswort abhängt - und ihn anschließend offline knacken. Schwache Passwörter solcher Konten sind damit leicht angreifbar.'
        Beispiel = 'Ein Angreifer mit reiner Netzwerksicht zählt Konten ohne Pre-Auth auf, holt deren AS-REP und knackt schwache Passwörter offline.'
        Empfehlung = 'Flag "Kerberos-Preauthentifizierung nicht erforderlich" entfernen, wo immer möglich; verbleibende Konten mit langen, starken Passwörtern versehen oder deaktivieren.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1558.004 - AS-REP Roasting'; Url = 'https://attack.mitre.org/techniques/T1558/004/' }
            @{ Titel = 'Microsoft Learn - UserAccountControl flags (DONT_REQ_PREAUTH 0x400000)'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/useraccountcontrol-manipulate-account-properties' }
        )
    }
    'delegation' = @{
        Titel = 'Delegation (Unconstrained / Constrained / RBCD)'; Schwere = 'Hoch'
        Zweck = 'Prüft Konten und Computer mit Kerberos-Delegationsrechten in den drei Varianten uneingeschränkt, eingeschränkt und ressourcenbasiert (RBCD). Domänencontroller werden ausgenommen.'
        Hintergrund = 'Delegation erlaubt einem Dienst, im Namen eines Benutzers auf weitere Dienste zuzugreifen. Bei der uneingeschränkten Variante hält der Server die Anmeldeinformationen jedes ankommenden Benutzers vor - wird er kompromittiert, lassen sich darüber auch Anmeldedaten von Administratoren abgreifen und die Domäne übernehmen. Die eingeschränkten Varianten sind sicherer, bei Fehlkonfiguration aber ebenfalls ein verbreiteter Weg zur Rechteausweitung (häufig kombiniert mit selbst angelegten Computerkonten).'
        Beispiel = 'Ein Computer mit uneingeschränkter Delegation, der von einem Domänenadministrator kontaktiert wird, hält dessen TGT vor - der Angreifer extrahiert es und übernimmt die Domäne. RBCD-Einträge erlauben oft eine direkte Übernahme des Zielsystems.'
        Empfehlung = 'Uneingeschränkte Delegation vermeiden; auf eingeschränkte Delegation ohne Protocol Transition umstellen; sensible Konten als "Konto ist vertraulich und kann nicht delegiert werden" (NOT_DELEGATED) markieren bzw. in "Protected Users" aufnehmen; RBCD-Einträge regelmäßig prüfen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Kerberos Constrained Delegation Overview'; Url = 'https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-constrained-delegation-overview' }
            @{ Titel = 'Microsoft Learn - UserAccountControl flags (TRUSTED_FOR_DELEGATION 0x80000)'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/useraccountcontrol-manipulate-account-properties' }
            @{ Titel = 'Shenanigans Labs (Elad Shamir) - Wagging the Dog: Abusing RBCD'; Url = 'https://shenaniganslabs.io/2019/01/28/Wagging-the-Dog.html' }
        )
    }
    'kerb_enc' = @{
        Titel = 'Schwache Kerberos-Verschlüsselung'; Schwere = 'Mittel'
        Zweck = 'Findet Konten, die auf die veraltete DES-Verschlüsselung festgelegt sind.'
        Hintergrund = 'Kerberos-Tickets können unterschiedlich stark verschlüsselt sein. DES und RC4 gelten als veraltet bzw. gebrochen und erleichtern Angriffe wie das Offline-Knacken von Tickets. Konten, die noch fest auf DES eingestellt sind, geben besonders schwach geschützte Tickets aus. Sicher ist die AES-Verschlüsselung, die sich pro Konto oder per Gruppenrichtlinie erzwingen lässt.'
        Beispiel = 'Ein Dienstkonto mit erzwungenem DES gibt schwach verschlüsselte Tickets aus, die sich mit heutiger Hardware extrem schnell brechen lassen.'
        Empfehlung = 'USE_DES_KEY_ONLY entfernen; msDS-SupportedEncryptionTypes der Konten auf AES (Wert 0x18 = AES128 + AES256) setzen und RC4 schrittweise per GPO "Configure encryption types allowed for Kerberos" abschalten.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Network security: Configure encryption types allowed for Kerberos'; Url = 'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/security-policy-settings/network-security-configure-encryption-types-allowed-for-kerberos' }
            @{ Titel = 'Microsoft Learn - UserAccountControl flags (USE_DES_KEY_ONLY 0x200000)'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/useraccountcontrol-manipulate-account-properties' }
        )
    }
    'machine_quota' = @{
        Titel = 'Computerkonten-Kontingent (MachineAccountQuota)'; Schwere = 'Hoch'
        Zweck = 'Liest aus, wie viele Computerkonten ein normaler Benutzer selbst in die Domäne aufnehmen darf (Einstellung ms-DS-MachineAccountQuota). Standard ist 10.'
        Hintergrund = 'Standardmäßig darf jeder Benutzer bis zu 10 eigene Computerkonten anlegen. Solche selbst angelegten Konten sind Baustein mehrerer bekannter Angriffe, mit denen sich Rechte bis zur Übernahme eines Servers oder sogar eines Domänencontrollers ausweiten lassen (u. a. die "noPac"-Lücke von 2021). In den meisten Umgebungen wird diese Möglichkeit nicht benötigt.'
        Beispiel = 'Ein Standardbenutzer legt ein Computerkonto an, trägt es als RBCD-Prinzipal an einem Zielserver ein und übernimmt diesen - ohne jegliche Adminrechte.'
        Empfehlung = 'ms-DS-MachineAccountQuota auf 0 setzen und das Anlegen von Computerkonten an eine dedizierte, delegierte Gruppe binden. noPac-Patches (November 2021) einspielen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - MS-DS-Machine-Account-Quota attribute'; Url = 'https://learn.microsoft.com/en-us/windows/win32/adschema/a-ms-ds-machineaccountquota' }
            @{ Titel = 'Shenanigans Labs (Elad Shamir) - Wagging the Dog (MachineAccountQuota + RBCD)'; Url = 'https://shenaniganslabs.io/2019/01/28/Wagging-the-Dog.html' }
            @{ Titel = 'NVD - CVE-2021-42278 (noPac, sAMAccountName Spoofing)'; Url = 'https://nvd.nist.gov/vuln/detail/CVE-2021-42278' }
        )
    }
    'privilegien' = @{
        Titel = 'Privilegien & ACLs'; Schwere = 'Hoch'
        Zweck = 'Bündelt Prüfungen rund um privilegierte Rechte und gefährliche Berechtigungen im Verzeichnis: DCSync-Rechte, riskante Operatoren-/Admin-Gruppen, die AdminSDHolder-ACL, die Nutzung von Protected Users und die Pre-Windows-2000-Kompatibilität.'
        Hintergrund = 'Nicht nur Gruppenmitgliedschaften, sondern auch einzelne Berechtigungen im Verzeichnis entscheiden, wer mächtige Aktionen ausführen darf. Schon ein einzelnes falsch vergebenes Recht (etwa zum Abruf aller Passwörter oder zum Ändern der Administrator-Vorlage) kann einem unprivilegierten Konto die vollständige Übernahme der Domäne ermöglichen. Diese Teilprüfungen lesen die betroffenen Gruppen und Berechtigungen read-only aus.'
        Beispiel = 'Ein an einen Anwendungs-Account delegiertes "Replicating Directory Changes All" genügt, um per DCSync alle Passwort-Hashes (inkl. krbtgt) abzuziehen.'
        Empfehlung = 'Privilegierte Rechte und ACLs auf das Nötigste reduzieren (Tier-0-Modell); delegierte Sonderrechte regelmäßig überprüfen und dokumentieren.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
            @{ Titel = 'Microsoft Learn - Appendix C: Protected Accounts and Groups in Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory' }
        )
    }
    'dcsync' = @{
        Titel = 'DCSync-Rechte (Verzeichnis-Replikation)'; Schwere = 'Kritisch'
        Zweck = 'Prüft, welche Konten und Gruppen die Replikationsrechte besitzen, mit denen sich Passwörter aus dem Verzeichnis abrufen lassen - und meldet alle, die dort nicht hingehören (normal nur Domänencontroller, Administratoren, SYSTEM).'
        Hintergrund = 'Mit diesen Replikationsrechten (bekannt als "DCSync") kann sich ein Konto wie ein Domänencontroller verhalten und die Passwörter aller Benutzer abfragen - inklusive des zentralen krbtgt-Schlüssels, mit dem sich praktisch beliebige Identitäten fälschen lassen. Das funktioniert aus der Ferne, ohne Anmeldung an einem DC, und ist ein direkter Weg zur vollständigen Domänenübernahme. Solche Rechte sollten ausschließlich Domänencontroller besitzen.'
        Beispiel = 'Ein übernommenes Konto mit diesen Rechten liest alle Passwort-Hashes aus und fälscht damit ein Admin-Ticket ("Golden Ticket") - die gesamte Domäne ist kompromittiert.'
        Empfehlung = 'Replikationsrechte ausschließlich Domänencontrollern gewähren; versehentlich delegierte Get-Changes(-All)-ACEs an Benutzern/Gruppen/Dienstkonten entfernen; Hybrid-Konten (z. B. Azure AD Connect / MSOL_) gezielt absichern und überwachen.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1003.006 - OS Credential Dumping: DCSync'; Url = 'https://attack.mitre.org/techniques/T1003/006/' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'operatoren' = @{
        Titel = 'Gefährliche Builtin-/Operatoren-Gruppen'; Schwere = 'Hoch'
        Zweck = 'Listet die Mitglieder sicherheitskritischer, oft übersehener Gruppen: Account/Server/Print/Backup Operators sowie Schema- und Enterprise Admins. Diese Gruppen sollten im Normalbetrieb leer oder minimal besetzt sein.'
        Hintergrund = 'Diese oft übersehenen Standardgruppen verleihen indirekt sehr weitreichende Macht: Account Operators können nahezu beliebige Konten verwalten; Backup-/Server Operators dürfen sich an Domänencontrollern anmelden bzw. die AD-Datenbank sichern und damit alle Passwörter auslesen; Schema- und Organisations-Admins sind forestweit allmächtig. Im Normalbetrieb sollten diese Gruppen leer oder minimal besetzt sein.'
        Beispiel = 'Ein Mitglied von Backup Operators sichert die NTDS.dit eines DCs und extrahiert offline alle Passwort-Hashes - ohne je Domain Admin gewesen zu sein.'
        Empfehlung = 'Mitgliedschaften dieser Gruppen entfernen bzw. auf das absolut Notwendige beschränken; statt dauerhafter Mitgliedschaft Just-in-time-Modelle nutzen; DnsAdmins gesondert prüfen (DLL-Ladepfad auf DCs).'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Appendix C: Protected Accounts and Groups in Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory' }
            @{ Titel = 'Microsoft Learn - Active Directory Security Groups'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-security-groups' }
        )
    }
    'adminsdholder' = @{
        Titel = 'AdminSDHolder-ACL'; Schwere = 'Hoch'
        Zweck = 'Prüft die Berechtigungen des AdminSDHolder-Objekts auf Konten, die dort nicht standardmäßig Schreib- oder Vollzugriff haben sollten.'
        Hintergrund = 'AdminSDHolder dient als Vorlage für die Berechtigungen aller privilegierten Konten und Gruppen: Das System überträgt seine Berechtigungen regelmäßig (etwa stündlich) auf diese Konten. Trägt ein Angreifer sich hier ein, erhält er dauerhaften Schreibzugriff auf sämtliche Administrator-Konten - eine versteckte Hintertür, die auch nach dem Zurücksetzen einzelner Konten bestehen bleibt.'
        Beispiel = 'Ein unscheinbares Konto mit Schreibrecht auf AdminSDHolder erhält darüber automatisch Zugriff auf die Gruppe der Domänen-Admins und kann sich selbst Vollzugriff auf jedes Admin-Konto verschaffen.'
        Empfehlung = 'Auf AdminSDHolder nur die Standard-Prinzipale (SYSTEM, Domain/Enterprise Admins, Administratoren) mit Schreibrechten zulassen; abweichende ACEs entfernen und ihre Herkunft untersuchen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Appendix C: Protected Accounts and Groups in Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'protected_users' = @{
        Titel = 'Protected Users (Nutzung)'; Schwere = 'Mittel'
        Zweck = 'Prüft, ob die Gruppe "Protected Users" verwendet wird, und listet ihre Mitglieder. Standardmäßig ist sie leer.'
        Hintergrund = 'Mitglieder der Gruppe "Protected Users" erhalten automatisch mehrere nicht abschaltbare Schutzmaßnahmen (u. a. keine NTLM-Anmeldung, nur starke Verschlüsselung, keine Delegation, kürzere Ticket-Gültigkeit). Das erschwert Diebstahl und Wiederverwendung ihrer Anmeldedaten erheblich. Dienst- und Computerkonten gehören hier nicht hinein.'
        Beispiel = 'Sind privilegierte Benutzerkonten nicht in Protected Users, lässt sich ihr NTLM-Hash per Pass-the-Hash weiterverwenden; als Mitglied wäre dieser Weg blockiert.'
        Empfehlung = 'Privilegierte Benutzerkonten (keine Dienst-/Computerkonten) nach sorgfältigem Test in Protected Users aufnehmen; Funktionsebene >= 2012 R2 und vorhandene AES-Schlüssel sicherstellen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Protected Users Security Group'; Url = 'https://learn.microsoft.com/en-us/windows-server/security/credentials-protection-and-management/protected-users-security-group' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'prewin2000' = @{
        Titel = 'Pre-Windows 2000 Compatible Access'; Schwere = 'Mittel'
        Zweck = 'Prüft die Mitgliedschaft der Gruppe "Pre-Windows 2000 Compatible Access" - kritisch ist die Aufnahme von "Jeder" (Everyone) oder "Anonymous-Anmeldung".'
        Hintergrund = 'Diese Gruppe existiert nur noch aus Kompatibilität mit sehr alten Systemen und gewährt Lesezugriff auf viele Verzeichnisobjekte. Enthält sie "Jeder" (Everyone) oder "Anonymous-Anmeldung", können selbst Benutzer ohne gültiges Konto umfangreiche Informationen wie Benutzer- und Gruppenlisten auslesen - wertvolle Aufklärung für einen Angreifer.'
        Beispiel = 'Ist "Anonymous Logon" Mitglied, kann ein Angreifer ohne gültiges Konto Benutzernamen und Gruppen enumerieren und so Spraying-/Roasting-Ziele finden.'
        Empfehlung = '"Everyone"/"Anonymous Logon" aus der Gruppe entfernen, sofern keine echten Legacy-Abhängigkeiten bestehen; die Mitgliedschaft auf das Nötigste reduzieren.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Active Directory Security Groups'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-security-groups' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'adcs' = @{
        Titel = 'AD CS - Zertifikatsdienste (ESC)'; Schwere = 'Hoch'
        Zweck = 'Untersucht die AD-Zertifikatsdienste (AD CS) auf die bekannten Eskalationspfade ESC1-ESC8: angreifbare Zertifikatvorlagen, manipulierbare Vorlagen-ACLs, gefährliche CA-Einstellungen und HTTP-Web-Enrollment.'
        Hintergrund = 'Die Zertifikatsdienste (AD CS) stellen u. a. Zertifikate aus, mit denen man sich anmelden kann. 2021 wurden acht Fehlkonfigurations-Klassen (ESC1-ESC8) bekannt, mit denen ein normaler Benutzer sich ein Zertifikat auf einen Administrator ausstellen und dessen Identität übernehmen kann - ein oft übersehener, direkter Weg zur Domänenübernahme. Vorlagen und Zertifizierungsstellen werden dafür read-only ausgelesen.'
        Beispiel = 'Eine einzige Vorlage, die "Antragsteller liefert Subject" erlaubt, Client-Authentication enthält und von "Domänen-Benutzer" angefordert werden darf, genügt für die vollständige Kompromittierung (ESC1).'
        Empfehlung = 'Vorlagen-Rechte und -Flags nach den ESC1-ESC8-Kriterien prüfen und härten; EDITF_ATTRIBUTESUBJECTALTNAME2 deaktivieren; Web-Enrollment absichern bzw. abschalten; Vorlagen-ACLs auf Schreibrechte für breite Gruppen kontrollieren.'
        Quellen = @(
            @{ Titel = 'SpecterOps (Schroeder/Christensen) - Certified Pre-Owned: Abusing AD CS'; Url = 'https://specterops.io/blog/2021/06/17/certified-pre-owned/' }
            @{ Titel = 'Microsoft Learn - What is Active Directory Certificate Services?'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview' }
        )
    }
    'esc1' = @{
        Titel = 'ESC1 (Enrollee Supplies Subject + Auth-EKU)'; Schwere = 'Kritisch'
        Zweck = 'Findet veröffentlichte Zertifikatvorlagen, die alle ESC1-Bedingungen erfüllen: "Antragsteller liefert Subject" (ENROLLEE_SUPPLIES_SUBJECT), eine Authentifizierungs-EKU, keine Manager-Genehmigung, keine geforderten Signaturen - und Enroll-Recht für niedrig privilegierte Prinzipale.'
        Hintergrund = 'Bei ESC1 darf der Antragsteller bei einer Vorlage selbst bestimmen, für wen das Zertifikat gilt. Erlaubt die Vorlage zusätzlich die Anmeldung, verlangt keine Genehmigung und ist für normale Benutzer beantragbar, kann sich ein berechtigter Benutzer ein Zertifikat auf "Administrator" ausstellen und sich damit anmelden - eine vollständige Übernahme.'
        Beispiel = 'Ein Mitglied von "Domänen-Benutzer" fordert ein Zertifikat dieser Vorlage mit SAN=Administrator an und meldet sich anschließend als Domain Admin an.'
        Empfehlung = 'ENROLLEE_SUPPLIES_SUBJECT entfernen oder Manager-Genehmigung erzwingen; Enroll-Rechte auf benötigte, nicht-breite Gruppen einschränken; Authentifizierungs-EKUs nur dort belassen, wo nötig.'
        Quellen = @(
            @{ Titel = 'SpecterOps - Certified Pre-Owned (ESC1)'; Url = 'https://specterops.io/blog/2021/06/17/certified-pre-owned/' }
            @{ Titel = 'Microsoft Learn - What is Active Directory Certificate Services?'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview' }
        )
    }
    'esc2_3' = @{
        Titel = 'ESC2/ESC3 (Any Purpose / Enrollment Agent)'; Schwere = 'Hoch'
        Zweck = 'Findet veröffentlichte Vorlagen mit Any-Purpose-EKU bzw. ohne EKU (ESC2) oder mit der Enrollment-Agent-EKU (ESC3), die von niedrig privilegierten Prinzipalen angefordert werden dürfen.'
        Hintergrund = 'ESC2: Vorlagen, deren Zertifikate für nahezu beliebige Zwecke (inklusive Anmeldung) nutzbar sind. ESC3: Vorlagen mit der Berechtigung "Enrollment Agent", mit der sich Zertifikate im Namen anderer beantragen lassen. Beides kann missbraucht werden, um die Identität z. B. eines Administrators zu übernehmen.'
        Beispiel = 'Mit einem Enrollment-Agent-Zertifikat (ESC3) beantragt der Angreifer ein Smartcard-Logon-Zertifikat im Namen eines Administrators.'
        Empfehlung = 'Any-Purpose-/EKU-lose Vorlagen vermeiden; Enrollment-Agent-Vorlagen streng auf wenige, dedizierte Konten beschränken; Enroll-Rechte einschränken.'
        Quellen = @(
            @{ Titel = 'SpecterOps - Certified Pre-Owned (ESC2/ESC3)'; Url = 'https://specterops.io/blog/2021/06/17/certified-pre-owned/' }
            @{ Titel = 'Microsoft Learn - What is Active Directory Certificate Services?'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview' }
        )
    }
    'esc4' = @{
        Titel = 'ESC4 (manipulierbare Vorlagen-ACL)'; Schwere = 'Hoch'
        Zweck = 'Findet Zertifikatvorlagen, deren ACL niedrig privilegierten Prinzipalen Schreibrechte (GenericAll/GenericWrite/WriteDacl/WriteOwner/WriteProperty) gewährt.'
        Hintergrund = 'Wer eine Zertifikatvorlage bearbeiten darf, kann sie selbst so umstellen, dass sie wie ESC1 angreifbar wird, und sich anschließend ein Admin-Zertifikat ausstellen. Schreibrechte auf Vorlagen für breite Gruppen (z. B. Authentifizierte Benutzer, Domänen-Benutzer) sind daher praktisch gleichbedeutend mit einer Übernahme-Möglichkeit.'
        Beispiel = 'Authenticated Users hat WriteProperty auf einer Vorlage - der Angreifer macht sie zu ESC1 und übernimmt die Domäne.'
        Empfehlung = 'Schreibrechte auf Zertifikatvorlagen ausschließlich PKI-Administratoren gewähren; breite Gruppen entfernen.'
        Quellen = @(
            @{ Titel = 'SpecterOps - Certified Pre-Owned (ESC4)'; Url = 'https://specterops.io/blog/2021/06/17/certified-pre-owned/' }
            @{ Titel = 'Microsoft Learn - What is Active Directory Certificate Services?'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview' }
        )
    }
    'esc6' = @{
        Titel = 'ESC6 (EDITF_ATTRIBUTESUBJECTALTNAME2)'; Schwere = 'Hoch'
        Zweck = 'Prüft je Zertifizierungsstelle, ob die gefährliche CA-Einstellung EDITF_ATTRIBUTESUBJECTALTNAME2 aktiv ist (Abfrage über certutil).'
        Hintergrund = 'Ist diese Einstellung an der Zertifizierungsstelle aktiv, darf bei jeder Vorlage jeder Antragsteller frei angeben, für wen das Zertifikat gelten soll. Damit wird praktisch jede anmeldefähige Vorlage angreifbar (wie ESC1) - ein Benutzer kann sich ein Zertifikat auf einen Administrator ausstellen. Es handelt sich um eine die gesamte CA betreffende Fehlkonfiguration.'
        Beispiel = 'Bei gesetztem Flag beantragt ein Benutzer ein gewöhnliches Zertifikat, schmuggelt aber SAN=Administrator hinein und meldet sich als Admin an.'
        Empfehlung = 'EDITF_ATTRIBUTESUBJECTALTNAME2 auf allen CAs deaktivieren (certutil -setreg policy\EditFlags -EDITF_ATTRIBUTESUBJECTALTNAME2; danach den Zertifikatdienst neu starten).'
        Quellen = @(
            @{ Titel = 'SpecterOps - Certified Pre-Owned (ESC6)'; Url = 'https://specterops.io/blog/2021/06/17/certified-pre-owned/' }
            @{ Titel = 'Microsoft Learn - What is Active Directory Certificate Services?'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview' }
        )
    }
    'esc8' = @{
        Titel = 'ESC8 (HTTP Web Enrollment / NTLM-Relay)'; Schwere = 'Mittel'
        Zweck = 'Prüft je CA-Host, ob die Rolle "Web Enrollment" (HTTP-Antragsschnittstelle certsrv) installiert ist.'
        Hintergrund = 'Die Web-Anmeldeschnittstelle der Zertifizierungsstelle akzeptiert Anmeldungen über HTTP. Ein Angreifer kann einen Domänencontroller dazu zwingen, sich dort anzumelden, diese Anmeldung weiterleiten ("Relay") und in dessen Namen ein Zertifikat ausstellen lassen - ein Weg zur Übernahme des Domänencontrollers bzw. der Domäne.'
        Beispiel = 'Per PetitPotam wird ein DC zur NTLM-Authentifizierung gezwungen; diese wird an http://CA/certsrv relayt und ein DC-Zertifikat ausgestellt.'
        Empfehlung = 'Web-Enrollment nur wenn nötig betreiben; HTTPS mit Extended Protection (Channel Binding) erzwingen und NTLM deaktivieren; CA-/RPC-Endpunkte gegen Relay härten (siehe Microsoft ADV210003).'
        Quellen = @(
            @{ Titel = 'SpecterOps - Certified Pre-Owned (ESC8)'; Url = 'https://specterops.io/blog/2021/06/17/certified-pre-owned/' }
            @{ Titel = 'Microsoft Learn - What is Active Directory Certificate Services?'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-cs/active-directory-certificate-services-overview' }
        )
    }
    'gpo_sysvol' = @{
        Titel = 'GPO & SYSVOL - Geheimnisse'; Schwere = 'Hoch'
        Zweck = 'Durchsucht SYSVOL und die GPO-Objekte nach hinterlegten Geheimnissen und gefährlichen Bearbeitungsrechten: GPP-Passwörter (cpassword), Klartext-Anmeldedaten in Skripten und breite Schreibrechte auf GPOs.'
        Hintergrund = 'SYSVOL ist eine für alle Domänenbenutzer lesbare Freigabe. Wer dort Passwörter oder Anmeldedaten (z. B. in Anmeldeskripten) ablegt, gibt sie faktisch jedem Benutzer preis. Außerdem entscheidet die Berechtigung auf den Gruppenrichtlinien, wer sie ändern darf - ein zu breit vergebenes Schreibrecht erlaubt das Ausrollen von Schadcode auf alle betroffenen Systeme.'
        Beispiel = 'Ein einziges cpassword in einer SYSVOL-XML genügt: Der AES-Schlüssel ist seit 2014 öffentlich, jeder Benutzer kann das Passwort entschlüsseln.'
        Empfehlung = 'GPP-Passwörter entfernen (MS14-025), Klartext-Credentials aus Skripten verbannen (gMSA/LAPS), GPO-Bearbeitungsrechte auf wenige Administratoren beschränken.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1552.006 - Group Policy Preferences'; Url = 'https://attack.mitre.org/techniques/T1552/006/' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'gpp_cpassword' = @{
        Titel = 'GPP-Passwörter (cpassword in SYSVOL)'; Schwere = 'Kritisch'
        Zweck = 'Durchsucht die GPP-XML-Dateien im SYSVOL (Groups.xml, Services.xml, ScheduledTasks.xml, DataSources.xml, Printers.xml, Drives.xml) nach cpassword-Werten und entschlüsselt sie zur Demonstration.'
        Hintergrund = 'Über Gruppenrichtlinien konnten früher Passwörter (z. B. für lokale Administratoren oder Dienstkonten) verteilt werden. Sie liegen verschlüsselt in SYSVOL-Dateien - doch der zugehörige Schlüssel wurde 2014 von Microsoft öffentlich gemacht (MS14-025). Da jeder Domänenbenutzer SYSVOL lesen kann, ist ein solches Passwort trivial zu entschlüsseln. Der Patch verhindert nur das Anlegen neuer solcher Passwörter; bereits vorhandene müssen manuell entfernt werden.'
        Beispiel = 'Ein cpassword in Groups.xml setzt das lokale Administrator-Passwort auf allen Clients - der Angreifer entschlüsselt es und hat überall lokalen Admin.'
        Empfehlung = 'Alle cpassword-Vorkommen aus SYSVOL entfernen und die betroffenen Passwörter rotieren; für lokale Admin-Passwörter LAPS, für Dienstkonten gMSA verwenden.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1552.006 - Group Policy Preferences'; Url = 'https://attack.mitre.org/techniques/T1552/006/' }
            @{ Titel = 'Microsoft Security Bulletin MS14-025 (CVE-2014-1812)'; Url = 'https://learn.microsoft.com/en-us/security-updates/securitybulletins/2014/ms14-025' }
        )
    }
    'sysvol_scripts' = @{
        Titel = 'Klartext-Credentials in SYSVOL-Skripten'; Schwere = 'Hoch'
        Zweck = 'Durchsucht Skriptdateien im SYSVOL (.bat/.cmd/.ps1/.vbs/.kix) nach Mustern, die auf hinterlegte Klartext-Anmeldedaten hindeuten (net use /user:, password=, ConvertTo-SecureString -AsPlainText u. a.).'
        Hintergrund = 'Logon-, Startup- und Wartungsskripte werden häufig in SYSVOL abgelegt und sind für alle Domänenbenutzer lesbar. Enthalten sie hartkodierte Anmeldedaten (z. B. für Laufwerks-Mappings, geplante Aufgaben oder Tool-Aufrufe), kann jeder Benutzer sie auslesen. Die Prüfung ist heuristisch (muster-basiert) und kann False Positives liefern - die Treffer sind manuell zu bewerten.'
        Beispiel = 'Ein logon.bat mit "net use Z: \\srv\share /user:svc-backup P@ssw0rd" verrät das Dienstkonto-Passwort an jeden Benutzer.'
        Empfehlung = 'Klartext-Anmeldedaten aus Skripten entfernen; stattdessen gMSA, LAPS oder einen Credential-Manager verwenden; betroffene Passwörter rotieren.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1552.001 - Credentials In Files'; Url = 'https://attack.mitre.org/techniques/T1552/001/' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'gpo_rights' = @{
        Titel = 'GPO-Bearbeitungsrechte'; Schwere = 'Hoch'
        Zweck = 'Prüft die ACLs der GPO-Objekte auf Schreibrechte (GenericAll/GenericWrite/WriteDacl/WriteOwner/WriteProperty) für niedrig privilegierte/breite Prinzipale.'
        Hintergrund = 'Wer eine verknüpfte Gruppenrichtlinie bearbeiten darf, kann auf allen davon betroffenen Systemen Einstellungen, Skripte oder Aufgaben ausrollen - bis hin zur Ausführung von Schadcode mit höchsten Rechten. Schreibrechte auf Gruppenrichtlinien für breite Gruppen (z. B. Authentifizierte Benutzer) sind daher hochkritisch, besonders bei Richtlinien, die auf Domänencontroller oder Server wirken.'
        Beispiel = 'Hat "Authenticated Users" GenericWrite auf der Default Domain Policy, kann jeder Benutzer ein Startup-Skript einschleusen, das auf allen Systemen als SYSTEM läuft.'
        Empfehlung = 'GPO-Bearbeitungsrechte ausschließlich dedizierten GPO-Administratoren gewähren; breite Gruppen entfernen; Änderungen überwachen.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1484.001 - Group Policy Modification'; Url = 'https://attack.mitre.org/techniques/T1484/001/' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'dc_haertung' = @{
        Titel = 'DC-Härtung (vertieft)'; Schwere = 'Hoch'
        Zweck = 'Vertieft die DC-Härtung über die Basis-Checks hinaus: LDAP-Signing & Channel Binding, erzwungenes SMB-Signing, Print Spooler auf DCs und anonyme LDAP-Binds.'
        Hintergrund = 'Domänencontroller sind die schützenswertesten Systeme der Domäne. Mehrere verbreitete Angriffe (Weiterleiten abgefangener Anmeldungen, erzwungene Authentifizierung über den Druckdienst) lassen sich durch wenige Härtungseinstellungen entschärfen. Diese Teilprüfungen lesen die relevanten Einstellungen und Dienste je Domänencontroller read-only aus.'
        Beispiel = 'Ist LDAP-Signing nicht erzwungen und der Print Spooler aktiv, kann ein Angreifer einen DC via PrinterBug zur NTLM-Authentifizierung zwingen und diese an einen anderen DC relayen.'
        Empfehlung = 'LDAP-Signing + Channel Binding erzwingen, SMB-Signing erforderlich setzen, Print Spooler auf DCs deaktivieren, anonyme LDAP-Binds unterbinden.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Windows security baselines'; Url = 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/device-management/windows-security-configuration-framework/windows-security-baselines' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'ldap_signing' = @{
        Titel = 'LDAP-Signing & Channel Binding'; Schwere = 'Hoch'
        Zweck = 'Liest je DC die Registry-Werte LDAPServerIntegrity (LDAP-Signing) und LdapEnforceChannelBinding (Channel Binding) aus.'
        Hintergrund = 'Ohne erzwungenes LDAP-Signing nimmt der Domänencontroller Verzeichnisanfragen ohne ausreichenden Schutz entgegen - das ermöglicht Abhören und das Weiterleiten abgefangener Anmeldungen ("Relay"). Channel Binding bindet die verschlüsselte Anmeldung zusätzlich an die jeweilige Verbindung und unterbindet dieses Weiterleiten. Beides sollte erzwungen sein.'
        Beispiel = 'Ein Angreifer relayt die NTLM-Authentifizierung eines Computerkontos an den LDAP-Dienst eines DCs und trägt z. B. ein RBCD-Recht ein - ohne LDAP-Signing/CBT gelingt das.'
        Empfehlung = 'LDAPServerIntegrity auf 2 (erforderlich) und LdapEnforceChannelBinding auf 2 (immer) setzen; vorab per Event 2887 Clients identifizieren, die ungesignt binden.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - How to enable LDAP signing'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/enable-ldap-signing-in-windows-server' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'smb_signing' = @{
        Titel = 'SMB-Signing (erforderlich)'; Schwere = 'Mittel'
        Zweck = 'Liest je DC den Registry-Wert RequireSecuritySignature (LanManServer) - ist SMB-Signing serverseitig erforderlich?'
        Hintergrund = 'SMB-Signing versieht jede Datei-/Netzwerkanfrage mit einer Signatur und verhindert so Manipulation und das Weiterleiten abgefangener Anmeldungen. Für die Freigaben SYSVOL/NETLOGON erzwingen Domänencontroller das bereits; ist Signing darüber hinaus nicht generell erforderlich, bleiben andere Verbindungen ungeschützt und angreifbar.'
        Beispiel = 'Bei nicht erzwungenem SMB-Signing kann ein Angreifer eine SMB-Authentifizierung abfangen und an einen anderen Dienst weiterleiten.'
        Empfehlung = 'RequireSecuritySignature serverseitig auf 1 setzen (GPO: "Microsoft network server: Digitally sign communications (always)").'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Overview of SMB signing'; Url = 'https://learn.microsoft.com/en-us/windows-server/storage/file-server/smb-signing-overview' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'print_spooler' = @{
        Titel = 'Print Spooler auf DCs'; Schwere = 'Hoch'
        Zweck = 'Prüft je DC, ob der Druckwarteschlangen-Dienst (Print Spooler) läuft.'
        Hintergrund = 'Über den Druckwarteschlangen-Dienst lässt sich ein Domänencontroller dazu bringen, sich gegen ein anderes System zu authentifizieren ("PrinterBug"). Diese Anmeldung kann ein Angreifer weiterleiten (etwa an die Verzeichnis- oder Zertifikatsdienste) und so den Domänencontroller bzw. die Domäne übernehmen. Auf Domänencontrollern wird der Dienst praktisch nie benötigt.'
        Beispiel = 'Per PrinterBug zwingt der Angreifer DC-A, sich gegen seinen Relay-Server zu authentifizieren, und leitet diese Authentifizierung an die AD-CS-Web-Enrollment-Schnittstelle weiter (ESC8).'
        Empfehlung = 'Print Spooler auf allen DCs deaktivieren (Set-Service Spooler -StartupType Disabled; Stop-Service Spooler), sofern nicht zwingend benötigt.'
        Quellen = @(
            @{ Titel = 'MITRE ATT&CK T1187 - Forced Authentication'; Url = 'https://attack.mitre.org/techniques/T1187/' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'anon_ldap' = @{
        Titel = 'Anonyme LDAP-Binds (dSHeuristics)'; Schwere = 'Mittel'
        Zweck = 'Liest das dSHeuristics-Attribut aus und prüft, ob anonyme LDAP-Operationen erlaubt sind (7. Zeichen = 2).'
        Hintergrund = 'Eine forestweite Verzeichnis-Einstellung (dSHeuristics) kann anonyme, nicht authentifizierte LDAP-Zugriffe erlauben. Ist das aktiviert, kann ein Angreifer ohne jedes Konto Benutzer-, Gruppen- und Konfigurationsdaten auslesen (Aufklärung). Standardmäßig ist das nicht gesetzt.'
        Beispiel = 'Bei erlaubten anonymen Binds enumeriert ein Angreifer aus dem Netz ohne gültiges Konto die gesamte Benutzerliste.'
        Empfehlung = 'Anonyme LDAP-Binds nicht erlauben (7. Zeichen von dSHeuristics nicht auf 2 setzen); falls für Altanwendungen nötig, eng begrenzen und überwachen.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - How to enable LDAP signing'; Url = 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/enable-ldap-signing-in-windows-server' }
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
        )
    }
    'delta' = @{
        Titel = 'Veränderungen seit letztem Lauf (Delta)'; Schwere = 'Info'
        Zweck = 'Vergleicht den aktuellen Lauf mit einem früheren JSON-Export und stellt gegenüber, welche Befunde neu hinzugekommen und welche behoben sind (Grundlage sind die rot/gelb markierten Einträge beider Läufe).'
        Hintergrund = 'Ein Assessment ist vor allem im Zeitverlauf aussagekräftig: Eine Momentaufnahme zeigt den Zustand, erst der Vergleich zweier Läufe zeigt Fortschritt (behobene Punkte) und Rückschritt (neue Funde). Der Delta-Modus liest dazu über -Vergleich einen früheren Export ein und stellt neue und behobene Befunde gegenüber. Der Abgleich erfolgt rein lokal über die Exporte, ohne zusätzliche AD-Abfragen - ideal, um Fortschritt gegenüber Audit/Management zu belegen.'
        Beispiel = 'Nach dem Entfernen der Enroll-Berechtigung an einer verwundbaren Zertifikatsvorlage erscheint der ESC1-Befund im nächsten Lauf unter "Behoben". Taucht hingegen ein neuer Domänen-Admin oder eine neue kerberoastbare Dienst-SPN auf, listet ihn der Bereich unter "Neu".'
        Empfehlung = 'Den JSON-Export jedes Assessments revisionssicher aufbewahren (Datum/DC im Dateinamen) und beim nächsten Lauf via -Vergleich referenzieren. Neue Befunde priorisiert prüfen; behobene als Nachweis dokumentieren. Den Vergleich möglichst gegen denselben DC bzw. dieselbe Domäne fahren, damit die Mengen vergleichbar sind.'
        Quellen = @(
            @{ Titel = 'Microsoft Learn - Best practices for securing Active Directory'; Url = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory' }
            @{ Titel = 'Microsoft Learn - Security baselines (Windows / Security Compliance Toolkit)'; Url = 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/device-management/windows-security-configuration-framework/windows-security-baselines' }
        )
    }
}
function Doku ($id) {
    # Emittiert den Katalogeintrag $id als 'Doku'-Ereignis (Begruendung im HTML/JSON).
    $k = $CheckKatalog[$id]
    if ($null -eq $k) { return }
    Merken 'Doku' @{
        CheckId     = $id
        DTitel      = $k.Titel
        Schwere     = $k.Schwere
        Zweck       = $k.Zweck
        Hintergrund = $k.Hintergrund   # optional: Technik/Protokoll/Schwachstelle
        Beispiel    = $k.Beispiel
        Empfehlung  = $k.Empfehlung
        Quellen     = $k.Quellen        # String ODER Liste @{ Titel; Url } (klickbare Links)
    }
}
####################################################################################################
# Design Funktionen (Header, Vollzeile, Leerzeile, Buttom, Bereichstitel, Subtitel)                #
####################################################################################################
function Header {                                                                                  #
    Merken 'Kopf' @{ Typ = $type; Titel = $maintitel; System = $system; Datum = $datum             #
                     Firma = $firma }                                                              #
    $header = New-Object 'object[,]' 5,14                                                          #
    $frame = New-Object 'object[]' 5                                                               #
    ######### Mit default Farbe fuellen ############################################################
    $header[0,1] = $header[0,3] = $header[0,5] = $header[0,7] = $header[0,9] = $header[0,11] `
    = $header[0,13] = "white"                                                                      #
    $header[1,1] = $header[1,3] = $header[1,5] = $header[1,7] = $header[1,9] = $header[1,11] `
    = $header[1,13] = "white"                                                                      #
    $header[2,1] = $header[2,3] = $header[2,5] = $header[2,7] = $header[2,9] = $header[2,11] `
    = $header[2,13] = "white"                                                                      #
    $header[3,1] = $header[3,3] = $header[3,5] = $header[3,7] = $header[3,9] = $header[3,11] `
    = $header[3,13] = "white"                                                                      #
    $header[4,1] = $header[4,3] = $header[4,5] = $header[4,7] = $header[4,9] = $header[4,11] `
    = $header[4,13] = "white"                                                                      #
    ######### Vertikalen Rahmen erstellen ##########################################################
    for ($a=0;$a -lt 5;$a++) {                                                                     #
    $header[$a,0] = $header[$a,4] = $header[$a,8] =$header[$a,12] = $zeichen                       #
    $header[$a,1] = $header[$a,5] = $header[$a,9] =$header[$a,13] = $F_Rahmen                      #
    }                                                                                              #
    ######### Horizontalen Rahmen erstellen ########################################################
    $header[0,2] = $header[0,10] = $header[2,2] = $header[2,10] = $header[4,2] = $header[4,10] `
    = $zeichen * 18                                                                                #
    $header[0,3] = $header[0,11] = $header[2,3] = $header[2,11] = $header[4,3] = $header[4,11] `
    = $F_Rahmen                                                                                    #
    $header[1,2] = $header[1,10] = $header[3,2] = $header[3,10] = $leer * 18                       #
    ######### Mitte aufbauen #######################################################################
    $leere_mitte = $sb - 40                                                                        #
    $header[0,6] = $header[4,6] = $zeichen * $leere_mitte                                          #
    $header[0,7] = $header[4,7] = $F_Rahmen                                                        #
    $header[1,6] = $header[2,6] = $header[3,6] = $leer * $leere_mitte                              #
    ######### Oben links ###########################################################################
    if($type.Length % 2 -ne 0) { $dif_type = $type.Length + 1 ; $type = "$type$leer" }             #
    else { $dif_type = $type.Length }                                                              #
    if($type.Length -gt 16) { $typ = $type.Substring(0,15) +"~" ; $header[1,2] = "$leer$typ$leer" }# 
    elseif($type.Length -eq 16) { $header[1,2] = "$leer$type$leer" }                               #
    else { $dif = (16 - $dif_type) / 2 ; $tmp = $leer * $dif                                       #
        $header[1,2] = "$leer$tmp$type$tmp$leer" }                                                 #
    $header[1,3] = $F_Ue_Schrift                                                                   #
    ######### Oben rechts ##########################################################################
    if($datum.Length % 2 -ne 0) { $dif_datum = $datum.Length + 1 ; $datus = "$datum$leer" }        #
    else { $dif_datum = $datum.Length ; $datus = "$datum" }                                        #
    if($datus.Length -gt 16)                                                                       #
    { $dat = $datum.Substring(0,15) +"~" ; $header[1,10] = "$leer$dat$leer" }                      #
    elseif ($datus.Length -eq 16) { $header[1,10] = "$leer$datum$leer" }                           #
    else { $dif = (16 - $dif_datum) / 2 ; $tmp = $leer * $dif                                      #
        $header[1,10] = "$leer$tmp$datus$tmp$leer" }                                               #
    $header[1,11] = $F_Ue_Schrift                                                                  #
    ######### Unten links ##########################################################################
    if($system.Length % 2 -ne 0) { $dif_sys = $system.Length + 1 ; $sys = "$system$leer" }         #
    else { $dif_sys = $system.Length ; $sys = $system }                                            #
    if($sys.Length -gt 16) { $sys = $system.Substring(0,15) +"~" ; $header[3,2] = "$leer$sys$leer"}#
    elseif ($sys.Length -eq 16) { $header[3,2] = "$leer$sys$leer" }                                #
    else { $dif = 16 - $dif_sys ; $tmp = $leer * $dif                                              #
        $header[3,2] = "$leer$sys$tmp$leer" }                                                      #
    $header[3,3] = $F_Ue_Schrift                                                                   #
    ######### Unten rechts #########################################################################
    if($firma.Length % 2 -ne 0) { $dif_firm = $firma.Length + 1 ; $firm = "$firma$leer" }          #
    else { $dif_firm = $firma.Length ; $firm = $firma }                                            #
    if($firm.Length -gt 16) { $firm = $firm.Substring(0,15) +"~"                                   #
        $header[3,10] = "$leer$firm$leer" }                                                        #
    elseif ($firm.Length -eq 16) { $header[3,10] = "$leer$firm$leer" }                             #
    else { $dif = 16 - $dif_firm ; $tmp = $leer * $dif                                             #
        $header[3,10] = "$leer$tmp$firm$leer" }                                                    #
    $header[3,11] = $F_Ue_Schrift                                                                  #
    ######### Zentral ##############################################################################
    if($maintitel.Length % 2 -ne 0) { $dif_mi = $maintitel.Length + 1 ; $main = "$maintitel$leer" }#
    else { $dif_mi = $maintitel.Length ; $main = "$maintitel" }                                    #
    if($dif_mi  -gt $leere_mitte) { $main = $maintitel.Substring(0,15) +"~"                        #
        $header[2,6] = "$leer$main$leer" }                                                         #
    elseif ($dif_mi -eq $leere_mitte) { $header[2,6] = "$leer$main$leer" }                         #
    else { $dif = ($leere_mitte - $dif_mi - 2) / 2 ; $tmp = $leer * $dif                           #
        $header[2,6] = "$leer$tmp$main$tmp$leer" }                                                 #            
    $header[2,7] = $F_Ue_Schrift                                                                   #
    ######### Zeilen fuer die Dateiausgabe zusammensetzen (unabhaengig von der Konsole) ###########
    for ($z=0;$z -lt 5;$z++) {                                                                     #
        $frame[$z]= $header[$z,0] + $header[$z,2] + $header[$z,4] + $header[$z,6] + $header[$z,8] `
        + $header[$z,10] + $header[$z,12]                                                          #
    }                                                                                              #
    ######### Ausgabe Console ######################################################################
    if ($A_Con -eq 1) {                                                                            #
        for ($z=0;$z -lt 5;$z++) {                                                                 #
        Write-Host $header[$z,0] -ForegroundColor $header[$z,1] -NoNewline                         #
        Write-Host $header[$z,2] -ForegroundColor $header[$z,3] -NoNewline                         #
        Write-Host $header[$z,4] -ForegroundColor $header[$z,5] -NoNewline                         #
        Write-Host $header[$z,6] -ForegroundColor $header[$z,7] -NoNewline                         #
        Write-Host $header[$z,8] -ForegroundColor $header[$z,9] -NoNewline                         #
        Write-Host $header[$z,10] -ForegroundColor $header[$z,11] -NoNewline                       #
        Write-Host $header[$z,12] -ForegroundColor $header[$z,13]                                  #
        }                                                                                          #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        for ($z=0;$z -lt 5;$z++) { Ausgabe $frame[$z] }                                            #
    }                                                                                              #
}                                                                                                  #
function Bottom {                                                                                  #
    ### Legende ####################################################################################
    #                                                                                              #
    ### Berechnung #################################################################################
    $zeile1u3 = "$zeichen" * $sb                                                                   #
    $zeile2b = $madeby                                                                             #
    if ($madeby.Length -gt "20") { $zeile2b = $leer + $madeby.Substring(0,19) + "~" + $leer }      #
    else { $zeile2b = $leer + $madeby.PadRight(20) + $leer }                                       #
    $dif = $sb - $zeile2b.Length                                                                   #
    if ($dif % 2 -ne 0) { $dif = $dif +1 }                                                         #
    $dif = $dif / 2                                                                                #
    $zeile2a = "$zeichen" * $dif                                                                   #
    $zeile2d = "$zeichen" * $dif                                                                   #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile1u3" -ForegroundColor $F_Rahmen                                          #
        Write-Host "$zeile2a" -NoNewline -ForegroundColor $F_Rahmen                                #
        Write-Host "$zeile2b" -NoNewline -ForegroundColor $F_Ue_Schrift                            #
        Write-Host "$zeile2c" -NoNewline -ForegroundColor $F_Rahmen                                #
        Write-Host "$zeile2d" -ForegroundColor $F_Rahmen                                           #
        Write-Host "$zeile1u3" -ForegroundColor $F_Rahmen                                          #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe $zeile1u3                                                                          #
        Ausgabe "$zeile2a$zeile2b$zeile2c$zeile2d"                                                 #
        Ausgabe $zeile1u3                                                                          #
    }                                                                                              #
}                                                                                                  #
function Vollzeile {                                                                               #
    ### Legende ####################################################################################
    #                                                                                              #
    ### Berechnung #################################################################################
    for ($a=0;$a -lt $sb;$a++) { $zeile = "$zeichen$zeile"}                                        #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile" -ForegroundColor $F_Rahmen                                             #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe $zeile                                                                             #
    }                                                                                              #
}                                                                                                  #
function Leerzeile {                                                                               #
    ### Legende ####################################################################################
    #                                                                                              #
    ### Berechnung #################################################################################
    $links = $rechts = "$zeichen" ; $dif_leer = $sb - $links.Length - $rechts.Length               #
    for ($a=0;$a -lt $dif_leer;$a++) { $zeile = "$leer$zeile"}                                     #
    $zeile = "$links$zeile$rechts"                                                                 #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile" -ForegroundColor $F_Rahmen                                             #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe $zeile                                                                             #
    }                                                                                              #
}                                                                                                  #
function Trennzeile ($sub) {                                                                       #
    ### Legende ####################################################################################
    # $sub = Einsatzbereich als Sublinie ja(s)                                                     #
    ### Berechnung #################################################################################
    if ($sub) { $links = "$zeichen$leer$leer" } else { $links = "$zeichen$leer" }                  #
    $rechts = "$leer$zeichen"                                                                      #
    $dif = $sb - $links.Length - $rechts.Length                                                    #
    $zeile = "$zeichen" * $dif                                                                     #
    $zeile = "$links$zeile$rechts"                                                                 #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile" -ForegroundColor $F_Rahmen                                             #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe $zeile                                                                             #
    }                                                                                              #
}                                                                                                  #                                                                                                #
function tablinie ($sub) {                                                                         #
    ### Legende ####################################################################################
    # $sub = Einsatzbereich als Subtabelle? nein(n), ja(s)                                         #
    ### Berechnung #################################################################################
    $rechts = "$leer$zeichen"                                                                      #
    if ($sub -eq 's') { $links = "$zeichen$leer$leer" } else { $links = "$zeichen$leer" }          #
    $dif_leer = $sb - $links.Length - $rechts.Length                                               #
    for ($a=0;$a -lt $dif_leer;$a++) { $zeile = "$tabzeichen$zeile"}                               #
    $zeile = "$links$zeile$rechts"                                                                 #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile" -ForegroundColor $F_Rahmen                                             #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe $zeile                                                                             #
    }                                                                                              #
}                                                                                                  #
function Bereich ($Wert1) {                                                                        #
    ### Legende ####################################################################################
    # Wert1  = Dieser Wert wird im Bereichsheader mittig plaziert                                  #
    ### Berechnung #################################################################################
    Merken 'Bereich' @{ Titel = $Wert1 }                                                           #
    for ($a=0;$a -lt $sb;$a++) { $zeile1u3 = "$zeichen$zeile1u3"}                                  #
    $zeile2a = "$zeichen$zeichen$zeichen$leer" ; $zeile2d = "$leer$zeichen$zeichen$zeichen"        #
    $zeile2b = "$wert1" ; $dif = $sb - $zeile2a.Length - $zeile2d.Length -$Wert1.Length            #
    for ($b=0;$b -lt $dif;$b++) { $zeile2c = " $zeile2c" }                                         #
    $zeile2c = "$zeile2c$zeile2d"                                                                  #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile1u3" -ForegroundColor $F_Rahmen                                          #
        Write-Host "$zeile2a" -NoNewline -ForegroundColor $F_Rahmen                                #
        Write-Host "$zeile2b" -NoNewline -ForegroundColor $F_Ue_Schrift                            #
        Write-Host "$zeile2c" -ForegroundColor $F_Rahmen                                           #
        Write-Host "$zeile1u3" -ForegroundColor $F_Rahmen                                          #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe $zeile1u3                                                                          #
        Ausgabe "$zeile2a$zeile2b$zeile2c"                                                         #
        Ausgabe $zeile1u3                                                                          #
    }                                                                                              #
}                                                                                                  #
function Bereichstitel ($Wert1,$sub) {                                                             #
    ### Legende ####################################################################################
    # Wert1  = Dieser Wert wird als Bereichstitel links plaziert                                   #
    # $sub = Einsatzbereich als Subtabelle? nein(n), ja(s)                                         #
    ### Berechnung #################################################################################
    Merken 'Titel' @{ Text = $Wert1 }                                                              #
    $zeile1u2d = "$leer$zeichen" ; $zeile1b = $Wert1                                               #
    if ($sub -eq 's') { $zeile1u2a = "$zeichen$leer$leer" } else { $zeile1u2a = "$zeichen$leer" }  #
    $dif1 = $sb - $zeile1u2a.Length - $zeile1b.Length - $zeile1u2d.Length                          #
    for ($a=0;$a -lt $dif1;$a++) { $zeile1c = "$leer$zeile1c"}                                     #
    $zeile1u2c = "$zeile1c$zeile1u2d"                                                              #
    for ($b=0;$b -lt $zeile1b.Length;$b++) { $zeile2b = "$zeichen$zeile2b" }                       #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile1u2a" -NoNewline -ForegroundColor $F_Rahmen                              #
        Write-Host "$zeile1b" -NoNewline -ForegroundColor $F_Ue_Schrift                            #
        Write-Host "$zeile1u2c" -ForegroundColor $F_Rahmen                                         #
        Write-Host "$zeile1u2a" -NoNewline -ForegroundColor $F_Rahmen                              #
        Write-Host "$zeile2b" -NoNewline -ForegroundColor $F_Rahmen                                #
        Write-Host "$zeile1u2c" -ForegroundColor $F_Rahmen                                         #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe "$zeile1u2a$zeile1b$zeile1u2c"                                                     #
        Ausgabe "$zeile1u2a$zeile2b$zeile1u2c"                                                     #
    }                                                                                              #
}                                                                                                  #
function Subtitel ($wert,$ein,$uz) {
    ### Legende ####################################################################################
    # Wert = Dieser Wert wird als Subtitel links plaziert                                          #
    # $ein = wie weitet der Titel eingerückt werden soll                                           #
    # $uz  = Welches Zeichen zum Unterstreichen verwendet werden soll (Std. "*")                   #
    ### Berechnung #################################################################################
    Merken 'Subtitel' @{ Text = $wert }                                                            #
    if (!($uz)) { $uz = "*" } ; if(!($ein)) { $ein = "0" } ; $ein = " " *$ein                      #
    $zei1 = "* " + "$ein" ; $zei4 = " *" ; $zei2 = $wert ; $zei5 = "$uz" *$wert.Length             #
    $dif = $sb - $zei1.Length - $zei2.Length - $zei4.Length ; $zei3 = " " *$dif                    #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zei1" -NoNewline -ForegroundColor $F_Rahmen                                   #
        Write-Host "$zei2" -NoNewline -ForegroundColor $F_Ue_Schrift                               #
        Write-Host "$zei3" -NoNewline                                                              #
        Write-Host "$zei4" -ForegroundColor $F_Rahmen                                              #
        Write-Host "$zei1" -NoNewline -ForegroundColor $F_Rahmen                                   #
        Write-Host "$zei5" -NoNewline -ForegroundColor $F_Rahmen                                   #
        Write-Host "$zei3" -NoNewline                                                              #
        Write-Host "$zei4" -ForegroundColor $F_Rahmen                                              #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe "$zei1$zei2$zei3$zei4"                                                             #
        Ausgabe "$zei1$zei5$zei3$zei4"                                                             #
    }                                                                                              #
}
####################################################################################################
# Ausgabe Funktionen fuer Werte                                                                    #
####################################################################################################
function 2werte ($Wert1,$Wert2,$sub,$farb) {                                                       #
    ### Legende ####################################################################################
    # Wert1  = Dieser Wert ist einer der auszugebenden Werte                                       #
    # Wert2  = Dieser Wert ist einer der auszugebenden Werte                                       #
    # $sub = Einsatzbereich als Subtabelle? nein(n), ja(s)                                         #
    # $farb = Farbe fuer den zweiten Wert                                                          #
    ### Berechnung #################################################################################
    Merken 'Wert' @{ Name = $Wert1; Wert = $Wert2; Farbe = $farb }                                 #
    $zeile_e = "$leer$zeichen"                                                                     #
    if ($sub -eq 's') { $zeile_a = "$zeichen$leer$leer" } else { $zeile_a = "$zeichen$leer" }      #
    $zeile_b = "$Wert1"                                                                            #
    $zeile_c = " $Wert2"                                                                           #
    $dif = $sb - $zeile_a.Length - $zeile_b.Length -$zeile_c.Length - $zeile_e.Length              #
    for ($a=0;$a -lt $dif;$a++) { $zeile_d = "$leer$zeile_d" }                                     #
    if ($null -eq $farb) { $fa = $F_Text} else { $fa = $farb  }                                    #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$zeile_a" -NoNewline -ForegroundColor $F_Rahmen                                #
        Write-Host "$zeile_b" -NoNewline -ForegroundColor $F_Text                                  #
        Write-Host "$zeile_c" -NoNewline -ForegroundColor $fa                                      #
        Write-Host "$zeile_d" -NoNewline                                                           #
        Write-Host "$zeile_e" -ForegroundColor $F_Rahmen                                           #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe "$zeile_a$zeile_b$zeile_c$zeile_d$zeile_e"                                         #
    }                                                                                              #
}                                                                                                  #
function new_2werte ($sub,$tz,$W1b,$W1,$Wf1,$W1p,$W2,$Wf2,$W2p) {                                  #
    ### Legende ####################################################################################
    # $sub = Einsatzbereich als Subtabelle? ja irgend ein Zeichen                                  #
    # $tz  = gewuenschtes Trennzeichen                                                             #
    # $W1b = Breite des ersten Wertes                                                              #
    # $W1  = Dieser Wert ist einer der auszugebenden Werte                                         #
    # $Wf1 = Farbe des ersten Wertes                                                               #
    # $W1p = Position des Wertes (r=rechts, l=links)                                               #
    # $W2  = Dieser Wert ist einer der auszugebenden Werte                                         #
    # $Wf2 = Farbe des zweiten Wertes                                                              #
    # $W2p = Position des Wertes (r=rechts, l=links)                                               #
    ### Berechnung #################################################################################
    Merken 'Wert' @{ Name = $W1; Wert = $W2; Farbe = $Wf2 }                                        #
    if ($sub) { $ze1 = "$zeichen$leer$leer" } else { $ze1 = "$zeichen$leer" }                      #
    $ze5 = "$leer$zeichen"                                                                         #
    if ($tz) { $ze3 = "$tz " } else { $ze3 = " " }                                                 #
    if ($W1b) {                                                                                    #
        if ($W1.Length -gt $W1b) { $ze2 = $W1.Substring(0,$W1b - 1) + "~" } else { $ze2 = $W1 } }  #
    if ($W1p) {                                                                                    #
        if ($W1p -eq "r") { $ze2 = $ze2.PadLeft($W1b) } else { $ze2 = $ze2.PadRight($W1b) } }      #
    $fr1 = $sb - $ze1.Length - $ze2.Length - $ze3.Length - $ze5.Length                             #
    if ($W2.Length -gt $fr1) { $ze4 = $w2.Substring(0,$fr1 - 1) + "~" }                            #
        elseif ($W2.Length -eq $fr1) { $ze4 = $W2 } else {                                         #
            if ($W2p -eq "l") { $ze4 = $W2.PadRight($fr1) }                                        #
            if ($W2p -eq "r") { $ze4 = $W2.PadLeft($fr1) }                                         #
        }                                                                                          #
    if (!($Wf1)) { $Wf1 = $F_Text}                                                                 #
    if (!($Wf2)) { $Wf2 = $F_Text}                                                                 #
    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {                                                                            #
        Write-Host "$ze1" -NoNewline -ForegroundColor $F_Rahmen                                    #
        Write-Host "$ze2" -NoNewline -ForegroundColor $Wf1                                         #
        Write-Host "$ze3" -NoNewline                                                               #
        Write-Host "$ze4" -NoNewline -ForegroundColor $Wf2                                         #
        Write-Host "$ze5" -ForegroundColor $F_Rahmen                                               #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        Ausgabe "$ze1$ze2$ze3$ze4$ze5"                                                             #
    }                                                                                              #
}
####################################################################################################
# Ausgabe Funktionen fuer Tabellen                                                                 #
####################################################################################################
function neu_tab_max6w_fb {                                                                        #
    param (                                                                                        #
        [int]$spa,                       # $spa   = Anzahl der Tabellenspalten                     #
        $pos,                            # $pos   = Tabelle links(l) oder rechts(r)                #  
        $sub,                            # $sub   = Einsatzbereich als Subtabelle? nein(n), ja(s)  #
        [int]$bre,                       # $bre   = Breite der Tabellenspalten                     #
        $wex,                            # $wex   = Text vor bzw. hinter den Tabellenspalten       #
        $we1,$we2,$we3,$we4,$we5,$we6,   # $we(n) = Wert der n.ten Spalte und Farben               #
        $we7,$we8,$we9,$we10,$we11,$we12 # Erst die Werte, dann die Farben !!                      #
    )                                                                                              #
    ### Variablen ##################################################################################
    # Remove-Variable -Name $all_we -Force -ErrorAction Ignore                                     #
    # Remove-Variable -Name $farben -Force -ErrorAction Ignore                                     #
    ################################################################################################
    $alleWe = @($we1,$we2,$we3,$we4,$we5,$we6,$we7,$we8,$we9,$we10,$we11,$we12)                    #
    Merken 'TabZeile' @{ Zellen = @($alleWe[0..($spa-1)]); Farben = @($alleWe[$spa..(2*$spa-1)])   #
                         Extra = $wex }                                                            #
    $trenner = "$leer$tabtrenner$leer"                                                             #
    $zeilen = ($spa*2)+3 ; $ende = "$leer$zeichen"                                                 #
    $Sammeln = @() ; $speicher = @()                                                               #
    for($a=0;$a -lt $zeilen;$a++) {                                                                #
        $Sammeln += New-Object psobject -Property @{W1= $null ;W2= $a}                             #
    }                                                                                              #
    if($sub -eq 's') { $Sammeln[0].W1 = "$zeichen$leer$leer" ; $Sammeln[0].W2 = $F_Rahmen }        #
    else { $Sammeln[0].W1 = "$zeichen$leer" ; $Sammeln[0].W2 = $F_Rahmen }                         #
    for ($b=2;$b -le $zeilen-2;$b=$b+2) { $Sammeln[$b].W1 = $trenner ; $Sammeln[$b].W2 = $F_Rahmen }
    $Sammeln[$zeilen-1].W1 = $ende ; $Sammeln[$zeilen-1].W2 = $F_Rahmen                            #
    switch ($spa) {                                                                                #
    "6" {                                                                                          #
        if($null -ne $we1) { $Wert1 += $we1 } else { $Wert1 += $null }                             #
        if($null -ne $we2) { $Wert2 += $we2 } else { $Wert2 += $null }                             #
        if($null -ne $we3) { $Wert3 += $we3 } else { $Wert3 += $null }                             #
        if($null -ne $we4) { $Wert4 += $we4 } else { $Wert4 += $null }                             #
        if($null -ne $we5) { $Wert5 += $we5 } else { $Wert5 += $null }                             #
        if($null -ne $we6) { $Wert6 += $we6 } else { $Wert6 += $null }                             #
        #######################################################################################    # 
        if($null -eq $we7 -or $we7.Length -eq 0) { $Wert7 += $F_Text } else { $Wert7 += $we7 }     #
        if($null -eq $we8 -or $we8.Length -eq 0) { $Wert8 += $F_Text } else { $Wert8 += $we8 }     #
        if($null -eq $we9 -or $we9.Length -eq 0) { $Wert9 += $F_Text } else { $Wert9 += $we9 }     #
        if($null -eq $we10 -or $we10.Length -eq 0) { $Wert10 += $F_Text } else { $Wert10 += $we10 }#
        if($null -eq $we11 -or $we11.Length -eq 0) { $Wert11 += $F_Text } else { $Wert11 += $we11 }#
        if($null -eq $we12 -or $we12.Length -eq 0) { $Wert12 += $F_Text } else { $Wert12 += $we12 }#
        #######################################################################################    # 
        $speicher += New-Object psobject -Property @{W1="$Wert1";W2="$wert7"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert2";W2="$wert8"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert3";W2="$wert9"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert4";W2="$wert10"}                     #
        $speicher += New-Object psobject -Property @{W1="$Wert5";W2="$wert11"}                     #
        $speicher += New-Object psobject -Property @{W1="$Wert6";W2="$wert12"}                     #
        }                                                                                          #
    "5" {                                                                                          #
        if($null -ne $we1) { $Wert1 += $we1 } else { $Wert1 += $null }                             #
        if($null -ne $we2) { $Wert2 += $we2 } else { $Wert2 += $null }                             #
        if($null -ne $we3) { $Wert3 += $we3 } else { $Wert3 += $null }                             #
        if($null -ne $we4) { $Wert4 += $we4 } else { $Wert4 += $null }                             #
        if($null -ne $we5) { $Wert5 += $we5 } else { $Wert5 += $null }                             #
        #######################################################################################    # 
        if($null -eq $we6 -or $we6.Length -eq 0) { $Wert6 += $F_Text } else { $Wert6 += $we6 }     #
        if($null -eq $we7 -or $we7.Length -eq 0) { $Wert7 += $F_Text } else { $Wert7 += $we7 }     #
        if($null -eq $we8 -or $we8.Length -eq 0) { $Wert8 += $F_Text } else { $Wert8 += $we8 }     #
        if($null -eq $we9 -or $we9.Length -eq 0) { $Wert9 += $F_Text } else { $Wert9 += $we9 }     #
        if($null -eq $we10 -or $we10.Length -eq 0) { $Wert10 += $F_Text } else { $Wert10 += $we10 }#  
        #######################################################################################    # 
        $speicher += New-Object psobject -Property @{W1="$Wert1";W2="$wert6"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert2";W2="$wert7"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert3";W2="$wert8"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert4";W2="$wert9"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert5";W2="$wert10"}                     #
        }                                                                                          #
    "4" {                                                                                          #
        if($null -ne $we1) { $Wert1 += $we1 } else { $Wert1 += $null }                             #
        if($null -ne $we2) { $Wert2 += $we2 } else { $Wert2 += $null }                             #
        if($null -ne $we3) { $Wert3 += $we3 } else { $Wert3 += $null }                             #
        if($null -ne $we4) { $Wert4 += $we4 } else { $Wert4 += $null }                             #
        #######################################################################################    # 
        if($null -eq $we5 -or $we5.Length -eq 0) { $Wert5 += $F_Text } else { $Wert5 += $we5 }     #
        if($null -eq $we6 -or $we6.Length -eq 0) { $Wert6 += $F_Text } else { $Wert6 += $we6 }     #
        if($null -eq $we7 -or $we7.Length -eq 0) { $Wert7 += $F_Text } else { $Wert7 += $we7 }     #
        if($null -eq $we8 -or $we8.Length -eq 0) { $Wert8 += $F_Text } else { $Wert8 += $we8 }     #
        #######################################################################################    # 
        $speicher += New-Object psobject -Property @{W1="$Wert1";W2="$wert5"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert2";W2="$wert6"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert3";W2="$wert7"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert4";W2="$wert8"}                      #
        }                                                                                          #
    "3" {                                                                                          #
        if($null -ne $we1) { $Wert1 += $we1 } else { $Wert1 += $null }                             #
        if($null -ne $we2) { $Wert2 += $we2 } else { $Wert2 += $null }                             #
        if($null -ne $we3) { $Wert3 += $we3 } else { $Wert3 += $null }                             #  
        #######################################################################################    #
        if($null -eq $we4 -or $we4.Length -eq 0) { $Wert4 += $F_Text } else { $Wert4 += $we4 }     #
        if($null -eq $we5 -or $we5.Length -eq 0) { $Wert5 += $F_Text } else { $Wert5 += $we5 }     #
        if($null -eq $we6 -or $we6.Length -eq 0) { $Wert6 += $F_Text } else { $Wert6 += $we6 }     #
        #######################################################################################    # 
        $speicher += New-Object psobject -Property @{W1="$Wert1";W2="$wert4"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert2";W2="$wert5"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert3";W2="$wert6"}                      #
        }                                                                                          #
    "2" {                                                                                          #
        if($null -ne $we1) { $Wert1 += $we1 } else { $Wert1 += $null }                             #
        if($null -ne $we2) { $Wert2 += $we2 } else { $Wert2 += $null }                             #
        #######################################################################################    #
        if($null -eq $we3 -or $we3.Length -eq 0) { $Wert3 += $F_Text } else { $Wert3 += $we3 }     #
        if($null -eq $we4 -or $we4.Length -eq 0) { $Wert4 += $F_Text } else { $Wert4 += $we4 }     #
        #######################################################################################    #
        $speicher += New-Object psobject -Property @{W1="$Wert1";W2="$wert3"}                      #
        $speicher += New-Object psobject -Property @{W1="$Wert2";W2="$wert4"}                      #
        }                                                                                          #
    "1" {                                                                                          #
        if($null -ne $we1) { $Wert1 += $we1 } else { $Wert1 += $null }                             #
        #######################################################################################    #
        if($null -eq $we2 -or $we2.Length -eq 0) { $Wert2 += $F_Text } else { $Wert2 += $we2 }     #
        #######################################################################################    # 
        $speicher += New-Object psobject -Property @{W1="$Wert1";W2="$wert2"}                      #
        }                                                                                          #
    default {}                                                                                     #
    }                                                                                              #
    if($pos -eq 'l') {                                                                             #
    $start = 3                                                                                     #
    $cut = 1                                                                                       #
    } else {                                                                                       #    
    $start = 1                                                                                     #
    $cut = $zeilen-2                                                                               #  
    }                                                                                              #
    foreach ($spei in $speicher) {                                                                 #
    $tmp = $null                                                                                   #
    if ($spei.W1.Length -lt $bre) {                                                                #
        $dif = $bre - $spei.W1.Length                                                              #
        for ($x=0;$x -lt $dif;$x++) { $tmp += "$leer" }                                            #
        if ($spei.W1 -match "^\d+$") { $spei.W1 = $tmp + $spei.W1 }                                #
            else { $spei.W1 = $spei.W1 + $tmp }                                                    #
        #if ($wer -match "^\d+$") { $wer = "$tmp$wer" } else { $wer = "$wer$tmp" }                 #
    } elseif ($null -eq $spei.W1) {                                                                #
        $dif = $bre + 1                                                                            #
        for ($x=0;$x -le $dif;$x++) { $tmp += "$leer" }                                            #
        $spei.W1 = "$tmp"                                                                          #
    } elseif ($spei.W1.Length -gt $bre) {                                                          #
        $spei.W1 = $spei.W1.Substring(0, $bre-1) + "~"                                             #
    }                                                                                              #
    $Sammeln[$start].W1 = $spei.W1                                                                 #  
    $Sammeln[$start].W2 = $spei.W2                                                                 #
    $start = $start + 2                                                                            #     
    }                                                                                              #
    $bese = $Sammeln[0].W1.Length + $Sammeln[$zeilen-1].W1.Length                                  #
    $besetzt = ($bre*$spa) + ($trenner.Length*$spa) + $bese                                        #
    $frei = $sb - $besetzt                                                                         #
    if($wex.Length -ge $frei) {                                                                    #
        $move = $wex.Substring(0,$frei-1)                                                          #
        $Sammeln[$cut].W1 = "$move"+"." ; $Sammeln[$cut].W2 = $F_Text                              #
    } elseif ($wex.Length -lt $frei) {                                                             #
        $tmp1 = $null                                                                              #
        $dif1 = $frei - $wex.Length                                                                #
        for($d=0;$d -lt $dif1;$d++) { $tmp1 = "$leer$tmp1" }                                       #
        $Sammeln[$cut].W1 = "$wex$tmp1" ; $Sammeln[$cut].W2 = $F_Text                              #
    } elseif ($null -eq $wex) {                                                                    #
        $Sammeln[$cut].W1 = "$wex" ; ; $Sammeln[$cut].W2 = $F_Text                                 #
    }                                                                                              #
    ### Ausgabe in Konsole #########################################################################
    if ($A_Con -eq 1) {                                                                            #
    for ($i=0;$i -lt $zeilen-1;$i++) {                                                             #
        Write-Host $Sammeln[$i].W1 -NoNewline -ForegroundColor $Sammeln[$i].W2                     #
    }                                                                                              #
    Write-Host $Sammeln[$zeilen-1].W1 -ForegroundColor $Sammeln[$zeilen-1].W2                      #
    }                                                                                              #
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {                                                                            #
        for ($i=0;$i -lt $zeilen-1;$i++) {                                                         #
            $buff = $Sammeln[$i].W1                                                                #
            $buffer = "$buffer$buff"                                                               #
            }                                                                                      #
        $buff2 = $Sammeln[$zeilen-1].W1 ; $buffer = "$buffer$buff2"                                #
        Ausgabe "$buffer"                                                                          #
    }                                                                                              #
}                                                                                                  #
####################################################################################################
# Ausgabe Funktionen fuer Text                                                                     #
####################################################################################################
function neu_text ($sub,$uze,[string]$ueb,[string]$text) {
    ### Legende ####################################################################################
    # $sub = um wieviele Zeichen eingerückt werden soll                                            #
    # $uze = Welches Zeichen für den Unterstrich verwendet werden soll                             #
    # $ueb = Ueberschrift oder Titel                                                               #
    # $text = Der auszugebende Text                                                                #
    ### Berechnung #################################################################################
    Merken 'Text' @{ Ueberschrift = $ueb; Text = $text }   # fuer HTML/JSON
    $anfang = "$zeichen " + (" " * $sub) ; $ende = " $zeichen" ; $Sammeln = @()
    $frei = $sb - $anfang.Length - $ende.Length
    # Umlaute bleiben erhalten (Text-Report ist jetzt UTF-8 mit BOM).
    if ($ueb) {
        $ueb_un = $uze * $ueb.Length
    }
    $worte = $text -split " "
    $akt_zeile = $worte[0]
    for ($i=1;$i -lt $worte.count;$i++) {
        $wort = $worte[$i]
        if (($akt_zeile + " " +$wort).Length -le $frei) { $akt_zeile += " " + $wort } 
        else { $Sammeln += $akt_zeile ; $akt_zeile = $wort }
        if($i -eq $worte.count - 1) { $Sammeln += $akt_zeile }
    }
    $Sammeln_count = $Sammeln.count

    ### Ausgabe Console ############################################################################
    if ($A_Con -eq 1) {
        if($ueb) {
            Write-Host $anfang -NoNewline -ForegroundColor $F_Rahmen
            Write-Host $ueb.PadRight($frei) -NoNewline -ForegroundColor $F_Text
            Write-Host $ende -ForegroundColor $F_Rahmen
            Write-Host $anfang -NoNewline -ForegroundColor $F_Rahmen
            Write-Host $ueb_un.PadRight($frei) -NoNewline -ForegroundColor $F_Text
            Write-Host $ende -ForegroundColor $F_Rahmen
        }
        for ($a=0;$a -lt $Sammeln_count;$a++) {
            Write-Host $anfang -NoNewline -ForegroundColor $F_Rahmen
            Write-Host $Sammeln[$a].PadRight($frei) -NoNewline -ForegroundColor $F_Text
            Write-Host $ende -ForegroundColor $F_Rahmen 
        }
    }
    ### Ausgabe in Datei ###########################################################################
    if ($A_Dat -eq 1) {
        if($ueb) {
            $tmp_1 = $anfang + $ueb.PadRight($frei) + $ende
            $tmp_2 = $anfang + $ueb_un.PadRight($frei) + $ende
            Ausgabe "$tmp_1"
            Ausgabe "$tmp_2"
        }
        for ($a=0;$a -lt $Sammeln_count;$a++) {
            $tmp_x = $anfang + $Sammeln[$a].PadRight($frei) + $ende
            Ausgabe "$tmp_x"
        }
    }
}
####################################################################################################
# Fehlerbehandlung: Kapselt einen Pruefbereich in try/catch                                        #
####################################################################################################
function Pruefbereich ($titel,[scriptblock]$aktion,$CheckId) {
    ### Legende ####################################################################################
    # $titel   = Ueberschrift des Bereichs (wird an Bereich durchgereicht)                         #
    # $aktion  = Scriptblock mit den Pruef-Funktionen des Bereichs                                 #
    # $CheckId = optionale Katalog-ID; emittiert die Begruendung (Doku) fuer HTML/JSON             #
    # Ein Fehler im Bereich bricht den Lauf nicht ab: Er wird im Report vermerkt, auf der          #
    # Konsole rot gemeldet, und es geht mit dem naechsten Bereich weiter.                          #
    ################################################################################################
    Bereich $titel
    if ($CheckId) { Doku $CheckId }
    try { & $aktion }
    catch {
        $f_meldung = $_.Exception.Message
        $f_zeile = $_.InvocationInfo.ScriptLineNumber
        Leerzeile
        neu_text 0 '!' "FEHLER - Bereich nur unvollständig geprüft" `
            "Meldung: $f_meldung (Skriptzeile $f_zeile). Der Lauf wird mit dem nächsten Bereich fortgesetzt."
        Leerzeile
        if ($A_Con -eq 1) {
            Write-Host "FEHLER im Bereich '$titel': $f_meldung" -ForegroundColor $F_Fehler
        }
    }
    finally { Puffer_leeren }                    # Bereich fertig -> Puffer in die Datei schreiben #
}
function Unterpruefung ($titel,$checkid,[scriptblock]$aktion) {
    ### Legende ####################################################################################
    # Teilpruefung innerhalb eines Bereichs (z.B. einzelner Kerberos-Check).                       #
    # $titel   = Unterueberschrift (Bereichstitel -> h3)                                           #
    # $checkid = Katalog-ID; emittiert direkt danach die Begruendung (Doku)                        #
    # $aktion  = Scriptblock mit der eigentlichen Pruefung; eigener try/catch, damit ein           #
    #            Fehler nur diese Teilpruefung ueberspringt, nicht den ganzen Bereich.             #
    ################################################################################################
    Bereichstitel $titel
    if ($checkid) { Doku $checkid }
    Leerzeile
    try { & $aktion }
    catch {
        neu_text 0 '!' "Hinweis: Teilprüfung übersprungen" `
            "Meldung: $($_.Exception.Message). Die übrigen Prüfungen dieses Bereichs laufen weiter."
    }
    Leerzeile
}
####################################################################################################
# Export-Funktionen: HTML-Report (im Stil von water.css) und JSON-Export                           #
####################################################################################################
function Farbklasse ($farbe) {
    # Konsolenfarbe -> CSS-Klasse fuer den HTML-Report (Befund-Status)
    switch -regex ("$farbe") {
        '^(dark)?red$'    { 'err'  ; break }
        '^(dark)?green$'  { 'ok'   ; break }
        '^(dark)?yellow$' { 'warn' ; break }
        default           { '' }
    }
}
function HTML_Report {
    # Erzeugt aus den gesammelten Ereignissen ($R_Daten) einen eigenstaendigen HTML-Report.
    # Styling: kompaktes, water.css-inspiriertes CSS (eingebettet, offline-faehig,
    # hell/dunkel folgt automatisch der Systemeinstellung).
    function Esc ($t) {
        # Nur HTML-kritische Zeichen escapen; Umlaute bleiben als echtes UTF-8 erhalten
        # (Datei ist UTF-8, charset=utf-8). Reihenfolge: & zuerst.
        "$t" -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
    }
    $ziel = [System.IO.Path]::ChangeExtension($path, 'html')
    $zielverz = Split-Path -Parent $ziel
    if (!(Test-Path $zielverz)) { New-Item -Path $zielverz -ItemType Directory | Out-Null }
    $css = @'
:root{--bg:#ffffff;--text:#363636;--muted:#70777f;--border:#dbdbdb;--accent:#0076d1;--zebra:rgba(125,125,125,.07);--ok:#2e7d32;--warn:#a86500;--err:#c62828;--fehler-bg:rgba(198,40,40,.08);--info-bg:rgba(0,118,209,.06)}
@media(prefers-color-scheme:dark){:root{--bg:#1c1f25;--text:#dbdbdb;--muted:#8b939e;--border:#3a3f4b;--accent:#41adff;--zebra:rgba(125,125,125,.10);--ok:#7bc67e;--warn:#e0a458;--err:#ef6461;--fehler-bg:rgba(239,100,97,.12);--info-bg:rgba(65,173,255,.10)}}
*{box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',sans-serif;max-width:960px;margin:0 auto;padding:1.5rem 1rem 3rem;background:var(--bg);color:var(--text);line-height:1.5}
h1{font-size:1.9rem;margin:.5rem 0 .2rem}
h2{font-size:1.35rem;margin:2.2rem 0 .6rem;padding-bottom:.25rem;border-bottom:2px solid var(--accent);scroll-margin-top:.5rem}
h3{font-size:1.1rem;margin:1.5rem 0 .4rem}
h4{font-size:1rem;margin:1.1rem 0 .3rem;color:var(--muted)}
p{margin:.4rem 0}
table{border-collapse:collapse;width:100%;margin:.4rem 0 1rem;font-size:.92rem}
td{border:1px solid var(--border);padding:.3rem .6rem;vertical-align:top;text-align:left}
tr:nth-child(even){background:var(--zebra)}
table.kv td:first-child{width:42%;color:var(--muted)}
.ok{color:var(--ok)}.warn{color:var(--warn)}.err{color:var(--err)}
.meta{color:var(--muted);font-size:.9rem}
.fehler{border-left:4px solid var(--err);background:var(--fehler-bg);padding:.6rem .9rem;margin:.8rem 0}
footer{margin-top:3rem;border-top:1px solid var(--border);padding-top:.6rem}
.badge{font-size:.72rem;font-weight:500;padding:1px 8px;border-radius:6px;vertical-align:middle;margin-left:8px;white-space:nowrap}
.sev-krit{background:#F7C1C1;color:#501313}.sev-hoch{background:#F5C4B3;color:#4A1B0C}.sev-mit{background:#FAC775;color:#412402}.sev-nied{background:#B5D4F4;color:#042C53}.sev-info{background:#D3D1C7;color:#2C2C2A}
details.doku{background:var(--info-bg);border:1px solid var(--border);border-radius:8px;padding:.3rem .8rem;margin:.2rem 0 1.2rem;font-size:.9rem}
details.doku summary{cursor:pointer;color:var(--accent);font-weight:500;padding:.3rem 0}
details.doku p{margin:.5rem 0;line-height:1.55}
details.doku .lbl{font-weight:500;color:var(--muted)}
.zus{margin:1rem 0 1.5rem;padding:.8rem 1rem;background:var(--zebra);border-radius:8px}
.zus h2{margin:.2rem 0 .6rem;border:0;padding:0}
.zus .counts{margin:.2rem 0 .8rem}
.zus ul{margin:.2rem 0;padding-left:1.1rem;font-size:.95rem}
.zus li{margin:.15rem 0}
.zus a{color:var(--accent);text-decoration:none}
'@
    $H = New-Object System.Text.StringBuilder
    $kopf = $R_Daten | Where-Object { $_.Art -eq 'Kopf' } | Select-Object -First 1
    $titelH = if ($kopf) { $kopf.Titel } else { $maintitel }
    [void]$H.AppendLine('<!DOCTYPE html>')
    [void]$H.AppendLine('<html lang="de">')
    [void]$H.AppendLine('<head>')
    [void]$H.AppendLine('<meta charset="utf-8">')
    [void]$H.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$H.AppendLine("<title>$(Esc $titelH) - $(Esc $system)</title>")
    [void]$H.AppendLine("<style>$css</style>")
    [void]$H.AppendLine('</head>')
    [void]$H.AppendLine('<body>')
    [void]$H.AppendLine("<header><h1>$(Esc $titelH)</h1>")
    if ($kopf) {
        [void]$H.AppendLine("<p class=""meta"">$(Esc $kopf.Typ) &middot; System: $(Esc $kopf.System) &middot; $(Esc $kopf.Datum) &middot; $(Esc $kopf.Firma)</p>")
    }
    [void]$H.AppendLine('</header>')
    function SevKlasse ($s) {
        switch ("$s") {
            'Kritisch' { 'sev-krit' ; break }
            'Hoch'     { 'sev-hoch' ; break }
            'Mittel'   { 'sev-mit'  ; break }
            'Niedrig'  { 'sev-nied' ; break }
            default    { 'sev-info' }
        }
    }
    ### Zusammenfassung: Pruefbereiche nach Einstufung, mit Sprungmarken ##########################
    $dokus = @($R_Daten | Where-Object { $_.Art -eq 'Doku' })
    if ($dokus.Count -gt 0) {
        $cnt = [ordered]@{ Kritisch = 0; Hoch = 0; Mittel = 0; Niedrig = 0; Info = 0 }
        foreach ($d in $dokus) { if ($cnt.Contains("$($d.Schwere)")) { $cnt["$($d.Schwere)"]++ } }
        [void]$H.AppendLine('<section class="zus">')
        [void]$H.AppendLine('<h2>Zusammenfassung</h2>')
        $teile = foreach ($s in $cnt.Keys) { "<span class=""badge $(SevKlasse $s)"">$s $($cnt[$s])</span>" }
        [void]$H.AppendLine("<div class=""counts"">$($teile -join ' ')</div>")
        [void]$H.AppendLine('<ul>')
        foreach ($d in $dokus) {
            [void]$H.AppendLine("<li><a href=""#chk-$($d.CheckId)"">$(Esc $d.DTitel)</a> <span class=""badge $(SevKlasse $d.Schwere)"">$(Esc $d.Schwere)</span></li>")
        }
        [void]$H.AppendLine('</ul>')
        [void]$H.AppendLine('<p class="meta">Hinweis: Die Einstufung bewertet die Wichtigkeit des Prüfbereichs, nicht zwingend einen konkreten Befund. Begründung und Empfehlung je Bereich im jeweiligen Block "Hintergrund &amp; Empfehlung".</p>')
        [void]$H.AppendLine('</section>')
    }
    $offen = ''                                  # aktuell geoeffnete Tabelle: '' | 'kv' | 'tab'
    for ($ix = 0; $ix -lt $R_Daten.Count; $ix++) {
        $e = $R_Daten[$ix]
        if ($e.Art -ne 'Wert' -and $e.Art -ne 'TabZeile' -and $offen) {
            [void]$H.AppendLine('</table>') ; $offen = ''
        }
        switch ($e.Art) {
            'Bereich'  {
                # Folgt direkt ein Doku-Ereignis, bekommt die Ueberschrift Anker-ID und Severity-Badge.
                $next = if ($ix + 1 -lt $R_Daten.Count) { $R_Daten[$ix + 1] } else { $null }
                if ($next -and $next.Art -eq 'Doku') {
                    # Katalog-Titel (mit Umlauten) statt des ASCII-Bereichstitels verwenden.
                    [void]$H.AppendLine("<h2 id=""chk-$($next.CheckId)"">$(Esc $next.DTitel) <span class=""badge $(SevKlasse $next.Schwere)"">$(Esc $next.Schwere)</span></h2>")
                } else {
                    [void]$H.AppendLine("<h2>$(Esc $e.Titel)</h2>")
                }
            }
            'Doku'     {
                [void]$H.AppendLine('<details class="doku"><summary>Hintergrund &amp; Empfehlung</summary>')
                [void]$H.AppendLine("<p><span class=""lbl"">Zweck:</span> $(Esc $e.Zweck)</p>")
                if ("$($e.Hintergrund)".Trim()) {
                    [void]$H.AppendLine("<p><span class=""lbl"">Technischer Hintergrund:</span> $(Esc $e.Hintergrund)</p>")
                }
                [void]$H.AppendLine("<p><span class=""lbl"">Beispiel:</span> $(Esc $e.Beispiel)</p>")
                [void]$H.AppendLine("<p><span class=""lbl"">Empfehlung:</span> $(Esc $e.Empfehlung)</p>")
                # Quellen: String -> Text; Liste @{Titel;Url} -> klickbare, geprueifte Links
                if ($e.Quellen -is [string]) {
                    [void]$H.AppendLine("<p><span class=""lbl"">Quellen:</span> $(Esc $e.Quellen)</p>")
                } else {
                    $links = @($e.Quellen | ForEach-Object {
                        "<a href=""$(Esc $_.Url)"" target=""_blank"" rel=""noopener noreferrer"">$(Esc $_.Titel)</a>"
                    }) -join ' &middot; '
                    [void]$H.AppendLine("<p><span class=""lbl"">Quellen:</span> $links</p>")
                }
                [void]$H.AppendLine('</details>')
            }
            'Titel'    {
                # Wie bei 'Bereich': folgt ein Doku-Ereignis, bekommt die Unterueberschrift
                # Anker-ID und Severity-Badge (fuer Unterpruefungen der Sicherheits-Pakete).
                $next = if ($ix + 1 -lt $R_Daten.Count) { $R_Daten[$ix + 1] } else { $null }
                if ($next -and $next.Art -eq 'Doku') {
                    # Katalog-Titel (mit Umlauten) statt des ASCII-Untertitels verwenden.
                    [void]$H.AppendLine("<h3 id=""chk-$($next.CheckId)"">$(Esc $next.DTitel) <span class=""badge $(SevKlasse $next.Schwere)"">$(Esc $next.Schwere)</span></h3>")
                } else {
                    [void]$H.AppendLine("<h3>$(Esc $e.Text)</h3>")
                }
            }
            'Subtitel' { [void]$H.AppendLine("<h4>$(Esc $e.Text)</h4>") }
            'Wert'     {
                if ($offen -ne 'kv') {
                    if ($offen) { [void]$H.AppendLine('</table>') }
                    [void]$H.AppendLine('<table class="kv">') ; $offen = 'kv'
                }
                $kl = Farbklasse $e.Farbe
                $td = if ($kl) { "<td class=""$kl"">" } else { '<td>' }
                [void]$H.AppendLine("<tr><td>$(Esc $e.Name)</td>$td$(Esc $e.Wert)</td></tr>")
            }
            'TabZeile' {
                if ($offen -ne 'tab') {
                    if ($offen) { [void]$H.AppendLine('</table>') }
                    [void]$H.AppendLine('<table>') ; $offen = 'tab'
                }
                $zellen = ''
                for ($i = 0; $i -lt $e.Zellen.Count; $i++) {
                    $kl = Farbklasse $e.Farben[$i]
                    if ($kl) { $zellen += "<td class=""$kl"">$(Esc $e.Zellen[$i])</td>" }
                    else     { $zellen += "<td>$(Esc $e.Zellen[$i])</td>" }
                }
                if ("$($e.Extra)".Trim()) { $zellen += "<td>$(Esc $e.Extra)</td>" }
                [void]$H.AppendLine("<tr>$zellen</tr>")
            }
            'Text'     {
                if ("$($e.Ueberschrift)" -match '^FEHLER') {
                    [void]$H.AppendLine("<div class=""fehler""><strong>$(Esc $e.Ueberschrift)</strong><br>$(Esc $e.Text)</div>")
                } else {
                    if ("$($e.Ueberschrift)".Trim()) { [void]$H.AppendLine("<h4>$(Esc $e.Ueberschrift)</h4>") }
                    [void]$H.AppendLine("<p>$(Esc $e.Text)</p>")
                }
            }
        }
    }
    if ($offen) { [void]$H.AppendLine('</table>') }
    [void]$H.AppendLine("<footer><p class=""meta"">$(Esc $madeby) &middot; $(Esc $firma)</p></footer>")
    [void]$H.AppendLine('</body>')
    [void]$H.AppendLine('</html>')
    # Mit BOM schreiben, damit Umlaute auch ausserhalb des Browsers (Editoren, PS 5.1) stimmen.
    [System.IO.File]::WriteAllText($ziel, $H.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    if ($A_Con -eq 1) { Write-Host "HTML-Report : $ziel" -ForegroundColor $F_Ue_Schrift }
}
function JSON_Export {
    # Schreibt die gesammelten Ereignisse als strukturierte JSON-Datei (fuer Auswertungen,
    # z.B. im Kontext einer geplanten AD-Abloesung).
    $ziel = [System.IO.Path]::ChangeExtension($path, 'json')
    $zielverz = Split-Path -Parent $ziel
    if (!(Test-Path $zielverz)) { New-Item -Path $zielverz -ItemType Directory | Out-Null }
    $export = [pscustomobject]@{
        Skript     = $maintitel
        Version    = $version
        System     = $system
        Datum      = $datum
        Ereignisse = $R_Daten
    }
    $json = $export | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($ziel, $json, (New-Object System.Text.UTF8Encoding($false)))
    if ($A_Con -eq 1) { Write-Host "JSON-Export : $ziel" -ForegroundColor $F_Ue_Schrift }
}
####################################################################################################
# Funktionen zum Auslesen                                                                          #
####################################################################################################
# Globale Variablen fuer den Bereich "Domain, Mode, Controller, FSMO"
#####################################################################
try { $DCs = (Get-ADForest).Domains | ForEach-Object{Get-ADDomainController -Filter * -Server $_ } }
catch { $DCs = $null }
if (-not $DCs) {
    Write-Host "FEHLER : Domain Controller konnten nicht ermittelt werden." -ForegroundColor $F_Fehler
    Write-Host "         Besteht eine Verbindung zur Domäne? Das Skript wird beendet." -ForegroundColor $F_Fehler
    exit 1
}
####################################################################################################
# Funktionen fuer den Bereich "Domain, Mode, FSMO"                                                 #
####################################################################################################
function dom_allgemein {
    ################################################################################################
    # Generierung der Variablen und Pruefung                                                       #
    ################################################################################################
    $dom_all = Get-ADDomain
    $for_rol = Get-ADForest
    [string]$Domain_mode = (Get-ADDomain).DomainMode ; $Domain_mode = $Domain_mode.TrimEnd("Domain")
    [string]$forest_mode = (Get-ADForest).ForestMode ; $forest_mode = $forest_mode.TrimEnd("Forest")
    $akt_com = $dom_all.ComputersContainer ; $std_com = "CN=Computers," + $dom_all.DistinguishedName
    $akt_usr = $dom_all.UsersContainer ; $std_usr = "CN=Users," + $dom_all.DistinguishedName
    $akt_dc = $dom_all.DomainControllersContainer
    $std_dc = "OU=Domain Controllers," + $dom_all.DistinguishedName
    if ($akt_com -ne $std_com) { $com_fa = "Green" } else { $com_fa = "Red" }
    if ($akt_usr -ne $std_usr) { $usr_fa = "Green" } else { $usr_fa = "Red" }
    if ($akt_dc -ne $std_dc) { $dc_fa = "Red" } else { $dc_fa = "Green" }
    $recy = (Get-ADOptionalFeature -Filter {Name -like "Recycle Bin Feature"}).EnabledScopes
    if ($recy) { $papier = "aktiviert" ; $pap_fa = "Green" } 
        else { $papier = "nicht aktiviert" ; $pap_fa = "Red" ; $recy = "nicht aktiviert" }
    $path_AD = (Get-ADDomain).DeletedObjectsContainer
    if ($path_AD) { $pat_fa = "Green" } else { $pat_fa = "Red" ; $path_AD = "nicht aktiviert" }
    $qu = "ms-DS-MachineAccountQuota"
    $quota = (Get-ADObject((Get-ADDomain).distinguishedname) `
     -Properties ms-DS-MachineAccountQuota).$qu
    if ($quota -gt 0) { $fa_qo = "Red" } else { $fa_qo = "Green" }
    $kerb_pls = Get-Date (Get-ADUser -Identity krbtgt -Properties *).PasswordLastSet
    $kerb_akt = Get-Date
    if(($kerb_akt - $kerb_pls).Days -gt 730) { $kerb_fa = "Red" } 
        elseif (($kerb_akt - $kerb_pls).Days -le 365) { $kerb_fa = "Green" } 
            else { $kerb_fa = "Yellow" }
    $kerb_pls = $kerb_pls.ToString("dd.MM.yyyy HH:mm:ss")
    if($schema -eq 1) {
        $w_laps = Invoke-Command -ComputerName $for_rol.SchemaMaster -ScriptBlock {
            $w_laps_tp = Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) `
            -Filter { name -like 'ms-LAPS*' }
            return $w_laps_tp }
        $w_laps_count = ($w_laps.Name).count
        if ($w_laps_count -ge 6) { $w_laps_text = "Attribute vorhanden" ; $w_laps_fa = "Green" }
            else { $w_laps_text = "Attribute fehlen" ; $w_laps_fa = "Red" }
        $l_laps = Invoke-Command -ComputerName $for_rol.SchemaMaster -ScriptBlock {
            $l_laps_tp = Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) `
            -Filter { name -like 'ms-Mcs-Adm*' }
            return $l_laps_tp
        }
        $l_laps_count = ($l_laps.Name).count
        if ($l_laps_count -ge 2) { $l_laps_text = "Attribute vorhanden" ; $l_laps_fa = "Green" } 
            else { $l_laps_text = "Attribute fehlen" ; $l_laps_fa = "Red" }
        $b_lock = Invoke-Command -ComputerName $for_rol.SchemaMaster -ScriptBlock {
            $b_lock_tp = Get-ADObject -SearchBase ((GET-ADRootDSE).SchemaNamingContext) `
            -Filter {Name -like 'ms-FVE-*'}
            return $b_lock_tp
        }
        $b_lock_count = ($b_lock.Name).count
        if ($b_lock_count -ge 5) { $b_lock_text = "Attribute vorhanden" ; $b_lock_fa = "Green" } 
            else { $b_lock_text = "Attribute fehlen" ; $b_lock_fa = "Red" }
        $jitjea = Invoke-Command -ComputerName $for_rol.SchemaMaster -ScriptBlock { 
            $jitjea_tp = Get-ADObject -SearchBase ((GET-ADRootDSE).SchemaNamingContext) `
            -Filter {Name -like 'ms-DS-Entry-Time*'}
            return $jitjea_tp
        }
        $jitjea_count = ($jitjea.Name).count
        if ($jitjea_count -ge 1) { $jitjea_text = "Attribute vorhanden" ; $jitjea_fa = "Green" } 
            else { $jitjea_text = "Attribute fehlen" ; $jitjea_fa = "Red" }
    }
    ################################################################################################
    # Ausgabe                                                                                      #
    ################################################################################################
    Bereichstitel "Domain allgemein:"
    Leerzeile # Domain und Dist. Name
    new_2werte "s" ":" "23" "Domain Name" "" "l" $dom_all.DNSRoot "Green" "l"
    new_2werte "s" ":" "23" "Distinguished Name" "" "l" $dom_all.DistinguishedName "Green" "l"
    Leerzeile # domain und forest mode
    new_2werte "s" ":" "23" "Domain Mode" "" "l" $Domain_mode "Green" "l"
    new_2werte "s" ":" "23" "Forest Mode" "" "l" $forest_mode "Green" "l"
    Leerzeile # Standard OU's
    Subtitel "Standard OU's:" "1" "-"
    new_2werte "s" ":" "23" "- Computer" "" "l" $akt_com $com_fa "l"
    new_2werte "s" ":" "23" "- Users" "" "l" $akt_usr $usr_fa "l"
    new_2werte "s" ":" "23" "- Domain Controllers" "" "l" $akt_dc $dc_fa "l"
    Leerzeile # FSMO Rollen
    Subtitel "FSMO Rollen:" "1" "-"
    new_2werte "s" ":" "23" "- PDC-Emulator" "" "l" $dom_all.PDCEmulator "" "l"
    new_2werte "s" ":" "23" "- Schema-Master" "" "l" $for_rol.SchemaMaster "" "l"
    new_2werte "s" ":" "23" "- RID-Pool-Master" "" "l" $dom_all.RIDMaster "" "l"
    new_2werte "s" ":" "23" "- Domain-Name-Master" "" "l" $for_rol.DomainNamingMaster "" "l"
    new_2werte "s" ":" "23" "- Infrastruktur-Master" "" "l" $dom_all.InfrastructureMaster "" "l"
    Leerzeile # Kerberos Key 
    Subtitel "Alter des Kerberos Keys:" "1" "-"
    new_2werte "s" ":" "23" "- Generierungsdatum" "" "l" "$kerb_pls" $kerb_fa "l"
    Leerzeile # AD-Papierkorb
    Subtitel "AD-Papierkorb:" "1" "-"
    new_2werte "s" ":" "23" "- Status" "" "l" $papier $pap_fa "l"
    new_2werte "s" ":" "23" "- Scope" "" "l" "$recy" $pap_fa "" "l"
    new_2werte "s" ":" "23" "- Pfad in der AD" "" "l" "$path_AD" $pat_fa "l"
    Leerzeile # msds MachineAccountQuota
    Subtitel "ms-DS-MachineAccountQuota:" "1" "-"
    new_2werte "s" ":" "23" "- Domain Joins pro User" "" "l" "$quota" $fa_qo "l"
    if ($schema -eq 1) {
        Leerzeile # Attribut Check
        Subtitel "Attribute zu Schema-Erweiterungen:" "1" "-"
        new_2werte "s" ":" "23" "- Legacy LAPS" "" "l" "$l_laps_text" "$l_laps_fa" "l"
        new_2werte "s" ":" "23" "- Windows LAPS" "" "l" "$w_laps_text" "$w_laps_fa" "l"
        new_2werte "s" ":" "23" "- BitLocker" "" "l" "$b_lock_text" "$b_lock_fa" "l"
        new_2werte "s" ":" "23" "- JiT/JEA" "" "l" "$jitjea_text" "$jitjea_fa" "l"
    }
    Leerzeile
}
####################################################################################################
# Funktionen fuer den Bereich "Central Store & Templates"                                          #
####################################################################################################
function centralstore {
    $fqdn = (Get-ADDomain).DNSRoot
    $weg = '\\' + $fqdn + '\Sysvol\' + $fqdn + '\Policies\PolicyDefinitions\'
    $weg_de = '\\' + $fqdn + '\Sysvol\' + $fqdn + '\Policies\PolicyDefinitions\de-DE\'
    $weg_en = '\\' + $fqdn + '\Sysvol\' + $fqdn + '\Policies\PolicyDefinitions\en-US\'
    $exist = Test-Path -path $weg
    $exist_de = Test-Path -path $weg_de
    $exist_en = Test-Path -path $weg_en
    if($exist -eq $true) 
        {
            $exist_fa = "Green"
            $exist = "Zentral Store wurde angelegt"
            $fi_cs_za = (Get-ChildItem -Path $weg -File).count
            if($fi_cs_za -eq "0") { $fi_cs_za_fa = "Red" } else { $fi_cs_za_fa = "Green" }
        }
        else { $exist_fa = "Red" ; $exist = "Zentral Store nicht vorhanden" }
    if($exist_de -eq $true) 
        { 
            $exist_de_fa = "Green"
            $exist_de = 'vorhanden'
            $fi_de_za = (Get-ChildItem -Path $weg_de -File).count
            if($fi_de_za -eq "0") { $fi_de_za_fa = "Red" } else { $fi_de_za_fa = "Green" }
        }
        else { $exist_de_fa = "Red" ; $exist_de = 'nicht vorhanden' }
    if($exist_en -eq $true) 
        {
            $exist_en_fa = "Green"
            $exist_en = 'vorhanden'
            $fi_en_za = (Get-ChildItem -Path $weg_en -File).count
            if($fi_en_za -eq "0") { $fi_en_za_fa = "Red" } else { $fi_en_za_fa = "Green" }
        }
        else { $exist_en_fa = "Red" ; $exist_en = 'nicht vorhanden' }
    ################################################################################################
    # Ausgabe                                                                                      #
    ################################################################################################
    Leerzeile
    Subtitel "Check Central Store:" "1" "-"
    new_2werte "s" ":" "23" ' Status Central Store  ' "" "l" "$exist" "$exist_fa" "l"
    new_2werte "s" ":" "23" ' - Anzahl der Templates' "" "l" "$fi_cs_za" "$fi_cs_za_fa" "l"
    new_2werte "s" ":" "23" ' Status "de_DE" Verz.  ' "" "l" "$exist_de" "$exist_de_fa" "l"
    new_2werte "s" ":" "23" ' - Anzahl der Templates' "" "l" "$fi_de_za" "$fi_de_za_fa" "l"
    new_2werte "s" ":" "23" ' Status "en_US" Verz.  ' "" "l" "$exist_en" "$exist_en_fa" "l"
    new_2werte "s" ":" "23" ' - Anzahl der Templates' "" "l" "$fi_en_za" "$fi_en_za_fa" "l"
    Leerzeile
}
function sec_templates {
    Subtitel "Check der Sec. Templates:" "1" "-"
    Leerzeile
    ################################################################################################
    # Pfad Variablen                                                                               #
    ################
    $fqdn = (Get-ADDomain).DNSRoot
    $we_ =   '\\' + $fqdn + '\Sysvol\' + $fqdn + '\Policies\PolicyDefinitions\'
    $we_de = '\\' + $fqdn + '\Sysvol\' + $fqdn + '\Policies\PolicyDefinitions\de-DE\'
    $we_en = '\\' + $fqdn + '\Sysvol\' + $fqdn + '\Policies\PolicyDefinitions\en-US\'
    ################################################################################################
    # Array füllen                                                                                 #
    ##############
    $templates = @()
    $templates += "AdmPwd"
    $templates += "MSS-legacy"
    $templates += "SecGuide"
    $templates += "LAPS"
    $templates += "Set-NetBIOS-node-type-KB160177"
    ################################################################################################
    foreach ($tem in $templates) 
    {
        Subtitel $tem "2" "-"
        $pa1 = $we_ + $tem + '.admx'
        $pa2 = $we_de + $tem + '.adml'
        $pa3 = $we_en + $tem + '.adml'
        if(Test-Path $pa1) 
        { new_2werte "s" ":" "23" ' Status ADMX-File' "" "l" "vorhanden" "green" "l" } else 
        { new_2werte "s" ":" "23" ' Status ADMX-File' "" "l" "nicht vorhanden" "red" "l" }
        if(Test-Path $pa2) 
        { new_2werte "s" ":" "23" ' - DE ADML-File' "" "l" "vorhanden" "green" "l" } else 
        { new_2werte "s" ":" "23" ' - DE ADML-File' "" "l" "nicht vorhanden" "red" "l" }
        if(Test-Path $pa3) 
        { new_2werte "s" ":" "23" ' - EN ADML-File' "" "l" "vorhanden" "green" "l" } else 
        { new_2werte "s" ":" "23" ' - EN ADML-File' "" "l" "nicht vorhanden" "red" "l" }
        Leerzeile
    }
}
####################################################################################################
# Funktionen fuer den Bereich "Domain Controller"                                                  #
####################################################################################################
function controller {
    foreach($DC in $DCs) {
        $tp1 = $DC.Hostname
        $tp2 = $DC.IPv4Address
        $tp3 = $DC.OperatingSystem
        $tp4 = $DC.OperatingSystemVersion
        $tp5 = Invoke-Command -ComputerName $tp1 -ScriptBlock {
            $installDate = (Get-WmiObject -Class Win32_OperatingSystem).InstallDate
            $dateTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($installDate)
            return (Get-Date -Date $dateTime -Format "dd.MM.yyyy HH:mm")
        }
        Subtitel $tp1 "1" "-"
        new_2werte "s" ":" "23" "- IPv4Address" "" "l" $tp2 "" "l"
        new_2werte "s" ":" "23" "- OS" "" "l" $tp3 "" "l"
        new_2werte "s" ":" "23" "- Version" "" "l" $tp4 "" "l"
        new_2werte "s" ":" "23" "- Installiert am" "" "l" $tp5 "" "l"
        Leerzeile
    }
}
####################################################################################################
# Funktionen fuer den Bereich "Logging auf Domain Controller(n)"                                   #
####################################################################################################
function Event_dienst {
    Bereichstitel "Status des EventLog Dienst:" 
    Leerzeile
    foreach($dc in $DCs){
        Subtitel "Server: $dc" "1" "-"
        $s_ev = Invoke-Command -ComputerName $dc -ScriptBlock{Get-Service -Name EventLog}
        if($s_ev.Status -eq "Running") { $s_ev_f = "Green"} else { $s_ev_f = "Red" }
        new_2werte "s" ":" "23" "EventLog Dienst Status" "" "l" $s_ev.Status $s_ev_f "l"
        Leerzeile
    }
}
function Auditcheck {
    Bereichstitel "Einstellung der lokalen Auditing Richtlinen:" 
    Leerzeile
    foreach($dc in $DCs){
        Subtitel "Server: $dc" "1" "-"
        $audit = Invoke-Command -ComputerName $dc -ScriptBlock{
            $buffer_1 = [System.IO.Path]::GetTempFileName()
            $buffer_2 = [System.IO.Path]::GetTempFileName()
            SecEdit /export /cfg $buffer_1 /log $buffer_2 | Out-Null
            Get-Content -Path $buffer_1 | Where-Object {$_ -like "Audit*"}
            Remove-Item -Path $buffer_1
            Remove-Item -Path $buffer_2
        } -ErrorAction SilentlyContinue
        foreach($au in $audit){
            $aux = $au.Replace(" = ",",")
            $au12 = $aux.Split(",") ; $au1 = $au12[0] ; $au2 = $au12[1]
            if($au2 -eq 0) { $aus = "Inaktiv!" ; $fx = "Red" }
            if($au2 -eq 1) { $aus = "Erfolgreich aktiv!" ; $fx = "Yellow" }
            if($au2 -eq 2) { $aus = "Fehler aktiv!" ; $fx = "Yellow" }
            if($au2 -eq 3) { $aus = "Erfolgreich und Fehler aktiv!" ; $fx = "Green" }
            $au1 = "- $au1"
            new_2werte "s" ":" "23" "$au1" "" "l" $aus $fx "l"
        }
    Leerzeile
    }
}
####################################################################################################
# Funktionen fuer den Bereich "AS-Trusts - Check"                                                  #
####################################################################################################
function trusts {
    foreach ($dc in $DCs){
        $dc_name = $dc.HostName
        Bereichstitel $dc_name
        Leerzeile 
        $trusts = Get-ADTrust -Filter * -Properties *
        $owndom = (Get-ADDomain).DNSRoot
        foreach ($trust in $trusts) {
            $t_source = (Get-ADDomain -Identity $trust.Source).DNSRoot
            $direct = (Get-ADTrust -Identity $trust.Name -Properties *).Direction
            if ($t_source -eq $owndom) { $t_source_f = "Green" } else { $t_source_f = "$F_Text"}
            $t_zinfo = Get-DnsServerZone -ComputerName $dc.HostName  | `
                Where-Object {$_.ZoneType -eq 'Forwarder' -and $_.ZoneName -eq $trust.Target}
            $t_ip = ($t_zinfo.MasterServers).IPAddressToString
            $t_wcreated = ($trust.whenCreated).ToString("dd.MM.yyyy HH:mm")
            $t_wchanged = ($trust.whenChanged).ToString("dd.MM.yyyy HH:mm")
            $tr_titel = "Trusted Domain: " + $trust.Name
            Subtitel $tr_titel "1" "-"
            new_2werte "s" ":" "23" "- Source Domain" "" "l" $t_source $t_source_f "l"
            new_2werte "s" ":" "23" "- Target Domain" "" "l" $trust.Target "" "l"
            Leerzeile
            new_2werte "s" ":" "23" "- Direction" "" "l" "$direct" "" "l"
            new_2werte "s" ":" "23" "- Gebunden an IP" "" "l" "$t_ip" "" "l"
            new_2werte "s" ":" "23" "- Wurde erstellt am" "" "l" "$t_wcreated" "" "l"
            new_2werte "s" ":" "23" "- Wurde geändert am" "" "l" "$t_wchanged" "" "l"
            Leerzeile
        }
    }
}
####################################################################################################
# Funktionen fuer den Bereich "DNS - Check"                                                        #
####################################################################################################
function aging { 
    foreach ($dc in $DCs){
        $globdns = Get-DnsServerScavenging -ComputerName $dc.Name -ErrorAction SilentlyContinue
        if($globdns){
            Bereichstitel $dc.HostName 
            Leerzeile
            if($globdns.ScavengingState -eq $false){
                $gloaus = "nicht aktiviert"
                $glonri = "n/a"
                $glori = "n/a"
                $gloaus_fa = "Red"
            } else {
                $gloaus = "aktiviert"
                $gloaus_fa = "Green"
                [string]$glonri = $globdns.NoRefreshInterval
                [string]$glori = $globdns.RefreshInterval
            }
            $weisrv = ((Get-DnsServerForwarder -ComputerName $dc.Name).IPAddress).IPAddressToString
            Subtitel "Weiterleitungseinstellungen:" "1" "-"
            $n=1
            foreach ($wei in $weisrv){
                [string]$weiip = "- Forwarder - DNS" + "$n"
                new_2werte "s" ":" "23" $weiip "" "l" $wei "" "l"
                $n++
            }
            Leerzeile
            Subtitel "Globale Alterungeinstellungen:" "1" "-"
            new_2werte "s" ":" "23" "- Status" "" "l" $gloaus $gloaus_fa "l"
            new_2werte "s" ":" "23" "- NoRefreshInterval" "" "l" $glonri $gloaus_fa "l"
            new_2werte "s" ":" "23" "- RefreshInterval" "" "l" $glori $gloaus_fa "l"
            Leerzeile
            Subtitel "Zonen Check" "1" "*"
            Leerzeile
            $zonen = Get-DnsServerZone -ComputerName $dc.name | `
            Where-Object {
            $_.IsDSIntegrated -ne $false -and 
            $_.ZoneName -notlike "_*" -and 
            $_.ZoneName -notlike "Trust*" -and
            $_.ZoneType -notlike "Forwarder"
            } | Sort-Object ZoneName -Descending
            foreach ($zone in $zonen){
                $zonalt = Get-DnsServerZoneAging -ComputerName $dc.Name -ZoneName $zone.ZoneName `
                          -ErrorAction SilentlyContinue
                $zoneServers = (Get-DnsServerResourceRecord -ComputerName $dc.Name `
                    -ZoneName $zone.ZoneName -RRType NS | `
                    Where-Object {$_.Hostname -notlike "_ms*"}).RecordData
                if($zonalt.AgingEnabled -eq $false){
                    $zonalt_aus = "nicht aktiviert"
                    $zonalt_fa = "Red"
                    $zonnri = "n/a"
                    $zonri = "n/a"
                } else {
                    $zonalt_aus = "aktiviert"
                    $zonalt_fa = "Green"
                    [string]$zonnri = $zonalt.NoRefreshInterval
                    [string]$zonri = $zonalt.RefreshInterval
                }
                $z_name = "Zone: " + $zone.ZoneName
                Subtitel $z_name "2" "="
                Leerzeile
                Subtitel "DNS-Server der Zone:" "3" "-"
                $i=1
                foreach ($zosrv in $zoneServers){
                    [string]$serv = "  " + $i + ". DNS"
                    new_2werte "s" ":" "23" $serv "" "l" $zosrv.NameServer "" "l"
                    $i++
                }
                Leerzeile
                Subtitel "Alterungsstatus der Zone:" "3" "-"
                new_2werte "s" ":" "23" "  - Status" "" "l" $zonalt_aus $zonalt_fa "l"
                new_2werte "s" ":" "23" "  - NoRefreshInterval" "" "l" $zonnri $zonalt_fa "l"
                new_2werte "s" ":" "23" "  - RefreshInterval" "" "l" $zonri $zonalt_fa "l"
                Leerzeile
            }
        }
    }
}
####################################################################################################
# Funktionen fuer den Bereich "Sysvol Replication & AD-Health"                                     #
####################################################################################################
function dfsr {
    $speak = ([CultureInfo]::InstalledUICulture).Name
    $mstate = $merg = $gstate = $gerg = "!! dfsmig fehlt !!"
    $fa1 = $fa2 = $fa3 = $fa4 = "Red"
    if(Test-Path "C:\Windows\System32\dfsrmig.exe"){
        [string]$dfs1 = dfsrmig /getmigrationstate
        [string]$dfs2 = dfsrmig /GetGlobalState
        if ($speak -eq "de-DE") {
            $erg1a = $dfs1.Contains("Entfernt")
            if ($erg1a -eq "$true") {$mstate = "Entfernt" ; $fa1 = "Green"} else {
                $mstate = "nicht Entfernt" ; $fa1 = "Red" }
            $erg2a = $dfs1.Contains("Erfolgreich")
            if ($erg2a -eq "$true") {$merg = "Erfolgreich" ; $fa2 = "Green"} else {
                $merg = "nicht Erfolgreich" ; $fa2 = "Red" }
            $erg3a = $dfs2.Contains("Entfernt")
            if ($erg3a -eq "$true") {$gstate = "Entfernt" ; $fa3 = "Green"} else {
                $gstate = "nicht Entfernt" ; $fa3 = "Red" }
            $erg4a = $dfs2.Contains("Erfolgreich")
            if ($erg4a -eq "$true") {$gerg = "Erfolgreich" ; $fa4 = "Green"} else {
                $gerg = "nicht Erfolgreich" ; $fa4 = "Red" }
        }
        else {
            $erg1b = $dfs1.Contains("Eliminated")
            if ($erg1b -eq "$true") {$mstate = "Eliminated" ; $fa1 = "Green"} else {
                $mstate = "Not Eliminated" ; $fa1 = "Red" }
            $erg2b = $dfs1.Contains("Succeeded")
            if ($erg2b -eq "$true") {$merg = "Succeeded" ; $fa2 = "Green"} else {
                $merg = "Not Succeeded" ; $fa2 = "Red" }
            $erg3b = $dfs2.Contains("Eliminated")
            if ($erg3b -eq "$true") {$gstate = "Eliminated" ; $fa3 = "Green"} else {
                $gstate = "Not Eliminated" ; $fa3 = "Red" }
            $erg4b = $dfs2.Contains("Succeeded")
            if ($erg4b -eq "$true") {$gerg = "Succeeded" ; $fa4 = "Green"} else {
                $gerg = "Not Succeeded" ; $fa4 = "Red" } 
        }
    }
    Bereichstitel "Migrationsstatus:"
    Leerzeile
    2werte "FS-Replikation     :" $mstate "s" $fa1
    2werte "Umstellung auf DFSR:" $merg "s" $fa2
    Leerzeile
    Bereichstitel "Globaler Status:"
    Leerzeile
    2werte "FS-Replikation     :" $gstate "s" $fa3
    2werte "Umstellung auf DFSR:" $gerg "s" $fa4
    Leerzeile
}
function Get-AllDomainControllers ($DomainNameInput) {
	[array]$allDomainControllers = Get-ADDomainController -Filter * -Server $DomainNameInput
	return $allDomainControllers
}
Function Get-DomainControllerNSLookup($DomainNameInput) {
	try
	{
		$domainControllerNSLookupResult = Resolve-DnsName $DomainNameInput -Type A | Select-Object -ExpandProperty IPAddress
		$domainControllerNSLookupResult = 'Bestanden'
	}
	catch
	{
		$domainControllerNSLookupResult = 'Fehler'
	}
	return $domainControllerNSLookupResult
}
Function Get-DomainControllerPingStatus($DomainNameInput) {
	If ((Test-Connection $DomainNameInput -Count 1 -quiet) -eq $True)
	{
		$domainControllerPingStatus = 'Bestanden'
	}
	Else
	{
		$domainControllerPingStatus = 'Fehler'
	}
	return $domainControllerPingStatus
}
Function Get-DomainControllerUpTime($DomainNameInput) {
	If ((Test-Connection $DomainNameInput -Count 1 -quiet) -eq $True)
	{
		try
		{
			$W32OS = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $DomainNameInput -ErrorAction SilentlyContinue
			$timespan = $W32OS.ConvertToDateTime($W32OS.LocalDateTime) - $W32OS.ConvertToDateTime($W32OS.LastBootUpTime)
			[int]$uptime = "{0:00}" -f $timespan.TotalHours
		}
		catch [exception] {
			$uptime = 'WMI Failure'
		}
		
	}
	Else
	{
		$uptime = '0'
	}
	return $uptime
}
Function Get-DITFileDriveSpace($DomainNameInput) {
	If ((Test-Connection $DomainNameInput -Count 1 -quiet) -eq $True)
	{
		try
		{
			$key = "SYSTEM\CurrentControlSet\Services\NTDS\Parameters"
			$valuename = "DSA Database file"
			$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $DomainNameInput)
			$regkey = $reg.opensubkey($key)
			$NTDSPath = $regkey.getvalue($valuename)
			$NTDSPathDrive = $NTDSPath.ToString().Substring(0, 2)
			$NTDSDiskDrive = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $DomainNameInput -ErrorAction SilentlyContinue | `
				Where-Object { $_.DeviceID -eq $NTDSPathDrive }
			$NTDSPercentFree = [math]::Round($NTDSDiskDrive.FreeSpace / $NTDSDiskDrive.Size * 100)
		}
		catch [exception] {
			$NTDSPercentFree = 'WMI Failure'
		}
	}
	Else
	{
		$NTDSPercentFree = '0'
	}
	return $NTDSPercentFree
}
function Get-DNSService ($DC) {
$session = New-PSSession -ComputerName $DC
$Status_DNS = Invoke-Command -Session $session -ScriptBlock {
    Get-Service -Name DNS | Select-Object -ExpandProperty Status
}
Remove-PSSession  -Session $session
if($Status_DNS.Value -eq 'Running'){
    $zurueck = 'Bestanden'
} else {
    $zurueck = 'Fehler'
}
return $zurueck
$zurueck = $null
}
function Get-NTDSService ($DC) {
$session = New-PSSession -ComputerName $DC
$Status_NTDS = Invoke-Command -Session $session -ScriptBlock {
    Get-Service -Name NTDS | Select-Object -ExpandProperty Status
}
Remove-PSSession  -Session $session
if($Status_NTDS.Value -eq 'Running'){
    $zurueck = 'Bestanden'
} else {
    $zurueck = 'Fehler'
}
return $zurueck
$zurueck = $null
}
function Get-NetlogonService ($DC) {
$session = New-PSSession -ComputerName $DC
$Status_Netlogon = Invoke-Command -Session $session -ScriptBlock {
    Get-Service -Name Netlogon | Select-Object -ExpandProperty Status
}
Remove-PSSession  -Session $session
if($Status_Netlogon.Value -eq 'Running'){
    $zurueck = 'Bestanden'
} else {
    $zurueck = 'Fehler'
}
return $zurueck
$zurueck = $null
}
Function Get-DomainControllerDCDiagTestResults($DomainNameInput) {
	$DCDiagTestResults = New-Object Object
	If ((Test-Connection $DomainNameInput -Count 1 -quiet) -eq $True)
	{
		$DCDiagTest = (Dcdiag.exe /s:$DomainNameInput /test:services /test:FSMOCheck /test:KnowsOfRoleHolders /test:Advertising /test:Replications) -split ('[\r\n]') 
		$DCDiagTestResults | Add-Member -Type NoteProperty -Name "ServerName" -Value $DomainNameInput
		$DCDiagTest | ForEach-Object {
			Switch -RegEx ($_)
			{
				"Starting" { $TestName = ($_ -Replace ".*Starting test: ").Trim() }
				"passed|failed|bestanden|Fehler" {
					If ($_ -Match "passed" -or $_ -match "bestanden")
					{
						$TestStatus = "Bestanden"
					}
					Else
					{
						$TestStatus = "Fehler"
					}
				}
			}
			If ($Null -ne $TestName -And $null -ne $TestStatus)
			{
				$DCDiagTestResults | Add-Member -Name $("$TestName".Trim()) -Value $TestStatus -Type NoteProperty -force
				$TestName = $Null; $TestStatus = $Null
			}
		}
		return $DCDiagTestResults
	}
	Else
	{
		$DCDiagTestResults | Add-Member -Type NoteProperty -Name "ServerName" -Value $DomainNameInput
		$DCDiagTestResults | Add-Member -Name Replications -Value 'Fehler' -Type NoteProperty -force
		$DCDiagTestResults | Add-Member -Name Advertising -Value 'Fehler' -Type NoteProperty -force
		$DCDiagTestResults | Add-Member -Name KnowsOfRoleHolders -Value 'Fehler' -Type NoteProperty -force
		$DCDiagTestResults | Add-Member -Name FSMOCheck -Value 'Fehler' -Type NoteProperty -force
		$DCDiagTestResults | Add-Member -Name Services -Value 'Fehler' -Type NoteProperty -force
	}
	return $DCDiagTestResults
}
Function Get-DomainControllerOSVersion ($DomainNameInput) {
	$W32OSVersion = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $DomainNameInput -ErrorAction SilentlyContinue).Caption
	return $W32OSVersion
}
Function Get-DomainControllerOSDriveFreeSpace ($DomainNameInput) {
	If ((Test-Connection $DomainNameInput -Count 1 -quiet) -eq $True)
	{
		try
		{
			$thisOSDriveLetter = (Get-WmiObject Win32_OperatingSystem -ComputerName $DomainNameInput -ErrorAction SilentlyContinue).SystemDrive
			$thisOSDiskDrive = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $DomainNameInput -ErrorAction SilentlyContinue | `
				Where-Object { $_.DeviceID -eq $thisOSDriveLetter }
			$thisOSPercentFree = [math]::Round($thisOSDiskDrive.FreeSpace / $thisOSDiskDrive.Size * 100)
		}
		catch [exception] {
			$thisOSPercentFree = 'WMI Failure'
		}
	}
	return $thisOSPercentFree
}
function controller_check {
	$allDomains = (Get-ADForest).Domains
	foreach ($domain in $allDomains)
	{
		[array]$allDomainControllers = Get-AllDomainControllers $domain	
		foreach ($domainController in $allDomainControllers)
		{
			$DCDiagTestResults = Get-DomainControllerDCDiagTestResults $domainController.HostName
			Vollzeile
            $s_titel = "Server: " + ($domainController.HostName).ToLower() 
            Bereichstitel $s_titel
            Leerzeile
            $ut1 = (Get-DomainControllerUpTime $domainController.HostName)
            $ut2 = $ut1 / 24
            $ut2 = [math]::Round($ut2, 2)
            $ut3 = "$ut1 Stunden => ca. $ut2 Tage"
			2werte "Site        :" $domainController.Site "s"
			2werte "OS-Version  :" (Get-DomainControllerOSVersion $domainController.hostname) "s"
            2werte "Uptime (hrs):" $ut3 "s"
			Leerzeile
			$roles = $domainController.OperationMasterRoles
            $su_ti1 = "FSMO Rollen Server: " + ($domainController.HostName).ToLower()
            Bereichstitel $su_ti1 "s"
			Leerzeile
			foreach($rol in $roles) {
				2werte "" $rol "s"
			}
			Leerzeile
            Bereichstitel "Netzwerk Prüfung:" "s"
            Leerzeile
            if((Get-DomainControllerNSLookup $domainController.HostName) -eq "Bestanden") { $fdnl = "Green" } 
                else { $fdnl = "Red"}
            if((Get-DomainControllerPingStatus $domainController.HostName) -eq "Bestanden") { $fdnp = "Green" } 
                else { $fdnp = "Red"}
			2werte " DNS Lookup:" (Get-DomainControllerNSLookup $domainController.HostName) "s" $fdnl
			2werte " Ping Test :" (Get-DomainControllerPingStatus $domainController.HostName) "s" $fdnp
            Leerzeile
            Bereichstitel "Belegter Platz NTDS und System Volumen:" "s"
            Leerzeile
            $d_free = (Get-DITFileDriveSpace $domainController.HostName)
            $o_free = (Get-DomainControllerOSDriveFreeSpace $domainController.HostName)
            if($d_free -le 10) { $d_freef = "Red" } elseif ($d_free -le 20) { $d_freef = "Yellow" }
                else { $d_freef = "Green" }
            if($o_free -le 10) { $o_freef = "Red" } elseif ($o_free -le 20) { $o_freef = "Yellow" }
                else { $o_freef = "Green" }
			2werte " NTDS/DIT   - Free Space (%):" $d_free "s" $d_freef
			2werte " OS-Volumen - Free Space (%):" $o_free "s" $o_freef
            Leerzeile
            Bereichstitel "Verfügbarkeit AD-Dienste:" "s"
            Leerzeile
            $stat_DNS = Get-DNSService $domainController.HostName
            if($stat_DNS -eq "Bestanden") { $stat_DNS_f = "Green"} else { $stat_DNS_f = "Red"}
            $stat_NTDS = Get-NTDSService $domainController.HostName
            if($stat_NTDS -eq "Bestanden") { $stat_NTDS_f = "Green"} else { $stat_NTDS_f = "Red"}
            $stat_Netlogon = Get-NetlogonService $domainController.HostName
            if($stat_Netlogon -eq "Bestanden") { $stat_Netlogon_f = "Green"} else { $stat_Netlogon_f = "Red"}
			2werte " DNS Service     :" $stat_DNS "s" $stat_DNS_f
			2werte " NTDS Service    :" $stat_NTDS "s" $stat_NTDS_f
			2werte " NetLogon Service:" $stat_Netlogon "s" $stat_Netlogon_f
			Leerzeile
            Bereichstitel "DC Diagnostic Prüfung:" "s"
            Leerzeile
            if($DCDiagTestResults.Replications -eq "Bestanden") { $dcr = "Green" } 
                else { $dcr = "Red"}
            if($DCDiagTestResults.Advertising -eq "Bestanden") { $dca = "Green" } 
                else { $dca = "Red"}
            if($DCDiagTestResults.KnowsOfRoleHolders -eq "Bestanden") { $dck = "Green" } 
                else { $dck = "Red"}
            if($DCDiagTestResults.FSMOCheck -eq "Bestanden") { $dcf = "Green" } 
                else { $dcf = "Red"}
            if($DCDiagTestResults.Services -eq "Bestanden") { $dcs = "Green" } 
                else { $dcs = "Red"}
			2werte " DCDIAG: Replications           :" $DCDiagTestResults.Replications "s" $dcr
			2werte " DCDIAG: Advertising            :" $DCDiagTestResults.Advertising "s" $dca
			2werte " DCDIAG: FSMO KnowsOfRoleHolders:" $DCDiagTestResults.KnowsOfRoleHolders "s" $dck
			2werte " DCDIAG: FSMO Check             :" $DCDiagTestResults.FSMOCheck "s" $dcf
			2werte " DCDIAG: Services               :" $DCDiagTestResults.Services "s" $dcs
            Leerzeile
            Vollzeile
		}
	}
}
####################################################################################################
# Funktionen fuer den Bereich "Administratoren und Builtin Benutzer"                               #
####################################################################################################
function Admins {
    $dom_sid = (Get-ADDomain).DomainSID ; $Rahmen = @() ; $Gruppe = @()
    $Gruppe += "Domain Admins"
    $Gruppe += "Schema Admins"
    $Gruppe += "Orga Admins"
    $Rahmen += "$dom_sid" + "-512"
    $Rahmen += "$dom_sid" + "-518"
    $Rahmen += "$dom_sid" + "-519"
    $lauf = $Rahmen.Count
    Bereichstitel "Prim. Admin-Gruppen:"
    Leerzeile
    for($i=0;$i -lt $lauf;$i++) {
        $a = (Get-ADGroupMember -Identity $Rahmen[$i] | `
            Where-Object { $_.objectClass -eq "user" }).count
        if($a -eq 0) { $a = "0" } elseif($a -ge 1) {} else { $a = "1" }
        $b = (Get-ADGroupMember -Identity $Rahmen[$i] | `
            Where-Object { $_.objectClass -eq "group" }).count
        if($b -eq 0) { $b = "0" } elseif($b -ge 1) {} else { $b = "1" }
        $c = (Get-ADGroupMember -Identity $Rahmen[$i] -Recursive).count
        if($c -eq 0) { $c = "0" } elseif($c -ge 1) {} else { $c = "1" }
        if($a -ge 10) { $fa1 = "Red" } elseif($a -ge 5) { $fa1 = "Yellow" } else { $fa1 = "Green"}
        if($c -gt $a) { $fa2 = "Yellow" } else { $fa2 = "White" } 
        $titel = "Gruppe der " + $Gruppe[$i] + ":"
        Subtitel $titel "1" "-"
        new_2werte "s" ":" "32" "- Benutzer direkt in der Gruppe" "" "l" "$a" $fa1 "l"
        new_2werte "s" ":" "32" "- Gruppen direkt in der Gruppe" "" "l" "$b" "" "l"
        new_2werte "s" ":" "32" "- Summe der Benutzer der Gruppe" "" "l" "$c" $fa2 "l"
        Leerzeile
    }
}
function lokale_AdmGru {
    $lokgru = @()
    $lokgru += "S-1-5-32-544"  # Administratoren
    $lokgru += "S-1-5-32-548"  # Konten-Operatoren
    $lokgru += "S-1-5-32-549"  # Server-Operatoren
    $lokgru += "S-1-5-32-550"  # Druck-Operatoren
    $lokgru += "S-1-5-32-551"  # Sicherungs-Operatoren
    $lokgru += "S-1-5-32-552"  # Wiederherstellungs-Operatoren
    #$lokgru += "S-1-5-32-554"  # Dokumente- und Einstellungen-Administratoren
    $lokgru += "S-1-5-32-555"  # Terminaldienste-Remotedesktopbenutzer
    $lokgru += "S-1-5-32-556"  # Netzwerkkonfigurations-Operatoren
    $lokgru += "S-1-5-32-557"  # IAS-Server
    $lokgru += "S-1-5-32-569"  # Cryptographische Operatoren
    $lokgru += "S-1-5-32-578"  # Hyper-V-Administratoren
    $lokgru += "S-1-5-32-579"  # Zugriffssteuerungs-Unterstützungsoperatoren
    $lokgru += "S-1-5-32-582"  # Storage Repl. Admin
    Subtitel "Check der lokalen Admin-Gruppen (DCs):"
    Leerzeile
    foreach($log in $lokgru) {
        $log_gruppe = Get-ADGroup -Filter { SID -eq $log }
        if($log_gruppe.Name) {
            $gruppenname = $log_gruppe.Name
            $log_member = Get-ADGroupMember -Identity "$gruppenname"
            $in_a = (Get-ADGroupMember -Identity "$gruppenname" -Recursive).count
            if($in_a -eq 0) { $in_a = "0" } elseif($in_a -ge 1) {} else { $in_a = "1" }
            [string]$inbenu = $in_a
            if($log_member) {
                $be_a = ($log_member | Where-Object {$_.objectClass -eq "user"}).count
                if($be_a -eq 0) { $be_a = "0" } elseif($be_a -ge 1) {} else { $be_a = "1" }
                [string]$benu = $be_a
                $gr_a = ($log_member | Where-Object {$_.objectClass -eq "group"}).count
                if($gr_a -eq 0) { $gr_a = "0" } elseif($gr_a -ge 1) {} else { $gr_a = "1" }
                [string]$grup = $gr_a
                Subtitel $log_gruppe.SamAccountName "1" "-"
                new_2werte "1" ":" "32" "- Benutzer direkt in der Gruppe" "" "l" $benu "" "l"
                new_2werte "1" ":" "32" "- Gruppen direkt in der Gruppe" "" "l" $grup "" "l"
                new_2werte "1" ":" "32" "- Summe der Benutzer der Gruppe" "" "l" $inbenu "" "l"
                Leerzeile
            }
        }
    }
}
function dom_AdmGri {
    $do_sid = (Get-ADDomain).DomainSID
    $admgru = @()
    $admgru += "$do_sid" + "-520" # Group Policy Creator Owners
    $admgru += "$do_sid" + "-526" # Key Admins
    $admgru += "$do_sid" + "-527" # Enterprise Key Admins
    $admgru += "$do_sid" + "-553" # RAS and IAS Servers
    Subtitel "Check weiterer Admin-Gruppen:"
    Leerzeile
    foreach ($adg in $admgru) {
        $dog_gruppe = Get-ADGroup -Filter { SID -eq $adg }
        if ($dog_gruppe.Name) {
            $dog_name = $dog_gruppe.Name
            $dog_member = Get-ADGroupMember -Identity "$dog_name"
            $in_da = (Get-ADGroupMember -Identity "$dog_name" -Recursive).count
            if($in_da -eq 0) { $in_da = "0" } elseif($in_da -ge 1) {} else { $in_da = "1" }
            [string]$in_dog = $in_da
            if($dog_member) {
                $be_a = ($dog_member | Where-Object {$_.objectClass -eq "user"}).count
                if($be_a -eq 0) { $be_a = "0" } elseif($be_a -ge 1) {} else { $be_a = "1" }
                [string]$benu = $be_a
                $gr_a = ($dog_member | Where-Object {$_.objectClass -eq "group"}).count
                if($gr_a -eq 0) { $gr_a = "0" } elseif($gr_a -ge 1) {} else { $gr_a = "1" }
                [string]$grup = $gr_a
                Subtitel $dog_gruppe.SamAccountName "1" "-"
                new_2werte "1" ":" "32" "- Benutzer direkt in der Gruppe" "" "l" $benu "" "l"
                new_2werte "1" ":" "32" "- Gruppen direkt in der Gruppe" "" "l" $grup "" "l"
                new_2werte "1" ":" "32" "- Summe der Benutzer der Gruppe" "" "l" $in_dog "" "l"
                Leerzeile
            }
        }
    }
    
}
function AdmCount {   
    Bereichstitel "Privilegierte Benutzer:"
    Leerzeile
    $ac_usr = Get-ADUser -Filter * -Properties * | Where-Object {$_.adminCount -eq 1} |`
        Sort-Object SamAccountName
    $ac_usr_za = (Get-ADUser -Filter * -Properties * | Where-Object {$_.adminCount -eq 1}).Count
    if($ac_usr_za -lt 5) { $ac_usr_za_fa = "Green" } 
        elseif ($ac_usr_za -gt 10) { $ac_usr_za_fa = "Red" } else {$ac_usr_za_fa = "Yellow" }
    $dom_sid = (Get-ADDomain).DomainSID ; $do_gru = "$dom_sid" + "-512"
    $do_gru_all = Get-ADGroupMember -Identity $do_gru -Recursive
    $do_gru_dir = Get-ADGroupMember -Identity $do_gru
    new_2werte "" ":" "26" "Benutzer mit Admin Count 1" "" "l" "$ac_usr_za" $ac_usr_za_fa "l"
    Leerzeile
    if($ac_usr_za){
        foreach($ac_u in $ac_usr){
            $ac_u_name = "Account Name: " + $ac_u.SamAccountName
            Subtitel $ac_u_name "1" "-"
            new_2werte "" ":" "22" " - User liegt unter" "" "l" $ac_u.CanonicalName "" "l"
            $ac_ud = [datetime]::ParseExact($ac_u.whenCreated, "MM/dd/yyyy HH:mm:ss", `
            [System.Globalization.CultureInfo]::InvariantCulture).ToString("dd.MM.yyyy HH:mm:ss")
            new_2werte "" ":" "22" " - Wurde erstellt am" "" "l" $ac_ud "" "l"
            if($ac_u.Enabled) { $stu = "aktiv"; $fa = "Red" } 
                else { $stu = "deaktiviert"; $fa = "Green" }
            new_2werte "" ":" "22" " - Benutzerstatus" "" "l" $stu $fa "l"
            if ($do_gru_all.SamAccountName -contains $ac_u.SamAccountName) {
                if($do_gru_dir.SamAccountName -contains $ac_u.SamAccountName) {
                    $do_gru_aus = "direktes Mitglied"
                    $do_gru_fa = "Yellow"
                } else {
                    $do_gru_aus = "indirektes Mitglied"
                    $do_gru_fa = "Red"
                }
                new_2werte "" ":" "22" " - MemberOf Dom-Admins" "" "l" $do_gru_aus $do_gru_fa "l"
            } else {
                new_2werte "" ":" "22" " - MemberOf Dom-Admins" "" "l" "nicht Mitglied" "Green" "l"
            }
            Leerzeile
        }
    }
}
function builtin_usr {
    $domain = (Get-ADDomain).DomainSID.value  
    ## Administrator ###############################################################################
    $std_admin = "$domain" + "-500"
    $s_admin = Get-ADUser $std_admin | Select-Object SamAccountName, Name, Enabled
    if($s_admin.Enabled -eq $true) { $s_admin_aktiv = "aktiv" ; $fa1 = "red"} else {
        $s_admin_aktiv = "deaktiviert" ; $fa1 = "Green" }
    ## Gast ########################################################################################
    $std_gast = "$domain" + "-501"
    $s_gast = Get-ADUser $std_gast | Select-Object SamAccountName, Name, Enabled
    if($s_gast.Enabled -eq $true)  { $s_gast_aktiv = "aktiv" ; $fa2 = "red"} else {
        $s_gast_aktiv = "deaktiviert" ; $fa2 = "Green" }
    ## krbtgt ######################################################################################
    $std_krbtgt = "$domain" + "-502"
    $s_krbtgt = Get-ADUser $std_krbtgt | Select-Object SamAccountName, Name, Enabled
    if($s_krbtgt.Enabled -eq $true)  { $s_krbtgt_aktiv = "aktiv" ; $fa3 = "red"} else {
        $s_krbtgt_aktiv = "deaktiviert" ; $fa3 = "Green" }
    ## DefaultAccount ##############################################################################
    $std_default = "$domain" + "-503"
    if(Get-ADUser -Filter {ObjectSid -eq $std_default}) {
        $s_default = Get-ADUser $std_default | Select-Object SamAccountName, Name, Enabled
        if($s_default) { if($s_default.Enabled -eq $true) { $s_default_aktiv = "aktiv" ; $fa4 = "red"} 
                else { $s_default_aktiv = "deaktiviert" ; $fa4 = "Green" } } }
    ## Ausgabe #####################################################################################
    Bereichstitel "Standard Builtin Accounts:" ""
    Leerzeile
    Subtitel "Zum Administrator Account:" "1" "-"
    new_2werte "1" ":" "16" "- SamAccountName" "" "l" $s_admin.SamAccountName "" "l"
    new_2werte "1" ":" "16" "- Name" "" "l" $s_admin.Name "" "l"
    new_2werte "1" ":" "16" "- Account Status" "" "l" $s_admin_aktiv $fa1 "l"
    Leerzeile
    Subtitel "Zum Gast Account:" "1" "-"
    new_2werte "1" ":" "16" "- SamAccountName" "" "l" $s_gast.SamAccountName "" "l"
    new_2werte "1" ":" "16" "- Name" "" "l" $s_gast.Name "" "l"
    new_2werte "1" ":" "16" "- Account Status" "" "l" $s_gast_aktiv $fa2 "l"
    Leerzeile
    Subtitel "Zum krbtgt Account:" "1" "-"
    new_2werte "1" ":" "16" "- SamAccountName" "" "l" $s_krbtgt.SamAccountName "" "l"
    new_2werte "1" ":" "16" "- Name" "" "l" $s_krbtgt.Name "" "l"
    new_2werte "1" ":" "16" "- Account Status" "" "l" $s_krbtgt_aktiv $fa3 "l"
    Leerzeile
    if($s_default) {
        Subtitel "Zum Default Account:" "1" "-"
        new_2werte "1" ":" "16" "- SamAccountName" "" "l" $s_default.SamAccountName "" "l"
        new_2werte "1" ":" "16" "- Name" "" "l" $s_default.Name "" "l"
        new_2werte "1" ":" "16" "- Account Status" "" "l" $s_default_aktiv $fa4 "l"
        Leerzeile
    }
}
####################################################################################################
# Funktionen fuer den Bereich "Benutzer und Benutzer Accounts"                                     #
####################################################################################################
function User_chk {
    Bereichstitel "Benutzer Accounts allgemein:" ""
    Leerzeile
    Subtitel 'Auf-/Verteilung:' "1" "-"
    ## Alle ##
    $all_usr = (Get-ADUser -Filter * | Select-Object name).count
    new_2werte "s" ":" "36" " Summe der Benutzer Accounts" "" "l" "$all_usr" "" "l"
    new_2werte "s" "" "36" " -----------------------------------" "" "l" "----" "" "l"
    ## Aktivierte ##
    $akt_usr = Get-ADUser -Filter * -Properties Name,SamAccountName,LastLogonDate | Where-Object {$_.Enabled -eq $true}
    [int]$akt_usr_c = 0 ; [int]$akt_usr_oan_c = 0
    foreach($akt in $akt_usr) { if($null -eq $akt.LastLogonDate) { $akt_usr_oan_c++ } $akt_usr_c++ }
    [string]$akt_usr_oan = "$akt_usr_oan_c" + "   (Admin Count 1 beachten!)"
    new_2werte "s" ":" "36" "  - Aktivierte Benutzer Accounts" "" "l" "$akt_usr_c" "" "l"
    new_2werte "s" ":" "36" "    - Davon Accounts ohne Anmeldung" "" "l" "$akt_usr_oan" "" "l"
    new_2werte "s" "" "36" "  ----------------------------------" "" "l" "----" "" "l"
    ## Deaktivierte ##
    $dea_usr = Get-ADUser -Filter * -Properties Name,SamAccountName,LastLogonDate | Where-Object {$_.Enabled -eq $false}
    [int]$dea_usr_c = 0 ; [int]$dea_usr_oan_c = 0
    foreach($dea in $dea_usr) { if($null -eq $dea.LastLogonDate) { $dea_usr_oan_c++ } $dea_usr_c++ }
    [string]$dea_usr_oan = "$dea_usr_oan_c" + "   (Admin Count 1 beachten!)"
    new_2werte "s" ":" "36" "  - Deaktivierte Benutzer Accounts" "" "l" "$dea_usr_c" "" "l"
    new_2werte "s" ":" "36" "    - Davon Accounts ohne Anmeldung" "" "l" "$dea_usr_oan" "" "l"
    Leerzeile
}
function inaktive_User {
    $epoch = New-Object System.DateTime(1601, 1, 1)
    $zeitraum = (Get-Date).AddMonths(-2)
    $inaktive = Get-ADUser -Filter {LastLogonTimeStamp -lt $zeitraum} `
                -Properties SamAccountName,Name,LastLogonTimeStamp,Enabled | `
                Sort-Object LastLogonTimeStamp
    $ina_anza = $inaktive.Count
    if($ina_anza -eq 0 -or $null -eq $ina_anza) { $ina_anza = "0" } elseif($ina_anza -gt 1) { } 
        else { $ina_anza = "1" }
    if($ina_anza -ne 0) {
        Bereichstitel "Inaktive Benutzer Accounts:" ""
        Leerzeile
        Subtitel "Benutzer (inaktiv), letzte Anmeldung über 2 Monate her:" "1" "-"
        new_2werte "s" ":" "31" "- Anzahl der inaktiven Benutzer" "" "l" "$ina_anza" "" "l"
        Leerzeile
        neu_tab_max6w_fb "3" "l" "s" "17" "Name" "SamAccountName" "Letzte Anmeldung" "Account Status"
        tablinie "s"
        foreach ($ina in $inaktive) 
        {
            $in_Name = $ina.Name
            $in_san = $ina.SamAccountName
            $llts = $ina.LastLogonTimeStamp
            if($ina.Enabled -eq $false) { $in_stat = "deaktiviert" ; $in_stat_f = "Green" }
                else { $in_stat = "aktiviert" ; $in_stat_f = "Red" }
            $lastLogon = ($epoch.AddSeconds($llts / (1e7))).ToString("dd.MM.yyyy HH:mm")
            neu_tab_max6w_fb "3" "l" "s" "17" $in_Name $in_san $lastLogon $in_stat "" "" $in_stat_f
        }
        Leerzeile
    }
}
function gesperrte_User {
    $gusr = Search-ADAccount -LockedOut
    $gusr_an = $gusr.Count
    if($gusr_an -eq 0 -or $null -eq $gusr_an) { $gusr_an = "0" } 
        elseif($gusr_an -gt 1) { } else { $gusr_an = "1" }
    if($null -ne $gusr){
        Bereichstitel "Gesperrte Benutzer Accounts:" ""
        Leerzeile
        Subtitel "Betroffene Benutzer Accounts:" "1" "-"
        new_2werte "s" ":" "41" "- Anzahl der aktuell gesperrten Benutzer" "" "l" "$gusr_an" "" "l"
        Leerzeile
        neu_tab_max6w_fb "3" "l" "s" "17" "Name" "SamAccountName" "Last Logon Try" "Fehlversuche"
        tablinie "s"
        foreach($gus in $gusr) {
            [string]$llotry = ((Get-ADUser -Identity $gus -Properties *).LastBadPasswordAttempt).ToString("dd.MM.yy HH:mm:ss")
            [string]$g_lout = (Get-ADUser -Identity $gus -Properties *).badPwdCount
            $g_lout = $g_lout.PadLeft(2)
            $g_name = $gus.Name
            $g_sama = $gus.SamAccountName
            neu_tab_max6w_fb "3" "l" "s" "17" "$g_name" "$g_sama" "$llotry" "$g_lout"
        }
    Leerzeile
    }
}
function ou_users {
    $dom = (Get-ADDomain).DistinguishedName ; $ou = "CN=Users,"+"$dom"
    $domain = (Get-ADDomain).DomainSID.value ; $stds = @()
    ################################################################################################
    $stds += "$domain" + "-500" # Administrator
    $stds += "$domain" + "-501" # Gast
    $stds += "$domain" + "-502" # krbtgt
    ################################################################################################
    $ou_all = (Get-ADUser -Filter * -SearchBase $ou).Count
    $ou_usr = (Get-ADUser -Filter * -SearchBase $ou | Where-Object {$stds -notcontains $_.SID}) | `
        Select-Object SID, SamAccountName,Name,LastLogonDate,Enabled
    $ou_usr_c = ($ou_usr.SID).Count
    if($ou_usr_c -gt 0){
        Bereichstitel 'Benutzer Accounts in der OU "User":' ""
        Leerzeile
        Subtitel "Aufteilung der Accounts:" "1" "-"
        $std_usr_c = $stds.Count
        new_2werte "s" ":" "36" " Summe aller Benutzer Accounts" "" "l" "$ou_all" "" "l"
        new_2werte "s" "" "36" " -----------------------------------" "" "l" "----" "" "l"
        new_2werte "s" ":" "36" " - Davon Standard Accounts" "" "l" "$std_usr_c" "" "l"
        new_2werte "s" ":" "36" " - Davon Sonstige Accounts" "" "l" "$ou_usr_c" "" "l"
        Leerzeile
        neu_tab_max6w_fb "3" "l" "s" "17" "Name" "SamAccountName" "Letzte Anmeldung" "Account Status"
        tablinie "s"
        foreach($ous in $ou_usr) {
            $t_nam = $ous.Name
            $t_san = $ous.SamAccountName
            if($ous.LastLogonDate) { $t_lld = ($ous.LastLogonDate).ToString("dd.MM.yyyy HH:mm") }
                else { $t_lld = " - n.a. - " }
            #$t_lld = ($ous.LastLogonDate).ToString("dd.MM.yyyy HH:mm")
            if($ous.Enabled -eq $false) { $t_sta = "deaktiviert" ; $t_sta_f = "Green" }
                else { $t_sta = "aktiviert" ; $t_sta_f = "yellow" } 
            neu_tab_max6w_fb "3" "l" "s" "17" "$t_nam" "$t_san" "$t_lld" "$t_sta" "" "" $t_sta_f
        }
    }
    Leerzeile
}
####################################################################################################
# Funktionen fuer den Bereich "Clients-/Serverkonten und nicht Windows Systeme"                    #
####################################################################################################
function sys_konten {
    $s_base = (Get-ADDomain).DistinguishedName
    $sum_konten = (Get-ADComputer -Filter * -SearchBase $s_base).count
    if($sum_konten -eq 0) { $sum_konten = "0" } elseif($sum_konten -gt 1) { } else { $sum_konten = "1" }    
    $akt_konten = (Get-ADComputer -Filter * -SearchBase $s_base | `
        Where-Object { $_.enabled -eq $true} ).count
    if($akt_konten -eq 0) { $akt_konten = "0" } elseif($akt_konten -gt 1) { } else { $akt_konten = "1" }
    $dea_konten = (Get-ADComputer -Filter * -SearchBase $s_base | `
        Where-Object { $_.enabled -ne $true} ).count
    if($dea_konten -eq 0) { $dea_konten = "0" } elseif($dea_konten -gt 1) { } else { $dea_konten = "1" }
    $sum_clt_konten = (Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object {
            $_.operatingsystem -like 'windows 2*' -or
            $_.operatingsystem -like 'windows X*' -or
            $_.operatingsystem -like 'windows 7*' -or
            $_.operatingsystem -like 'windows 8*' -or
            $_.operatingsystem -like 'windows 10*' -or
            $_.operatingsystem -like 'windows 11*'
        } ).count
    if($sum_clt_konten -eq 0) { $sum_clt_konten = "0" } elseif($sum_clt_konten -gt 1) { } else { $sum_clt_konten = "1" }
    $akt_clt_konten = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { 
            $_.operatingsystem -like 'windows 2*' -or
            $_.operatingsystem -like 'windows X*' -or
            $_.operatingsystem -like 'windows 7*' -or
            $_.operatingsystem -like 'windows 8*' -or
            $_.operatingsystem -like 'windows 10*' -or
            $_.operatingsystem -like 'windows 11*'
        } ) | Where-Object {$_.Enabled -eq $true}).count
    if($akt_clt_konten -eq 0) { $akt_clt_konten = "0" } elseif($akt_clt_konten -gt 1) { } else { $akt_clt_konten = "1" }
    $dea_clt_2 = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows 2*' }) | Where-Object {$_.Enabled -ne $true}).count
    if($dea_clt_2 -eq 0) { $dea_clt_2 = "0" } elseif($dea_clt_2 -gt 1) { } else { $dea_clt_2 = "1" }
    [int]$dea_clt_konten = $dea_clt_2
    $dea_clt_x = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows X*' }) | Where-Object {$_.Enabled -ne $true}).count
    if($dea_clt_x -eq 0) { $dea_clt_x = "0" } elseif($dea_clt_x -gt 1) { } else { $dea_clt_x = "1" }
    [int]$dea_clt_konten = $dea_clt_konten + $dea_clt_x
    $dea_clt_7 = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows 7*' }) | Where-Object {$_.Enabled -ne $true}).count
    if($dea_clt_7 -eq 0) { $dea_clt_7 = "0" } elseif($dea_clt_7 -gt 1) { } else { $dea_clt_7 = "1" }
    [int]$dea_clt_konten = $dea_clt_konten + $dea_clt_7
    $dea_clt_8 = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows 8*' }) | Where-Object {$_.Enabled -ne $true}).count
    if($dea_clt_8 -eq 0) { $dea_clt_8 = "0" } elseif($dea_clt_8 -gt 1) { } else { $dea_clt_8 = "1" }
    [int]$dea_clt_konten = $dea_clt_konten + $dea_clt_8
    $dea_clt_10 = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows 10*' }) | Where-Object {$_.Enabled -ne $true}).count
    if($dea_clt_10 -eq 0) { $dea_clt_10 = "0" } elseif($dea_clt_10 -gt 1) { } else { $dea_clt_10 = "1" }
    [int]$dea_clt_konten = $dea_clt_konten + $dea_clt_10
    $dea_clt_11 = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows 11*' }) | Where-Object {$_.Enabled -ne $true}).count
    if($dea_clt_11 -eq 0) { $dea_clt_11 = "0" } elseif($dea_clt_11 -gt 1) { } else { $dea_clt_11 = "1" }
    [int]$dea_clt_konten = $dea_clt_konten + $dea_clt_11
    $sum_srv_konten = (Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows server 2*' } ).count
    if($sum_srv_konten -eq 0) { $sum_srv_konten = "0" } elseif($sum_srv_konten -gt 1) { } else { $sum_srv_konten = "1" }
    $akt_srv_konten = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows server 2*' } ) | `
        Where-Object { $_.enabled -eq $true} ).count
    if($akt_srv_konten -eq 0) { $akt_srv_konten = "0" } elseif($akt_srv_konten -gt 1) { } else { $akt_srv_konten = "1" }
    $dea_srv_konten = ((Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -like 'windows server 2*' } ) | `
        Where-Object { $_.enabled -ne $true} ).count
    if($dea_srv_konten -eq 0) { $dea_srv_konten = "0" } elseif($dea_srv_konten -gt 1) { } else { $dea_srv_konten = "1" }
    $sum_oth_konten = (Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -notlike 'windows*' } ).count
    if($sum_oth_konten -eq 0) { $sum_oth_konten = "0" } elseif($sum_oth_konten -gt 1) { } else { $sum_oth_konten = "1" }
    $akt_oth_konten = (Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -notlike 'windows*' } | `
        Where-Object { $_.enabled -eq $true} ).count
    if($akt_oth_konten -eq 0) { $akt_oth_konten = "0" } elseif($akt_oth_konten -gt 1) { } else { $akt_oth_konten = "1" }
    $dea_oth_konten = (Get-ADComputer -Filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -notlike 'windows*' } | `
        Where-Object { $_.enabled -ne $true} ).count
    if($dea_oth_konten -eq 0) { $dea_oth_konten = "0" } elseif($dea_oth_konten -gt 1) { } else { $dea_oth_konten = "1" }
    Bereichstitel "Computer- und Systemkonten allgemein:"
    Leerzeile
    Subtitel "Summe der Computer-/Systemkonten:" "2" "-"
    new_2werte "s" ":" "34" " - Summe aller Konten" "" "l" "$sum_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe aktiv. Konten" "" "l" "$akt_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe deakt. Konten" "" "l" "$dea_konten" "" "l"
    Leerzeile
    Subtitel "Windows Client Konten:" "2" "-"
    new_2werte "s" ":" "34" " - Summe aller Client Konten" "" "l" "$sum_clt_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe aktiv. Client Konten" "" "l" "$akt_clt_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe deakt. Client Konten" "" "l" "$dea_clt_konten" "" "l"
    Leerzeile
    Subtitel "Windows Server Konten:" "2" "-"
    new_2werte "s" ":" "34" " - Summe aller Server Konten" "" "l" "$sum_srv_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe aktiv. Server Konten" "" "l" "$akt_srv_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe deakt. Server Konten" "" "l" "$dea_srv_konten" "" "l"
    Leerzeile
    Subtitel 'Nicht "Windows" Betriebssystem Konten:' "2" "-"
    new_2werte "s" ":" "34" " - Summe der nicht Windows Konten" "" "l" "$sum_oth_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe aktiv. nicht Win. Konten" "" "l" "$akt_oth_konten" "" "l"
    new_2werte "s" ":" "34" " - Summe deakt. nicht Win. Konten" "" "l" "$dea_oth_konten" "" "l"
    Leerzeile
}
####################################################################################################
# Funktionen fuer den Bereich "Client Check"                                                       #
####################################################################################################
function clt_chk {
    $zaehler_c = 0
    ##### Windows XP ###############################################################################
    $winxp = (Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows X*' }).count
    if($winxp -eq 0) { $winxp = "0" ; $fxp = "Green" } elseif($winxp -gt 1) { $fxp = "Red" } 
        else { $winxp = "1" ; $fxp = "Red" }
    $winxpko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows X*' }
    $zaehler_c = $zaehler_c + $winxp
    ##### Windows 7 ################################################################################
    $win7 = (Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 7*' }).count
    if($win7 -eq 0) { $win7 = "0" ; $f7 = "Green" } elseif($win7 -gt 1) { $f7 = "Red" } 
        else { $win7 = "1" ; $f7 = "Red" }
    $win7ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 7*' }
    $zaehler_c = $zaehler_c + $win7
    ##### Windows 8 ################################################################################
    $win8 = (Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 8*' }).count
    if($win8 -eq 0) { $win8 = "0" ; $f8 = "Green" } elseif($win8 -gt 1) { $f8 = "Red" } 
        else { $win8 = "1" ; $f8 = "Red" }
    $win8ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 8*' }
    $zaehler_c = $zaehler_c + $win8
    ##### Windows 10 ###############################################################################
    $win10 = (Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 10*' }).count
    if($win10 -eq 0) { $win10 = "0" ; $f10 = "Green" } elseif($win10 -gt 1) { $f10 = "Yellow" } 
        else { $win10 = "1" ; $f10 = "Yellow" }
    $win10ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 10*' } | Sort-Object OperatingSystemVersion
    $zaehler_c = $zaehler_c + $win10
    ##### Windows 11 ###############################################################################
    $win11 = (Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 11*' }).count
    if($win11 -eq 0) { $win11 = "0" ; $f11 = "Yellow" } elseif($win11 -gt 1) { $f11 = "Green" } 
        else { $win11 = "1" ; $f11 = "Green" }
    $win11ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows 11*' } | Sort-Object OperatingSystemVersion
    $zaehler_c = $zaehler_c + $win11
    ##### Zusammenfassung ##########################################################################
    Bereichstitel "Clients nach OS:"
    Leerzeile
    2werte " Windows XP          :" $winxp "s" $fxp
    2werte " Windows 7           :" $win7 "s" $f7
    2werte " Windows 8 bzw. 8.1  :" $win8 "s" $f8
    2werte " Windows 10          :" $win10 "s" $f10
    2werte " Windows 11          :" $win11 "s" $f11
    Leerzeile
    if ($zaehler_c -ne 0){
        if ($cltchk -eq 2) { Bereichstitel "Liste der Clients nach OS:" } else {
            Bereichstitel "Liste veralteter Clients nach OS:" }
        Leerzeile
        Subtitel "Legende: Systeme Out of Service" "1" "-"
        new_2werte "s" ":" "20" "- OS End of Life" "" "l" "Rot" "Red" "l"
        new_2werte "s" ":" "20" "- Systemstatus EoL" "" "l" "aktiviert" "Red" "l"
        new_2werte "s" ":" "20" "- Systemstatus Eol" "" "l" "deaktiviert" "Green" "l"
        Leerzeile
        Subtitel "Legende: Systeme supported" "1" "-"
        new_2werte "s" ":" "20" "- Systemstatus" "" "l" "aktiviert" "Green" "l"
        new_2werte "s" ":" "20" "- Systemstatus" "" "l" "deaktiviert" "Yellow" "l"
        Leerzeile
        neu_tab_max6w_fb "3" "l" "s" "15" "Betriebssystem" "Systemname" "IPv4-Adresse" "EoS"
        tablinie
        if ($winxp -gt "0") { 
            foreach ($xpko in $winxpko) {
                $ena_xp = $xpko.Enabled 
                if($ena_xp -eq $true) {$ena_x_fa = "Red"} else {$ena_x_fa = "Green"}
                $end_xp = "08.04.2014"
                neu_tab_max6w_fb "3" "l" "s" "15" $xpko.OperatingSystem $xpko.Name $xpko.IPv4Address $end_xp $ena_x_fa "" "Red"
            }
        tablinie
        }
        if ($win7 -gt "0") { 
            foreach ($7ko in $win7ko) { 
                $ena_7 = $7ko.Enabled
                if($ena_7 -eq $true) {$ena_7_fa = "Red"} else {$ena_7_fa = "Green"}
                $end_7 = "14.01.2020"
                neu_tab_max6w_fb "3" "l" "s" "15" $7ko.OperatingSystem $7ko.Name $7ko.IPv4Address $end_7 $ena_7_fa "" "Red"
            }
        tablinie
        }
        if ($win8 -gt "0") { 
            foreach ($8ko in $win8ko) { 
                $ena_8 = $8ko.Enabled
                if($ena_8 -eq $true) {$ena_8_fa = "Red"} else {$ena_8_fa = "Green"}
                $end_8 = "10.01.2023"
                neu_tab_max6w_fb "3" "l" "s" "15" $8ko.OperatingSystem $8ko.Name $8ko.IPv4Address $end_8 $ena_8_fa "" "Red"
            }
        tablinie
        }
        if ($win10 -gt "0") { 
            foreach ($10ko in $win10ko) { 
                $buf_1 = (Get-ADComputer -Identity $10ko -Properties *).OperatingSystemVersion
                $buf_1 = (($buf_1.Replace(")","")).Split("("))[1]
                if ($buf_1 -lt "19044") { 
                    $sys_1 = "vor Windows 10 21H2"
                    $EoS_1 = "End of Life" ; $EoS_1_fa = "Red"
                    if($10ko.Enabled -eq $true) {$ena_10_fa = "Red"} else { $ena_10_fa = "Green" }
                    neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $10ko.Name $10ko.IPv4Address $EoS_1 $ena_10_fa "" $EoS_1_fa
                }
                if ($buf_1 -eq "19044") { 
                    $sys_1 = "Windows 10 21H2"
                    $EoS_1 = "11.06.2024" ; $EoS_1_fa = "Green"
                    if($10ko.Enabled -eq $true) {$ena_10_fa = "Green"} else { $ena_10_fa = "Yellow" }
                    neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $10ko.Name $10ko.IPv4Address $EoS_1 $ena_10_fa "" $EoS_1_fa
                }
                if ($cltchk -eq 2) {
                    if ($buf_1 -eq "19045") {
                        $sys_1 = "Windows 10 22H2"
                        $EoS_1 = "14.10.2025" ; $EoS_1_fa = "Green"
                        if($10ko.Enabled -eq $true) {$ena_10_fa = "Green"} else { $ena_10_fa = "Yellow" }
                        neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $10ko.Name $10ko.IPv4Address $EoS_1 $ena_10_fa "" $EoS_1_fa
                    }
                    if ($buf_1 -gt "19045") {
                        $sys_1 = "neuer als Basis 22H2"
                        $EoS_1 = "14.10.2025" ; $EoS_1_fa = "Green"
                        if($10ko.Enabled -eq $true) {$ena_10_fa = "Green"} else { $ena_10_fa = "Yellow" }
                        neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $10ko.Name $10ko.IPv4Address $EoS_1 $ena_10_fa "" $EoS_1_fa
                    }
                }
            }
        tablinie
        }
        if ($win11 -gt "0") { 
            foreach ($11ko in $win11ko) { 
                $buf_2 = (Get-ADComputer -Identity $11ko -Properties *).OperatingSystemVersion
                $buf_2 = (($buf_2.Replace(")","")).Split("("))[1]
                #$ena_11 = $11ko.Enabled
                if ($buf_2 -lt "22000") { 
                    $sys_1 = "vor Windows 11 21H2"
                    $EoS_1 = "End of Life" ; $EoS_1_fa = "Red"
                    if($11ko.Enabled -eq $true) {$ena_11_fa = "Red"} else {$ena_11_fa = "Green"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $11ko.Name $11ko.IPv4Address $EoS_1 $ena_11_fa "" $EoS_1_fa
                }
                if ($buf_2 -eq "22000") { 
                    $sys_1 = "Windows 11 21H2"
                    $EoS_1 = "08.10.2024" ; $EoS_1_fa = "Green"
                    if($11ko.Enabled -eq $true) {$ena_11_fa = "Green"} else {$ena_11_fa = "Red"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $11ko.Name $11ko.IPv4Address $EoS_1 $ena_11_fa "" $EoS_1_fa
                }
                if ($cltchk -eq 2) {
                    if ($buf_2 -eq "22621") { 
                        $sys_1 = "Windows 11 22H2"
                        $EoS_1 = "14.10.2025" ; $EoS_1_fa = "Green"
                        if($11ko.Enabled -eq $true) {$ena_11_fa = "Green"} else {$ena_11_fa = "Red"}
                        neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $11ko.Name $11ko.IPv4Address $EoS_1 $ena_11_fa "" $EoS_1_fa
                    }
                    if ($buf_2 -gt "22621") { 
                        $sys_1 = "neuer als Basis 22H2"
                        $EoS_1 = "14.10.2025" ; $EoS_1_fa = "Green"
                        if($11ko.Enabled -eq $true) {$ena_11_fa = "Green"} else {$ena_11_fa = "Red"}
                        neu_tab_max6w_fb "3" "l" "s" "15" $sys_1 $11ko.Name $11ko.IPv4Address $EoS_1 $ena_11_fa "" $EoS_1_fa
                    }
                }
            }
        }
        Leerzeile
    }
}
####################################################################################################
# Funktionen fuer den Bereich "Server Check"                                                       #
####################################################################################################
function srv_chk {
    $zaehler_s = 0
    ##### Server 2000 ##############################################################################
    $2000 = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2000' }).Count
    if($2000 -eq 0) { $2000 = "0" ; $f2k = "Green" } elseif($2000 -gt 1) { $f2k = "Red" } 
        else { $2000 = "1" ; $f2k = "Red" }
    $2000ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows server 2000' }
    $zaehler_s = $zaehler_s + $2000
    ##### Server 2003 ##############################################################################
    $2003 = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2003' }).Count
    if($2003 -eq 0) { $2003 = "0" ; $f2k3 = "Green" } elseif($2003 -gt 1) { $f2k3 = "Red" } 
        else { $2003 = "1" ; $f2k3 = "Red" }
    $2003ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows server 2003' }
    $zaehler_s = $zaehler_s + $2003
    ##### Server 2008 ##############################################################################
    $2008 = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2008*' }).Count
    $2008r = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2008 R2*' }).Count
    $2008 = [int]$2008 - [int]$2008r
    if($2008 -eq 0) { $2008 = "0" ; $f2k8 = "Green" } elseif($2008 -gt 1) { $f2k8 = "Red" } 
        else { $2008 = "1" ; $f2k8 = "Red" }
    if($2008r -eq 0) { $2008r = "0" ; $f2k8r = "Green" } elseif($2008r -gt 1) { $f2k8r = "Red" } 
        else { $2008r = "1" ; $f2k8r = "Red" }
    $2008ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows server 2008*' } | Sort-Object operatingsystem
    $zaehler_s = $zaehler_s + $2008r + $2008
    ##### Server 2012 ##############################################################################
    $2012 = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2012*' }).Count
    $2012r = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2012 R2*' }).Count
    $2012 = [int]$2012 - [int]$2012r
    if($2012 -eq 0) { $2012 = "0" ; $f2k12 = "Green" } elseif($2012 -gt 1) { $f2k12 = "Red" } 
        else { $2012 = "1" ; $f2k12 = "Red" }
    if($2012r -eq 0) { $2012r = "0" ; $f2k12r = "Green" } elseif($2012r -gt 1) { $f2k12r = "Red" } 
        else { $2012r = "1" ; $f2k12r = "Red" }
    $2012ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows server 2012*' } | Sort-Object operatingsystem
    $zaehler_s = $zaehler_s + $2012
    ##### Server 2016 ##############################################################################
    $2016 = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2016*' }).Count
    if($2016 -eq 0) { $2016 = "0" ; $f2k16 = "Green" } elseif($2016 -gt 1) { $f2k16 = "Yellow" } 
        else { $2016 = "1" ; $f2k16 = "Yellow" }
    $2016ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows server 2016*' }
    $zaehler_s = $zaehler_s + $2016
    ##### Server 2019 ##############################################################################
    $2019 = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2019*' }).Count
    if($2019 -eq 0) { $2019 = "0" ; $f2k19 = "Green" } elseif($2019 -gt 1) { $f2k19 = "Green" } 
        else { $2019 = "1" ; $f2k19 = "Green" }
    $2019ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows server 2019*' }
    $zaehler_s = $zaehler_s + $2019
    ##### Server 2022 ##############################################################################
    $2022 = @(Get-ADComputer -filter * -Properties OperatingSystem | `
        Where-Object {$_.operatingsystem -like 'windows server 2022*' }).Count
    if($2022 -eq 0) { $2022 = "0" ; $f2k22 = "Green" } elseif($2022 -gt 1) { $f2k22 = "Green" } 
        else { $2022 = "1" ; $f2k22 = "Green" }
    $2022ko = Get-ADComputer -filter * -Properties * | `
        Where-Object {$_.operatingsystem -like 'windows server 2022*' }
    $zaehler_s = $zaehler_s + $2022
    ##### Zusammenfassung ##########################################################################
    Bereichstitel "Server nach OS:"
    Leerzeile
    2werte " Windows Server 2000    :" $2000 "s" $f2k
    2werte " Windows Server 2003    :" $2003 "s" $f2k3
    2werte " Windows Server 2008    :" $2008 "s" $f2k8
    2werte " Windows Server 2008 R2 :" $2008r "s" $f2k8r
    2werte " Windows Server 2012    :" $2012 "s" $f2k12
    2werte " Windows Server 2012 R2 :" $2012r "s" $f2k12r
    2werte " Windows Server 2016    :" $2016 "s" $f2k16
    2werte " Windows Server 2019    :" $2019 "s" $f2k19
    2werte " Windows Server 2022    :" $2022 "s" $f2k22
    Leerzeile
    if ($zaehler_s -ne 0) {
        if ($srvchk -eq 2) { Bereichstitel "Liste der Server nach OS:" } else {
            Bereichstitel "Liste veralteter Server nach OS:" }
        Leerzeile
        Subtitel "Legende: Systeme Out of Service" "1" "-"
        new_2werte "s" ":" "20" "- OS End of Life" "" "l" "Rot" "Red" "l"
        new_2werte "s" ":" "20" "- Systemstatus EoL" "" "l" "aktiviert" "Red" "l"
        new_2werte "s" ":" "20" "- Systemstatus Eol" "" "l" "deaktiviert" "Green" "l"
        Leerzeile
        Subtitel "Legende: Systeme supported" "1" "-"
        new_2werte "s" ":" "20" "- Systemstatus" "" "l" "aktiviert" "Green" "l"
        new_2werte "s" ":" "20" "- Systemstatus" "" "l" "deaktiviert" "Yellow" "l"
        new_2werte "s" ":" "20" "- Extended Support" "" "l" "Enddatum" "Yellow" "l"
        Leerzeile
        neu_tab_max6w_fb "3" "l" "s" "15" "Betriebssystem" "Systemname" "IPv4-Adresse" "EoS"
        tablinie
        if ($2000 -gt "0") { 
            foreach ($0ko in $2000ko) { 
                $EoS_0 = "End of Life" ; $EoS_0_fa = "Red"
                $ena_0 = $0ko.Enabled
                if($ena_0 -eq $true) {$ena_0_fa = "Red"} else {$ena_0_fa = "Green"}
                neu_tab_max6w_fb "3" "l" "s" "15" $0ko.OperatingSystem $0ko.Name $0ko.IPv4Address $EoS_0 $ena_0_fa "" $EoS_0_fa
            }
            tablinie
        }
        if ($2003 -gt "0") { 
            foreach ($3ko in $2003ko) { 
                $EoS_3 = "End of Life" ; $EoS_3_fa = "Red"
                $ena_3 = $3ko.Enabled
                if($ena_3 -eq $true) {$ena_3_fa = "Red"} else {$ena_3_fa = "Green"}
                neu_tab_max6w_fb "3" "l" "s" "15" $3ko.OperatingSystem $3ko.Name $3ko.IPv4Address $EoS_3 $ena_3_fa "" $EoS_3_fa
            }
            tablinie
        }
        if ($2008 -gt "0" -or $2008r -gt "0") { 
            foreach ($8ko in $2008ko) {
                if($8ko.OperatingSystem -notlike "*R2*")
                {
                    $EoS_8 = "14.01.2020" ; $EoS_8_fa = "Red"
                    $ena_8 = $8ko.Enabled
                    if($ena_8 -eq $true) {$ena_8_fa = "Red"} else {$ena_8_fa = "Green"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $8ko.OperatingSystem $8ko.Name $8ko.IPv4Address $EoS_8 $ena_8_fa "" $EoS_8_fa
                }
            }
            tablinie
            foreach ($8ko in $2008ko) {
                if($8ko.OperatingSystem -like "*R2*")
                {
                    $EoS_8 = "14.01.2020 (R2)" ; $EoS_8_fa = "Red"
                    $ena_8 = $8ko.Enabled
                    if($ena_8 -eq $true) {$ena_8_fa = "Red"} else {$ena_8_fa = "Green"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $8ko.OperatingSystem $8ko.Name $8kor.IPv4Address $EoS_8 $ena_8_fa "" $EoS_8_fa
                }
            }
            tablinie
        }
        if ($2012 -gt "0") { 
            foreach ($12ko in $2012ko) { 
                if($12ko.OperatingSystem -notlike "*R2*")
                {
                    $EoS_12 = "10.10.2023 (ES)" ; $EoS_12_fa = "Red"
                    $ena_12 = $12ko.Enabled
                    if($ena_12 -eq $true) {$ena_12_fa = "Red"} else {$ena_12_fa = "Green"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $12ko.OperatingSystem $12ko.Name $12ko.IPv4Address $EoS_12 $ena_12_fa "" $EoS_12_fa
                }
            }
            tablinie
            foreach ($12ko in $2012ko) { 
                if($12ko.OperatingSystem -like "*R2*")
                {
                    $EoS_12 = "10.10.2023 (ES)" ; $EoS_12_fa = "Red"
                    $ena_12 = $12ko.Enabled
                    if($ena_12 -eq $true) {$ena_12_fa = "Red"} else {$ena_12_fa = "Green"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $12ko.OperatingSystem $12ko.Name $12ko.IPv4Address $EoS_12 $ena_12_fa "" $EoS_12_fa
                }
            }
            tablinie
        }
        if ($2016 -gt "0") { 
            foreach ($16ko in $2016ko) { 
                $EoS_16 = "12.01.2027 (ES)" ; $EoS_16_fa = "Yellow"
                $ena_16 = $16ko.Enabled
                if($ena_16 -eq $true) {$ena_16_fa = "Green"} else {$ena_16_fa = "Yellow"}
                neu_tab_max6w_fb "3" "l" "s" "15" $16ko.OperatingSystem $16ko.Name $16ko.IPv4Address $EoS_16 $ena_16_fa "" $EoS_16_fa
            }
            tablinie
        }
        if($srvchk -eq 2){
            if ($2019 -gt "0") {
                foreach ($19ko in $2019ko) { 
                    $EoS_19 = "09.01.2024" ; $EoS_19_fa = "Green"
                    #$EoS_19 = "09.01.2029" ; $EoS_19_fa = "yellow"
                    $ena_19 = $19ko.Enabled
                    if($ena_19 -eq $true) {$ena_19_fa = "Green"} else {$ena_19_fa = "Yellow"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $19ko.OperatingSystem $19ko.Name $19ko.IPv4Address $EoS_19 $ena_19_fa "" $EoS_19_fa
                }
                tablinie
            }
            if ($2022 -gt "0") { 
                foreach ($22ko in $2022ko) { 
                    $EoS_22 = "13.10.2026" ; $EoS_22_fa = "Green"
                    #$EoS_22 = "14.10.2031" ; $EoS_22_fa = "yellow"
                    $ena_22 = $22ko.Enabled
                    if($ena_22 -eq $true) {$ena_22_fa = "Green"} else {$ena_22_fa = "Yellow"}
                    neu_tab_max6w_fb "3" "l" "s" "15" $22ko.OperatingSystem $22ko.Name $22ko.IPv4Address $EoS_22 $ena_22_fa "" $EoS_22_fa
                }
                tablinie
            }
        }
        Leerzeile
    }
}
####################################################################################################
# Funktionen fuer den Bereich "Nicht Windows Systeme"                                              #
####################################################################################################
function oth_chk {
    $s_base = (Get-ADDomain).DistinguishedName
    $other = @()
    $other = Get-ADComputer -filter * -SearchBase $s_base -Properties * | `
        Where-Object { $_.operatingsystem -notlike 'windows*' } | `
        Select-Object SID, Name, Created, DNSHostName, CanonicalName, Enabled, LastLogonDate, `
            Modified, OperatingSystem, IPv4Address, IPv6Address
    Bereichstitel 'Nicht "Windows" Betriebssystem Konten:'
    Leerzeile
    foreach($ot in $other) {
        $sub_txt = "Systemname: " + $ot.Name + " "
        Subtitel $sub_txt "2" "-"
        $cut = (Get-ADDomain).DNSRoot
        $pf = ($ot.CanonicalName).Replace($cut,"")
        [string]$dns = $ot.DNSHostName
        [string]$sid = $ot.SID
        if($ot.OperatingSystem) { [string]$os = $ot.OperatingSystem ; $os_f = "Green" }
            else { [string]$os = "nicht in AD hinterlegt" ; $os_f = $F_Fehler }
        if($ot.Enabled -eq $true) { $status = "aktiviert" ; $status_f = "yellow" }
            else { $status = "deaktiviert" ; $status_f = "green" }
        $cre_date = $ot.Created.ToString("dd.MM.yyyy HH:mm")
        $cha_date = $ot.Modified.ToString("dd.MM.yyyy HH:mm")
        if($ot.LastLogonDate) { 
            $lld_date = $ot.LastLogonDate.ToString("dd.MM.yyyy HH:mm")
            $lld_date_f = "yellow"
        } else { 
            $lld_date = "nicht in AD hinterlegt"
            $lld_date_f = "Red"
        }
        new_2werte "s" ":" "25" "  - DNSHostName" "" "l" $dns "" "l"
        new_2werte "s" ":" "25" "  - SID" "" "l" $sid "" "l"
        new_2werte "s" ":" "25" "  - Pfad in der AD" "" "l" $pf "" "l"
        new_2werte "s" "" "25" "  -----------------------" "Yellow" "l" `
            "----------------------------------------------" "Yellow" "l"
        new_2werte "s" ":" "25" "  - Betriebsystem" "" "l" $os $os_f "l"
        new_2werte "s" ":" "25" "  - System Status" "" "l" $status $status_f "l"
        new_2werte "s" "" "25" "  -----------------------" "Yellow" "l" `
        "----------------------------------------------" "Yellow" "l"
        new_2werte "s" ":" "25" "  - Account Created on" "" "l" $cre_date "" "l"
        new_2werte "s" ":" "25" "  - Account Modified on" "" "l" $cha_date "" "l"
        new_2werte "s" ":" "25" "  - Last Logon Date" "" "l" $lld_date $lld_date_f "l"
        new_2werte "s" "" "25" "  -----------------------" "Yellow" "l" `
        "----------------------------------------------" "Yellow" "l"
        if($ot.IPv4Address) { 
            new_2werte "s" ":" "25" "  - IPv4 Adresse" "" "l" $ot.IPv4Address "Green" "l" }
        if($ot.IPv6Address) { 
            new_2werte "s" ":" "25" "  - IPv6 Adresse" "" "l" $ot.IPv6Address "Yellow" "l" }
        Leerzeile
    }
}
####################################################################################################
# Funktionen fuer den Bereich "gMSA - Group Managed Service Accounts"                              #
####################################################################################################
function KDSR {
    $r_keys = Get-KdsRootKey
    Bereichstitel "Vorhandene(r) KDSRootKey(s):"
    Leerzeile
    foreach ($key in $r_keys) {
        2werte "KeyId                  :" $key.KeyId "s"
        $da_us = $key.CreationTime
        $da_de = [datetime]::ParseExact($da_us, "MM/dd/yyyy HH:mm:ss", `
            [System.Globalization.CultureInfo]::InvariantCulture).ToString("dd.MM.yyyy HH:mm:ss")
        2werte "CreationTime           :" $da_de "s"
        $pass = Test-KdsRootKey -KeyId $key.KeyId
        if ($pass -eq $true) { $faps = "Green" } else { $faps = $F_Fehler }
        2werte "Verify KDS-Conf. to Key:" $pass "s" $faps
        Leerzeile
    }
}
function MSA {
    $msa_all = Get-adServiceAccount -Filter * -Properties * | `
        Where-Object {$_.ObjectClass -eq "msDS-ManagedServiceAccount"}
    $msa_za = ((Get-adServiceAccount -Filter * -Properties * | `
        Where-Object {$_.ObjectClass -eq "msDS-ManagedServiceAccount"}).Name).count
    if($msa_za -gt 1) { } elseif($msa_za -eq 1) { $msa_za = 1 } else { $msa_za = 0 }
    if ($msa_za -ge 1) {
        Bereichstitel "Managed Service Accounts:" 
        Leerzeile
        2werte "Vorhandene(r) Account(s):" $msa_za "s"
        Leerzeile
        foreach ($msa in $msa_all) {
            If ($msa.PrincipalsAllowedToRetrieveManagedPassword) {
                $rights = (($msa.PrincipalsAllowedToRetrieveManagedPassword).Split(",")).`
                    Replace("CN=","")
                $rights = $rights[0]
            } else { $rights = "n.a." }
            if ($msa.Enabled -eq $true) { $smsa = "aktiv" ; $fmsa = "Yellow" } 
                else { $smsa = "inaktiv" ; $fmsa = $F_Fehler }
            Bereichstitel $msa.Name "s"
            Leerzeile
            $temas = $msa.objectClass.Replace("msDS-","")
            2werte " SamAccountName :" $msa.SamAccountName "s"
            2werte " Account Type   :" $temas "s"
            2werte " Status des MSA :" $smsa "s" $fmsa
            2werte " Erlaubte Nutzer:" $rights "s"
            2werte " Erstellt am    :" $msa.whenCreated "s"
            Leerzeile
        }
    }
}
function gMSA {
    $gmsa_all = Get-adServiceAccount -Filter * -Properties * | `
        Where-Object {$_.ObjectClass -eq "msDS-GroupManagedServiceAccount"}
    $gmsa_za = ((Get-adServiceAccount -Filter * -Properties * | `
        Where-Object {$_.ObjectClass -eq "msDS-GroupManagedServiceAccount"}).Name).Count
    if($gmsa_za -gt 1) { } elseif($gmsa_za -eq 1) { $gmsa_za = 1 } else { $gmsa_za = 0 }
    if ($gmsa_za -ge 1) {
        Bereichstitel "Group Managed Service Accounts:" 
        Leerzeile
        2werte "Vorhandene(r) Account(s):" $gmsa_za "s"
        Leerzeile
        foreach ($gmsa in $gmsa_all) {
            If ($gmsa.PrincipalsAllowedToRetrieveManagedPassword) {
                $rights = (($gmsa.PrincipalsAllowedToRetrieveManagedPassword).Split(",")).`
                    Replace("CN=","")
                $rights = $rights[0]
            } else { $rights = "n.a." }
            if ($gmsa.Enabled -eq $true) { $sgmsa = "aktiv" ; $fgmsa = "Yellow" } 
                else { $sgmsa = "inaktiv" ; $fgmsa = $F_Fehler }
            Bereichstitel $gmsa.Name "s"
            Leerzeile
            $tegmas = $gmsa.objectClass.Replace("msDS-","")
            2werte " SamAccountName :" $gmsa.SamAccountName "s"
            2werte " Account Type   :" $tegmas "s"
            2werte " Status des MSA :" $sgmsa "s" $fgmsa
            2werte " Erlaubte Nutzer:" $rights "s"
            2werte " Erstellt am    :" $gmsa.whenCreated "s"
            Leerzeile
        }
    }
}
####################################################################################################
# Funktionen fuer den Bereich "AD-Gruppen"                                                         #
####################################################################################################
function ad_gruppen {
    $fix = @() ; $dsid = (Get-ADDomain).DomainSID.value
    #######################################################################################
    $fix += "$dsid" + "-498"  #  Schreibgeschützte Domänencontroller der Organisation  -498
    $fix += "$dsid" + "-512"  #  Domänen-Admins                                        -512
    $fix += "$dsid" + "-513"  #  Domänen-Benutzer                                      -513
    $fix += "$dsid" + "-514"  #  Domänen-Gäste                                         -514
    $fix += "$dsid" + "-515"  #  Domänencomputer                                       -515
    $fix += "$dsid" + "-516"  #  Domänencontroller                                     -516
    $fix += "$dsid" + "-517"  #  Zertifikatherausgeber                                 -517
    $fix += "$dsid" + "-518"  #  Schema-Admins                                         -518
    $fix += "$dsid" + "-519"  #  Organisations-Admins                                  -519
    $fix += "$dsid" + "-520"  #  Richtlinien-Ersteller-Besitzer                        -520
    $fix += "$dsid" + "-521"  #  Schreibgeschützte Domänencontroller                   -521
    $fix += "$dsid" + "-522"  #  Klonbare Domänencontroller                            -522
    $fix += "$dsid" + "-525"  #  Protected Users                                       -525
    $fix += "$dsid" + "-526"  #  Schlüsseladministratoren                              -526
    $fix += "$dsid" + "-527"  #  Unternehmenssschlüsseladministratoren                 -527
    $fix += "$dsid" + "-553"  #  RAS- und IAS-Server                                   -553
    $fix += "$dsid" + "-571"  #  Zulässige RODC-Kennwortreplikationsgruppe             -571
    $fix += "$dsid" + "-572"  #  Abgelehnte RODC-Kennwortreplikationsgruppe            -572
    $fix += "$dsid" + "-1101" #  DnsAdmins                                            -1101
    $fix += "$dsid" + "-1102" #  DnsUpdateProxy                                       -1102
    $fix += "$dsid" + "-1103" #  DHCP-Benutzer                                        -1103
    $fix += "$dsid" + "-1104" #  DHCP-Administratoren                                 -1104
    #######################################################################################
    #$gruppen = Get-ADGroup -Filter * | `
    #Where-Object { $_.DistinguishedName -notlike "*,CN=Users*" `
    #-and $_.DistinguishedName -notlike "*,CN=Builtin*" }
    $adgruppen = (Get-ADGroup -Filter * | Select-Object name).count
    $adgruppeno = (Get-ADGroup -Filter *  -Properties Members | `
     Where-Object { -not $_.Members}).count
    $adgruppeno1 = Get-ADGroup -Filter *  -Properties Members | `
     Where-Object { -not $_.Members}
    Bereichstitel "AD-Gruppen:"
    Leerzeile
    Bereichstitel "AD-Gruppen allgemein:" "s"
    Leerzeile
    2werte "Vorhandene AD-Gruppen           :" $adgruppen "s"
    2werte "AD-Gruppen ohne Mitglieder      :" "$adgruppeno (inkl. BuildIn)" "s"
    Leerzeile
    Bereichstitel "Namen der leeren AD-Gruppen ohne BuildIn:" "s"
    Leerzeile
    $add_zahl = 1
    neu_tab_max6w_fb "1" "l" "s" "4" "Gruppe" "Nr."
    tablinie "s"
    foreach($gru in $adgruppeno1) {
        if($fix -notcontains $gru.SID) {
            neu_tab_max6w_fb "1" "l" "s" "4" $gru.Name $add_zahl
            $add_zahl = $add_zahl + 1
        }
    }
    Leerzeile
}
####################################################################################################
# Funktionen fuer den Bereich "GPOs"                                                               #
####################################################################################################
function GPO_all {
    $gpos = (Get-GPO -all | Select-Object name).count
    $InvGPO = @()
    $AllGPO = (Get-GPO -All | Where-Object {$_.DisplayName -notmatch "default*"})
    foreach ($GPO in $AllGPO)
    {
    $GPOName = " " ; $stat = " " ; $alink = " " ; $nlink = " " ; $nop = " " ; $emp = " " ; $zahl = 0
    if($GPO.GpoStatus -eq "AllSettingsDisabled") 
    { $stat = "1" ; $zahl = 1 }
    $GPOName = $GPO.Displayname
    [XML]$GPOReport = Get-GPOReport $GPOName -ReportType XML
    $GPOLinks = $GPOReport.GPO.LinksTo
    $GPOApplyPermission = Get-GPPermission $GPOName -All | `
     Where-Object {$_.Permission -match "GpoApply"}
    if ($GPOLinks)
    {
      $GPOLinkCount = $GPOLinks.Count
      $DisabledGPOLinksCount = ($GPOLinks | Where-Object {$_.enabled -eq "false"}).Count
      if ($GPOLinkCount -eq $DisabledGPOLinksCount)
      { $alink = "2" ; $zahl = 1 }
    }
    if (!$GPOLinks)
    {
      $Sitelinked = Get-ADObject -LDAPFilter '(objectClass=site)' `
      -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" `
      -SearchScope OneLevel -Properties gPLink | Where-Object { $_.gpLink -match $GPO.Id}
      if (!$Sitelinked)
      { $nlink= "3" ; $zahl = 1 }
    }
    if (!$GPOApplyPermission)
    { $nop="4" ; $zahl = 1 }
    if (!$GPOReport.GPO.Computer.ExtensionData -and !$GPOReport.GPO.User.ExtensionData)
    { $emp = "5" ; $zahl = 1 }
    if ($zahl -eq 1)
    { $InvGPO += new-object PSObject `
        -property @{GPOName="$GPOName";Dis="$stat";Links="$alink";NoLink="$nlink";NoP="$nop";Empty="$emp"}}
    }
    $InvGPO = $InvGPO | Sort-Object
    $InvAnz = $InvGPO.Count #+ 1
    if($InvAnz -gt 1 ) { $InvAnz = $InvAnz -1 }
    Bereichstitel "Group Policy Objekte (GPO)"
    Leerzeile
    Bereichstitel "GPOs allgemein:" "s" 
    Leerzeile
    2werte "Anzahl der GPOs                 :" $gpos "s"
    2werte "Anzahl der invaliden GPOs       :" $InvAnz "s"
    Leerzeile
    Bereichstitel "Namen der invaliden GPOs:" "s"
    Leerzeile
    2werte " Legende" "" "s"
    2werte " -------" "" "s"
    2werte " 1 = " "Disabled" "s"
    2werte " 2 = " "All Links disabled" "s"
    2werte " 3 = " "Not Linked" "s"
    2werte " 4 = " "No Permissions" "s"
    2werte " 5 = " "Empty" "s"
    Leerzeile
    neu_tab_max6w_fb "5" "r" "s" "1" "GPO-Name" "1" "2" "3" "4" "5"
    tablinie "s"
    foreach ($igpo in $InvGPO)
    {
      [string]$Name = $igpo.GPOName
      [string]$w1 = $igpo.Dis
      [string]$w2 = $igpo.Links
      [string]$w3 = $igpo.NoLink
      [string]$w4 = $igpo.NoP
      [string]$w5 = $igpo.Empty
      neu_tab_max6w_fb "5" "r" "s" "1" $Name $w1 $w2 $w3 $w4 $w5
    }
    tablinie "s"
}
####################################################################################################
# Funktionen fuer den Bereich "dDP Password Settings"                                              #
####################################################################################################
function ddomainpol {
    $ddp = Get-ADDefaultDomainPasswordPolicy
    $ddpre = (Get-ADDefaultDomainPasswordPolicy).ReversibleEncryptionEnabled
    Bereichstitel "Default Domain Password Policy"
    Leerzeile
    Subtitel "Settings:" "1" "-"
    [string]$ddp_com = $ddp.ComplexityEnabled
    if ($ddp.ComplexityEnabled -ne $true) { $f1 = "Red" } else { $f1 = "Green" }
    new_2werte "s" ":" "29" "- ComplexityEnabled" "" "l" $ddp_com $f1 "l"
    [string]$ddp_max = $ddp.MaxPasswordAge
    if ($ddp.MaxPasswordAge -gt '90.00:00:00') { $f2 = "Red" } else { $f2 = "Green" }
    new_2werte "s" ":" "29" "- MaxPasswordAge" "" "l" $ddp_max $f2 "l"
    [string]$ddp_min = $ddp.MinPasswordAge
    if ($ddp.MinPasswordAge -lt '1.00:00:00') { $f3 = "Red" } else { $f3 = "Green" }
    new_2werte "s" ":" "29" "- MinPasswordAge" "" "l" $ddp_min $f3 "l"
    [string]$ddp_mpl = $ddp.MinPasswordLength
    if ($ddp.MinPasswordLength -lt '8') { $f4 = "Red" } 
    elseif ($ddp.MinPasswordLength -lt '12') { $f4 = "Yellow" } else { $f4 = "Green" }
    new_2werte "s" ":" "29" "- MinPasswordLength" "" "l" $ddp_mpl $f4 "l"
    [string]$ddp_phc = $ddp.PasswordHistoryCount
    if ($ddp.PasswordHistoryCount -lt '12') { $f5 = "Red" } else { $f5 = "Green" }
    new_2werte "s" ":" "29" "- PasswordHistoryCount" "" "l" $ddp_phc $f5 "l"
    [string]$ddp_ld = $ddp.LockoutDuration
    new_2werte "s" ":" "29" "- LockoutDuration" "" "l" $ddp_ld "" "l"
    [string]$ddp_low = $ddp.LockoutObservationWindow
    new_2werte "s" ":" "29" "- LockoutObservationWindow" "" "l" $ddp_low "" "l"
    [string]$ddp_lot = $ddp.LockoutThreshold
    if ($ddp.LockoutThreshold -eq '0') { $f6 = "Red" } 
    elseif ($ddp.LockoutThreshold -gt '5') { $f6 = "Yellow" }
    elseif ($ddp.LockoutThreshold -gt '10') { $f6 = "Red" } else { $f6 = "Green" }
    new_2werte "s" ":" "29" "- LockoutThreshold" "" "l" $ddp_lot $f6 "l"
    [string]$ddp_ree = $ddpre
    if ($ddpre -eq $false) { $f7 = "Green" } else { $f7 = "Red" }
    new_2werte "s" ":" "29" "- ReversibleEncryptionEnabled" "" "l" $ddp_ree $f7 "l"
    Leerzeile
}
####################################################################################################
# Funktionen fuer den Bereich "fGPP - fine Grained Password Policies"                              #
####################################################################################################
function fGPO {
    $fin = @(Get-ADFineGrainedPasswordPolicy -Filter *) | Sort-Object Precedence
    $fine = (Get-ADFineGrainedPasswordPolicy -Filter *).count
    if($fine -eq 0) { $fine1 = "0" } elseif($fine -ge 1) { $fine1 = $fine } else { $fine1 = "1" }
    Bereichstitel "Fine Grained Password Policy(s)" ""
    Leerzeile
    2werte "- Anzahl der vorhandenen fGPOs:" $fine1 "s"
    Leerzeile
    Foreach($fines in $fin)
    {
        $fgpo_name = "fGPO: " + '"' + $fines.Name + '"' + " Settings:"
        Subtitel $fgpo_name "1" "-"
        [string]$fi_name = $fines.Name
        new_2werte "s" ":" "29" "- Name" "" "l" $fi_name "" "l"
        [string]$fi_pre = $fines.Precedence
        new_2werte "s" ":" "29" "- Precedence" "" "l" $fi_pre "" "l"
        $appl = $fines.AppliesTo
        $appl_count = $appl.Count
        for ($i=0;$i -lt $appl_count;$i++) {
            $apx = Get-ADObject -Identity $appl[$i] | Select-Object Name,ObjectClass
            $apx_txt = $apx.Name
            if($apx.ObjectClass -eq "group") { $apx_fa = "Green" }
                else { $apx_fa = "Red" ; $apx_txt = $apx_txt + " (Benutzer)" }
            if($i -eq 0) { new_2werte "s" ":" "29" "- AppliesTo" "" "l" $apx_txt $apx_fa "l" }
                else { new_2werte "s" ":" "29" "           " "" "l" $apx_txt $apx_fa "l" }
        }
        [string]$fi_ce = $fines.ComplexityEnabled
        if ($fines.ComplexityEnabled -ne $true) { $f1 = "Red" } else { $f1 = "Green" }
        new_2werte "s" ":" "29" "- ComplexityEnabled" "" "l" $fi_ce "$f1" "l"
        [string]$fi_max = $fines.MaxPasswordAge
        if ($fines.MaxPasswordAge -gt '90.00:00:00') { $f2 = "Red" } else { $f2 = "Green" }
        new_2werte "s" ":" "29" "- MaxPasswordAge" "" "l" $fi_max "$f2" "l"
        [string]$fi_min = $fines.MinPasswordAge
        if ($fines.MinPasswordAge -lt '1.00:00:00') { $f3 = "Red" } else { $f3 = "Green" }
        new_2werte "s" ":" "29" "- MinPasswordAge" "" "l" $fi_min "$f3" "l"
        [string]$fi_mpl = $fines.MinPasswordLength
        if ($fines.MinPasswordLength -lt '8') { $f4 = "Red" } 
        elseif ($fines.MinPasswordLength -lt '12') { $f4 = "Yellow" } else { $f4 = "Green" }
        new_2werte "s" ":" "29" "- MinPasswordLength" "" "l" $fi_mpl "$f4" "l"
        [string]$fi_phc = $fines.PasswordHistoryCount
        if ($fines.PasswordHistoryCount -lt '12') { $f5 = "Red" } else { $f5 = "Green" }
        new_2werte "s" ":" "29" "- PasswordHistoryCount" "" "l" $fi_phc "$f5" "l"
        [string]$fi_ld = $fines.LockoutDuration
        new_2werte "s" ":" "29" "- LockoutDuration" "" "l" $fi_ld "" "l"
        [string]$fi_low = $fines.LockoutObservationWindow
        new_2werte "s" ":" "29" "- LockoutObservationWindow" "" "l" $fi_low "" "l"
        [string]$fi_lot = $fines.LockoutThreshold
        if ($fines.LockoutThreshold -eq '0') { $f6 = "Red" } 
        elseif ($fines.LockoutThreshold -gt '5') { $f6 = "Yellow" }
        elseif ($fines.LockoutThreshold -gt '10') { $f6 = "Red" } else { $f6 = "Green" }
        new_2werte "s" ":" "29" "- LockoutThreshold" "" "l" $fi_lot "$f6" "l"
        [string]$fi_ree = $fines.ReversibleEncryptionEnabled
        if ($fines.ReversibleEncryptionEnabled -eq $false) { $f7 = "Green" } else { $f7 = "Red" }
        new_2werte "s" ":" "29" "- ReversibleEncryptionEnabled" "" "l" $fi_ree "$f7" "l"
        Leerzeile
    }  
}
####################################################################################################
# Funktionen fuer den Bereich "User vs Password Policies"                                          #
####################################################################################################
function pw_user ($Wrt1,$user) {
    Subtitel $Wrt1 "1" "-"
    Leerzeile
    neu_tab_max6w_fb "3" "l" "s" "14" "User Principal Name" "SamAccountName" "PW Last Set" "Status"
    tablinie "s"
    foreach ($use in $user) {
        $use_san = $use.SamAccountName
        $use_upn = $use.UserPrincipalName
        $use_ena = $use.Enabled
        if($use_ena -eq $true) {
            $use_ena_txt = "aktiviert"
            $use_ena_fa = "Red"
        } else {
            $use_ena_txt = "deaktiviert"
            $use_ena_fa = "Green"
        }
        if((Get-ADUser -Identity $use_san -Properties *).PasswordLastSet) {
            [string]$use_pls = (Get-ADUser -Identity $use_san -Properties *).PasswordLastSet
            $use_pls = $use_pls.Split(" ")[0]
            $use_pls_temp = [DateTime]::ParseExact($use_pls,"MM/dd/yyyy", $null)
            $use_pls = $use_pls_temp.ToString("dd.MM.yyyy")
            $use_pls_fa = "Yellow"
        } else { 
            $use_pls = "kein Kennwort"
            $use_pls_fa = "Red"
        }
        neu_tab_max6w_fb "3" "l" "s" "14" $use_upn $use_san $use_pls $use_ena_txt "" $use_pls_fa $use_ena_fa        
    }
    Leerzeile
}
function spezial_user {
    Bereichstitel 'Benutzer, die die Password Policies "umgehen"!'
    Leerzeile
    $uohne = Get-ADUser -Filter * -Property * | Where-Object { $_.userAccountControl -eq "544" }
    if($uohne) { 
        pw_user "Benutzer, die kein Kennwort hinterlegen müssen (544):" $uohne
    }
    $ulauf = Get-ADUser -Filter * -Property * | Where-Object { $_.userAccountControl -eq "66048" }
    if($ulauf) { 
        pw_user "Benutzer, deren Kennwort nie abläuft (66048):" $ulauf
    }
    $ohneuablauf = Get-ADUser -Filter * -Property * | Where-Object { $_.userAccountControl -eq "66080" }
    if($ohneuablauf) { 
        pw_user "Benutzer, die kein Kennwort brauchen, das auch nie abläuft (66080):" $ohneuablauf
    }
    $uohnedeak = Get-ADUser -Filter * -Property * | Where-Object { $_.userAccountControl -eq "546" }
    if ($uohnedeak) { 
        pw_user "Deak. Benutzer, die kein Kennwort hinterlegt haben müssen (546):" $uohnedeak
    }
    $de_op_nie = Get-ADUser -Filter * -Property * | Where-Object { $_.userAccountControl -eq "66082" }
    if ($de_op_nie) {
        pw_user "Deak. Benutzer, die kein Kennwort brauchen, das auch nie abläuft (66082):" $de_op_nie
    }
}
####################################################################################################
# Funktionen fuer den Bereich "Organisation Units"                                                 #
####################################################################################################
function OUS {
    ### Legende ####################################################################################
    # neu_tab_max6w_fb ([int]$spa,$pos,$sub,[int]$bre,$txt,$we1,$we2,$we3,$we4,$we5,$we6)          #
    # $spa = Anzahl der Tabellenspalten                                                            #
    # $pos = Tabelle links(l) oder rechts(r)                                                       #
    # $sub = Einsatzbereich als Subtabelle? nein(n), ja(s)                                         #
    # $bre = Breite der Tabellenspalten                                                            #
    # $txt = Text der Spalte vor bzw. hinter der Tabelle                                           #
    # $we(n) = Wert der n. Spalte                                                                  #
    ### Head #######################################################################################
    Bereichstitel "Aufteilung OU zu User,Systemen, AD-Gruppen und gMSA" "s"
    Leerzeile
    if ($OrgUni -eq '2') { $OU_Head = ("Organisation Unit").PadRight(18) + "- Pfad" } 
        else { $OU_Head = "Organisation Unit" }
    neu_tab_max6w_fb "5" "l" "s" "4" "$OU_Head" "User" "Sys" "ADGr" "gMSA" "OU's"
    tablinie "s"
    ### gen. Var. Global ###########################################################################
    $dom = Get-ADDomain ; $dist = $dom.DistinguishedName ; $dom_root = $dom.DNSRoot
    $pc = "CN=Computers,"+"$dist" ; $us = "CN=Users,"+"$dist"
    $bu = "CN=Builtin,"+"$dist" ; $ms = "CN=Managed Service Accounts,"+"$dist"
    ### Builtin Computer ###########################################################################
    $pc_path = (Get-ADObject -Identity $pc -Properties * | Select-Object CanonicalName).CanonicalName
    $pc_path = $pc_path.Replace("$dom_root/","~/").Replace("Computers","")
    if ($OrgUni -eq '2') { $N1 = ("Computers").PadRight(25) + "- $pc_path" } else { $N1 = "Computers" }
    [string]$u1 = (Get-ADUser -Filter * -SearchBase $pc -SearchScope OneLevel).count
    if($u1 -eq 0) { $u11 = "0" } elseif($u1 -ge 1) {$u11 = $u1} else { $u11 = "1" }
    [string]$s1 = (Get-AdComputer -Filter * -SearchBase $pc -SearchScope OneLevel).Count
    if($s1 -eq 0) { $s11 = "0" } elseif($s1 -ge 1) {$s11 = $s1} else { $s11 = "1" }
    [string]$AD1 = (Get-ADGroup -Filter * -SearchBase $pc -Searchscope OneLevel).count
    if($AD1 -eq 0) { $AD11 = "0" } elseif($AD1 -ge 1) {$AD11 = $AD1} else { $AD11 = "1" }
    [string]$gMS1 = (Get-ADServiceAccount -Filter * -SearchBase $pc -Searchscope OneLevel).count
    if($gMS1 -eq 0) { $gMS11 = "0" } elseif($gMS1 -ge 1) {$gMS11 = $gMS1} else { $gMS11 = "1" }
    [string]$o_u1 = (Get-ADOrganizationalUnit -Filter * -SearchBase $pc).count
    if($o_u1 -eq 0) { $o_u11 = "0" } elseif($o_u1 -ge 1) {$o_u11 = $o_u1} else { $o_u11 = "1" }
    if ($o_u1) { 
        if($o_u1 -eq 0) { $o_u11 = "0" }
        elseif($o_u1 -ge 1) {$o_u11 = $o_u1 - 1}
        else { $o_u11 = "1" }
    } else {
        $o_u11 = "0"
    }
    neu_tab_max6w_fb "5" "l" "s" "4" $N1 $u11 $s11 $AD11 $gMS11 $o_u11
    ### Builtin Users ##############################################################################
    $us_path = (Get-ADObject -Identity $us -Properties * | Select-Object CanonicalName).CanonicalName
    $us_path = $us_path.Replace("$dom_root/","~/").Replace("Users","")
    if ($OrgUni -eq '2') { $N2 = ("Users").PadRight(25) + "- $us_path" } else { $N2 = "Users" }
    [string]$u2 = (Get-ADUser -Filter * -SearchBase $us -SearchScope OneLevel).count
    if($u2 -eq 0) { $u21 = "0" } elseif($u2 -ge 1) {$u21 = $u2} else { $u21 = "1" }
    [string]$s2 = (Get-AdComputer -Filter * -SearchBase $us -SearchScope OneLevel).Count
    if($s2 -eq 0) { $s21 = "0" } elseif($s2 -ge 1) {$s21 = $s2} else { $s21 = "1" }
    [string]$AD2 = (Get-ADGroup -Filter * -SearchBase $us -Searchscope OneLevel).count
    if($AD2 -eq 0) { $AD21 = "0" } elseif($AD2 -ge 1) {$AD21 = $AD2} else { $AD21 = "1" }
    [string]$gMS2 = (Get-ADServiceAccount -Filter * -SearchBase $us -Searchscope OneLevel).count
    if($gMS2 -eq 0) { $gMS21 = "0" } elseif($gMS2 -ge 1) {$gMS21 = $gMS2} else { $gMS21 = "1" }
    [string]$o_u2 = (Get-ADOrganizationalUnit -Filter * -SearchBase $us).count
    if($o_u2 -eq 0) { $o_u21 = "0" } elseif($o_u2 -ge 1) {$o_u21 = $o_u2} else { $o_u21 = "1" }
    if ($o_u2) { 
        if($o_u2 -eq 0) { $o_u21 = "0" }
        elseif($o_u2 -ge 1) {$o_u21 = $o_u2 - 1}
        else { $o_u21 = "1" }
    } else {
        $o_u21 = "0"
    }
    neu_tab_max6w_fb "5" "l" "s" "4" $N2 $u21 $s21 $AD21 $gMS21 $o_u21
    ### Builtin Builtin ############################################################################
    $bu_path = (Get-ADObject -Identity $bu -Properties * | Select-Object CanonicalName).CanonicalName
    $bu_path = $bu_path.Replace("$dom_root/","~/").Replace("Builtin","")
    if ($OrgUni -eq '2') { $N3 = ("Builtin").PadRight(25) + "- $bu_path" } else { $N3 = "Builtin" }
    [string]    $u3 = (Get-ADUser -Filter * -SearchBase $bu -SearchScope OneLevel).count
    if($u3 -eq 0) { $u31 = "0" } elseif($u3 -ge 1) {$u31 = $u3} else { $u31 = "1" }
    [string]$s3 = (Get-AdComputer -Filter * -SearchBase $bu -SearchScope OneLevel).Count
    if($s3 -eq 0) { $s31 = "0" } elseif($s3 -ge 1) {$s31 = $s3} else { $s31 = "1" }
    [string]$AD3 = (Get-ADGroup -Filter * -SearchBase $bu -Searchscope OneLevel).count
    if($AD3 -eq 0) { $AD31 = "0" } elseif($AD3 -ge 1) {$AD31 = $AD3} else { $AD31 = "1" }
    [string]$gMS3 = (Get-ADServiceAccount -Filter * -SearchBase $bu -Searchscope OneLevel).count
    if($gMS3 -eq 0) { $gMS31 = "0" } elseif($gMS3 -ge 1) {$gMS31 = $gMS3} else { $gMS31 = "1" }
    [string]$o_u3 = (Get-ADOrganizationalUnit -Filter * -SearchBase $bu).count
    if($o_u3 -eq 0) { $o_u31 = "0" } elseif($o_u3 -ge 1) {$o_u31 = $o_u3} else { $o_u31 = "1" }
    if ($o_u3) { 
        if($o_u3 -eq 0) { $o_u31 = "0" }
        elseif($o_u3 -ge 1) {$o_u31 = $o_u3 - 1}
        else { $o_u31 = "1" }
    } else {
        $o_u31 = "0"
    }
    neu_tab_max6w_fb "5" "l" "s" "4" $N3 $u31 $s31 $AD31 $gMS31 $o_u31
    ### Builtin Managed Service Accounts ###########################################################
    $ms_path = (Get-ADObject -Identity $ms -Properties * | Select-Object CanonicalName).CanonicalName
    $ms_path = $ms_path.Replace("$dom_root/","~/").Replace("Managed Service Accounts","")
    if ($OrgUni -eq '2') { $N4 = ("Managed Service Accounts").PadRight(25) + "- $ms_path" } 
        else { $N4 = "Managed Service Accounts" }
    [string]$u4 = (Get-ADUser -Filter * -SearchBase $ms -SearchScope OneLevel).count
    if($u4 -eq 0) { $u41 = "0" } elseif($u4 -ge 1) {$u41 = $u4} else { $u41 = "1" }
    [string]$s4 = (Get-AdComputer -Filter * -SearchBase $ms -SearchScope OneLevel).Count
    if($s4 -eq 0) { $s41 = "0" } elseif($s4 -ge 1) {$s41 = $s4} else { $s41 = "1" }
    [string]$AD4 = (Get-ADGroup -Filter * -SearchBase $ms -Searchscope OneLevel).count
    if($AD4 -eq 0) { $AD41 = "0" } elseif($AD4 -ge 1) {$AD41 = $AD4} else { $AD41 = "1" }
    [string]$gMS4 = (Get-ADServiceAccount -Filter * -SearchBase $ms -Searchscope OneLevel).count
    if($gMS4 -eq 0) { $gMS41 = "0" } elseif($gMS4 -ge 1) {$gMS41 = $gMS4} else { $gMS41 = "1" }
    [string]$o_u4 = (Get-ADOrganizationalUnit -Filter * -SearchBase $ms).count
    if($o_u4 -eq 0) { $o_u41 = "0" } elseif($o_u4 -ge 1) {$o_u41 = $o_u4} else { $o_u41 = "1" }
    if ($o_u4) { 
        if($o_u4 -eq 0) { $o_u41 = "0" }
        elseif($o_u4 -ge 1) {$o_u41 = $o_u4 - 1}
        else { $o_u41 = "1" }
    } else {
        $o_u41 = "0"
    }
    neu_tab_max6w_fb "5" "l" "s" "4" $N4 $u41 $s41 $AD41 $gMS41 $o_u41
    tablinie "s"
    ### other Organisation Units ###################################################################
    $OUs = Get-ADOrganizationalUnit -Properties CanonicalName -Filter * | Sort-Object CanonicalName
    $OUlength = (Get-ADOrganizationalUnit -Properties Name -Filter * | Sort-Object Name).Name
    $len = 0 ; foreach ($oul in $OUlength) { if ($oul.Length -gt $len) { $len = $oul.Length } }
    foreach ($OU in $OUs) {
        $Name = Split-Path $OU.CanonicalName -Leaf
        $ou_path = (Get-ADObject -Identity $ou -Properties * | Select-Object CanonicalName).CanonicalName
        $ou_path = $ou_path.Replace("$dom_root/","~/").Replace("$Name","")
        if ($OrgUni -eq '2') { $ou_name = ("$Name").PadRight($len) + " - $ou_path" } 
            else { $ou_name = "$Name" }
        $User = (Get-AdUser -Filter * -SearchBase $OU.DistinguishedName -SearchScope OneLevel).Count
        if($User -eq 0) { $User1 = "0" } elseif($User -ge 1) {$User1 = $User} else { $User1 = "1" }
        $Sys = (Get-AdComputer -Filter * -SearchBase $OU.DistinguishedName -SearchScope OneLevel).Count
        if($Sys -eq 0) { $Sys1 = "0" } elseif($Sys -ge 1) {$Sys1 = $Sys} else { $Sys1 = "1" }
        $ADG = (Get-ADGroup -Filter * -SearchBase $OU.DistinguishedName -Searchscope OneLevel).count
        if($ADG -eq 0) { $ADG1 = "0" } elseif($ADG -ge 1) {$ADG1 = $ADG} else { $ADG1 = "1" }
        $gMSA = (Get-ADServiceAccount -Filter * -SearchBase $OU.DistinguishedName -Searchscope OneLevel).count
        if($gMSA -eq 0) { $gMSA1 = "0" } elseif($gMSA -ge 1) {$gMSA1 = $gMSA} else { $gMSA1 = "1" }
        [string]$o_u5 = (Get-ADOrganizationalUnit -Filter * -SearchBase $ou | Select-Object Name).count
        if ($o_u5) { 
            if($o_u5 -eq 0) { $o_u51 = "0" }
            elseif($o_u5 -ge 1) {$o_u51 = $o_u5 - 1}
            else { $o_u51 = "1" }
        } else {
            $o_u51 = "0"
        }
        neu_tab_max6w_fb "5" "l" "s" "4" $ou_name $User1 $Sys1 $ADG1 $gMSA1 $o_u51
    }
    tablinie "s"
    Leerzeile
}
####################################################################################################
# Funktionen fuer den Bereich "DACL, Rechte-Delegierung"                                           #
####################################################################################################
function dacls {
    $doma = Get-ADDomain
    $drin = $doma.NetBIOSName
    $pfad = "AD:\" + $doma.DistinguishedName
    $pfad_orig = Get-Location
    Set-Location $pfad
    $O_Units = Get-ChildItem -Recurse | Where-Object ObjectClass -eq organizationalUnit
    foreach ($O_U in $O_Units)
	{
		$ARs = (Get-ACL -Path $O_U.PSpath).Access
        foreach ($AR in $ARs) {
            if($AR.ActiveDirectoryRights -notlike "gen*") {
                if($AR.IdentityReference -match $drin) {
                    if ($AR.IdentityReference -notmatch "Unter" -and
                        $AR.IdentityReference -notmatch "Schl" -and
                        $AR.IdentityReference -notmatch "Dom" ) 
                    {
                        if($AR.IdentityReference.Value) {
                            $name = $AR.IdentityReference.Value
                            [string]$rech = $AR.ActiveDirectoryRights
                            [string]$vere = $AR.IsInherited
                            if($vere -eq $true) { $vtxt = "ja" } else { $vtxt = "nein" }
                            $name_txt = "Organisation Unit:  " + $O_U.Name + " "
                            Subtitel $name_txt "1" "-"
                            new_2werte "s" ":" "15" "OU Pfad" "" "l" $O_U.distinguishedName "" "l"
                            new_2werte "s" ":" "15" "Rechteinhaber" "" "l" $name "" "l"
                            new_2werte "s" ":" "15" "Berechtigungen" "" "l" $rech "" "l"
                            new_2werte "s" ":" "15" "Rechte geerbt?" "" "l" $vtxt "" "l"
                            tablinie "s"
                        }
                    }
                }
            }
        }		
	}
    Set-Location $pfad_orig
}
####################################################################################################
# Funktionen fuer den Bereich "Zertifizierungsstelle(n)"                                           #
####################################################################################################
function ca_root {
    Bereichstitel "Root CA's:"
    Leerzeile
    $d_dis = (Get-ADDomain).DistinguishedName
    $cp1 = "CN=Enrollment Services,CN=Public Key Services,CN=Services,CN=Configuration," + $d_dis
    $cp2 = "CN=Certification Authorities,CN=Public Key Services,CN=Services,CN=Configuration," + $d_dis
    $subs = Get-ADObject -Filter { objectClass -eq 'pKIEnrollmentService'} `
        -SearchBase $cp1 | Select-Object Name, distinguishedName, objectClass, DNSHostName
    $roots = Get-ADObject -Filter { objectClass -eq 'certificationAuthority'} `
        -SearchBase $cp2 -Properties * | Select-Object Name, distinguishedName, objectClass | `
        Sort-Object Name  
    foreach($ro in $roots){
        if($subs.Name -contains $ro.Name) { $fa_su = "Red" } else { $fa_su = "Green" }
        2werte "Zertifizierungsstellenname:" $ro.Name "s" $fa_su
    }
    Leerzeile
}
function ca_sub {
    Bereichstitel "Sub CA's:"
    Leerzeile
    $d_dis = (Get-ADDomain).DistinguishedName
    $cp1 = "CN=Enrollment Services,CN=Public Key Services,CN=Services,CN=Configuration," + $d_dis
    $cp2 = "CN=Certification Authorities,CN=Public Key Services,CN=Services,CN=Configuration," + $d_dis
    $subs = Get-ADObject -Filter { objectClass -eq 'pKIEnrollmentService'} `
        -SearchBase $cp1 | Select-Object Name, distinguishedName, objectClass, DNSHostName | `
        Sort-Object name
    $roots = Get-ADObject -Filter { objectClass -eq 'certificationAuthority'} `
        -SearchBase $cp2 -Properties * | Select-Object Name, distinguishedName, objectClass   
    foreach($su in $subs){
        if($roots.Name -contains $su.Name) { $fa_su = "Red" } else { $fa_su = "Green" }
        $ca_name = Get-ADObject -Identity $su.DistinguishedName -Properties * | `
            Select-Object DNSHostName
        $puf = ($ca_name.DNSHostName).Split(".")
        $su_ca = Get-ADComputer -Identity $puf[0] -Properties *
        $netco = Test-Connection $puf[0] -Count 1 -ErrorAction SilentlyContinue
        if($netco) { $err = "aktiv" ; $fa_err = "Green" } else { $err = "inaktiv" ; $fa_err = "Red" }



        $au_ti = "Zertifizierungsstellenname: " + $su.Name
        $au_ti_ul = "-" * $au_ti.Length
        2werte "Zertifizierungsstellenname:" $su.Name "s" $fa_su
        2werte $au_ti_ul "" "s"
        2werte " - Hostname        :" $su_ca.Name "s"
        2werte " - Betriebssystem  :" $su_ca.OperatingSystem "s"
        2werte " - IPv4-Adresse    :" $su_ca.IPv4Address "s"
        2werte " - AD-Lokalisierung:" $su_ca.CanonicalName "s"
        2werte " - Pingbar via Nic :" $err "s" $fa_err
        Leerzeile
    }
    Leerzeile
}
function ca_templates {
    Bereichstitel "Vorhandene Zertifikatsvorlagen im AD:"
    Leerzeile
    neu_tab_max6w_fb "1" "l" "s" "19" "Name der Vorlage" "Modifiziert am"
    tablinie
    $d_dis = (Get-ADDomain).DistinguishedName
    $cp = "CN=Services,CN=Configuration," + $d_dis
    $templ = Get-ADObject -Filter { objectClass -eq 'pKICertificateTemplate'} `
        -SearchBase $cp -SearchScope Subtree -Properties * | `
        Select-Object Name, Modified | Sort-Object Modified
    foreach($te in $templ) {
        neu_tab_max6w_fb "1" "l" "s" "19" $te.Name $te.Modified
        #2werte "$tee :" $te.Modified "s"
    }
    Leerzeile
}
####################################################################################################
# Globale Variablen fuer den Bereich "Domain Controller"                                           # 
####################################################################################################
$dcons = ((Get-ADForest).Domains | `
    ForEach-Object {Get-ADDomainController -Filter * -Server $_}).Hostname
####################################################################################################
# Funktionen fuer den Bereich "Domain Controller"                                                  #
####################################################################################################
function dcdienste ($dcho) {
    $dienste = @()
    ################################################################################################
    $dienste += "AxInstSV"
    $dienste += "ALG"
    $dienste += "bthserv"
    $dienste += "cdpsvc"
    $dienste += "MapsBroker"
    $dienste += "lfsvc"
    $dienste += "SharedAccess"
    $dienste += "lltdsvc"
    $dienste += "wlidsvc"
    $dienste += "AppVClient"
    $dienste += "NgcSvc"
    $dienste += "NgcCtnrSvc"
    $dienste += "NetTcpPortSharing"
    $dienste += "NcbService"
    $dienste += "CscService"
    $dienste += "PhoneSvc"
    $dienste += "PrintNotify"
    $dienste += "PcaSvc"
    $dienste += "QWAVE"
    $dienste += "RmSvc"
    $dienste += "RemoteAccess"
    $dienste += "SensorDataService"
    $dienste += "SensrSvc"
    $dienste += "SensorService"
    $dienste += "SCardSvr"
    $dienste += "ScDeviceEnum"
    $dienste += "SSDPSRV"
    $dienste += "WiaRpc"
    $dienste += "TabletInputService"
    $dienste += "upnphost"
    $dienste += "UevAgentService"
    $dienste += "WalletService"
    $dienste += "Audiosrv"
    $dienste += "AudioEndpointBuilder"
    $dienste += "FrameServer"
    $dienste += "stisvc"
    $dienste += "wisvc"
    $dienste += "icssvc"
    ################################################################################################
    # Datensammlung remote auf dem DC (read-only); Ausgabe erfolgt anschliessend lokal.
    # (frueher: Enter-/Exit-PSSession - im Skript-Kontext wirkungslos, lief lokal statt remote)
    $remote = Invoke-Command -ComputerName $dcho -ScriptBlock {
        $alle = Get-Service -Name *
        $wmi  = Get-WmiObject win32_service | Where-Object {
                    ($_.startname -ne "LocalSystem") -and
                    ($_.startname -ne "NT AUTHORITY\NetworkService") -and
                    ($_.startname -ne "NT AUTHORITY\NETWORK SERVICE") -and
                    ($_.startname -ne "NT AUTHORITY\LocalService") }
        # Status/Starttyp als String, damit die Werte die Deserialisierung sauber ueberstehen.
        [pscustomobject]@{
            Dienst = @($alle | Select-Object Name,
                        @{ n = 'StartType'; e = { [string]$_.StartType } },
                        @{ n = 'Status';    e = { [string]$_.Status } })
            AnzGes = @($alle).Count
            AnzDis = @($alle | Where-Object { $_.StartType -eq 'Disabled'  }).Count
            AnzMan = @($alle | Where-Object { $_.StartType -eq 'Manual'    }).Count
            AnzAut = @($alle | Where-Object { $_.StartType -eq 'Automatic' }).Count
            AnzRun = @($alle | Where-Object { $_.Status    -eq 'Running'   }).Count
            AnzSto = @($alle | Where-Object { $_.Status    -eq 'Stopped'   }).Count
            LokCon = @($wmi | Select-Object name, startmode)
            LokAnz = @($wmi).Count
        }
    }
    $dienst = $remote.Dienst
    $anzges = $remote.AnzGes
    $anzdis = $remote.AnzDis
    $anzman = $remote.AnzMan
    $anzaut = $remote.AnzAut
    $anzrun = $remote.AnzRun
    $anzsto = $remote.AnzSto
    $lokcon = $remote.LokCon
    $Lokanz = $remote.LokAnz
    Bereichstitel "Zu den Diensten:" "s"
    Leerzeile
    2werte "Anzahl der Dienste Gesamt:" $anzges "s"
    Leerzeile
    Bereichstitel "Dienste nach Starttype:" "s"
    Leerzeile
    2werte "Automatic:" $anzaut "s"
    2werte "Manual   :" $anzman "s"
    2werte "Disabled :" $anzdis "s"
    Leerzeile
    Bereichstitel "Dienste nach Status:" "s"
    Leerzeile
    2werte "Running:" $anzrun "s"
    2werte "Stopped:" $anzsto "s"
    Leerzeile
    $dienst = $dienst | Sort-Object Status
    $dienrun = $dienst | Where-Object {($_.Status -eq "Running")}
    Bereichstitel "Dienste mit Status Running:" "s"
    Leerzeile
    neu_tab_max6w_fb "2" "r" "s" "9" "Dienstname" "Starttype" "Status"
    tablinie "s"
    foreach ($die in $dienrun) {
      [string]$name = $die.name
      $name = $name.Trim()
      [string]$styp = $die.starttype
      $styp = $styp.Trim()
      [string]$stat = $die.Status
      $stat = $stat.Trim()
      if($dienste -contains $name) { $farbe = "Red" } else { $farbe = "White" }
      neu_tab_max6w_fb "2" "r" "s" "9" $name $styp $stat "$farbe" "$farbe"
    }
    Leerzeile
    $lokcon = $lokcon | Sort-Object startmode
    Bereichstitel "Dienste die nicht im Maschinen Kontext gestartet werden:" "s"
    Leerzeile
    2werte "Anzahl der Dienste:" $Lokanz "s"
    Leerzeile
    neu_tab_max6w_fb "1" "l" "s" "9" "Dienstname" "Starttype"
    tablinie "s"
    foreach ($lcon in $lokcon) {
      $na = $lcon.name
      $sty = $lcon.startmode
      neu_tab_max6w_fb "1" "l" "s" "9" $na $sty
    }
    Leerzeile
}
function dcprog ($dcpro) {
$prog = Invoke-Command -ComputerName $dcpro -ScriptBlock `
    { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | `
    Sort-Object -Property PSChildName | Select-Object DisplayName, PSChildName, Publisher, DisplayVersion, InstallDate }
Bereichstitel "Installierte Programme:" "s"
Leerzeile
foreach ($pro in $prog) {
    $nam = $pro.DisplayName
    if ( $null -eq $nam) { $nam = $pro.PSChildName } ; 2werte "Programm      :" "$nam" "s"
    $pub = $pro.Publisher
    if ( $null -eq $pub) { $pub = "n/a" }            ; 2werte "Publisher     :" "$pub" "s"
    $ver = $pro.DisplayVersion
    if ( $null -eq $ver) { $ver = "n/a" }            ; 2werte "Version       :" "$ver" "s"
    $ins = $pro.InstallDate
    if ( $null -eq $ins) { $ins = "n/a" }            ; 2werte "Installiert am:" "$ins" "s"
    Leerzeile
}
}
function dchot ($dcakt) {
    $fixe =  Invoke-Command -ComputerName $dcakt -ScriptBlock `
     { Get-HotFix | Sort-Object HotFixID -Descending | `
      Select-Object HotFixID, Description, InstalledBy, InstalledOn }
    Bereichstitel "Installierte HotFixe:" "s"
    Leerzeile
    neu_tab_max6w_fb "3" "l" "s" "13" "HitFixID" "Beschreibung" "Inst. von" "Datum"
    tablinie "s"
    foreach ($fix in $fixe) {
        [string]$hf = $fix.HotFixID
        [string]$di = $fix.Description
        $di = $di.Split(" ")[0]
        [string]$ib = $fix.InstalledBy
        if([string]::IsNullOrEmpty($ib)) { $ib = "n.a." } 
        else { $ib = $ib.Split("\")[1] }
        [string]$io = $fix.InstalledOn
        $io = $io.Split(" ")[0]
        $io_temp = [DateTime]::ParseExact($io,"MM/dd/yyyy", $null)
        $io = $io_temp.ToString("dd.MM.yyyy")
        neu_tab_max6w_fb "3" "l" "s" "13" $hf $di $ib $io
    }
    Leerzeile
}
function dcroles ($dcakt) {
    $rollen = Invoke-Command -ComputerName $dcakt -ScriptBlock `
     { Get-WindowsFeature | Where-Object { $_.installstate -eq "installed" } }
    Bereichstitel "Installierte Windows Rollen:" "s"
    Leerzeile
    neu_tab_max6w_fb "1" "l" "s" "9" "Displayname - Name" "Status"
    tablinie "s"
    foreach ($rol in $rollen) {
      $dnam = $rol.DisplayName
      $name = $rol.Name
      $toge = "$dnam - $name"
      $inst = $rol.Installstate
      neu_tab_max6w_fb "1" "l" "s" "9" $toge $inst
    }
    Leerzeile
}
function dcfeature ($dcakt) {
    $features = Invoke-Command -ComputerName $dcakt -ScriptBlock `
    { Get-WindowsOptionalFeature -Online | Where-Object { $_.state -eq "Enabled" } } 
    Bereichstitel "Installierte Windows Optional Feature:" "s"
    Leerzeile
    neu_tab_max6w_fb "1" "l" "s" "9" "Featurename" "Status"
    tablinie
    foreach ($feat in $features) {
      [string]$name = $feat.Featurename
      if([string]::IsNullOrEmpty($name)) { $name = "n.a." }
      [string]$inst = $feat.state
      neu_tab_max6w_fb "1" "l" "s" "9" $name $inst
    }
    Leerzeile
}
function dc_ldaps ($dc){
    $TcpClient1 = $null
    $TcpClient2 = $null
    $HostName = $dc
    $Port1 = 636  # Default LDAPS port
    $Port2 = 3269 # Sysvol LDAPS port
    $TcpClient1 = New-Object System.Net.Sockets.TcpClient($HostName,$Port1)
    $TcpClient2 = New-Object System.Net.Sockets.TcpClient($HostName,$Port2)
    if ($null -eq $TcpClient1) { $wert1 = "n.a." ; $fa1 = "Red" } else { $wert1 = "enabled" ; $fa1 = "Green" }
    if ($null -eq $TcpClient2) { $Wert2 = "n.a." ; $fa2 = "Red" } else { $Wert2 = "enabled" ; $fa2 = "Green" }
    Bereichstitel "Voraussetzungen für LDAPS:" "s"
    Leerzeile
    2werte " Zugriff über Port  636 möglich:" $wert1 "s" $fa1
    2werte " Zugriff über Port 3269 möglich:" $wert2 "s" $fa2
    Leerzeile
}
function NTLM ($dcakt){
    $regkey = Invoke-Command -ComputerName $dcakt -ScriptBlock `
    { 
        Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\' `
         -ErrorAction SilentlyContinue
    }
    if($null -eq $regkey) { 
        $aktiv = "Reg-Key nicht vorhanden" ; $fa = "Red"
        } 
        else { 
            $aktiv = "Reg-Key vorhanden" ; $fa = "Green" 
            $wert = $regkey.lmcompatibilitylevel
            switch ($wert) {
                "0" { $fa2 = "Red" }
                "1" { $fa2 = "Red" }
                "2" { $fa2 = "Red" }
                "3" { $fa2 = "Red" }
                "4" { $fa2 = "Yellow" }
                "5" { $fa2 = "Green" }
                Default { $wert = "Fehler" ; $fa2 = "Red" }
            }
        }
    Bereichstitel "NTLM Einstellungen:" "s"
    Leerzeile
    2werte "Reg-Key vorhanden:" $aktiv "s" $fa
    Leerzeile
    2werte "Gesetzter Wert   :" $wert "s" $fa2
    Leerzeile
    2werte "Legende" "" "s"
    Leerzeile
    2werte " 0 =" "LM- und NTLM-Antworten senden" "s" "Red"
    2werte " 1 =" "LM- und NTLM-Antworten senden" "s" "Red"
    2werte "    " "(NTLMv2-Sitzungssicherheit verwenden)" "s" "Red"
    2werte " 2 =" "Nur NTLM-Antworten senden" "s" "Red"
    2werte " 3 =" "Nur NTLMv2 Antworten senden" "s" "Red"
    2werte " 4 =" "Nur NTLMv2 Antworten senden" "s" "Yellow"
    2werte "    " "LM verweigern" "s" "Yellow"
    2werte " 5 =" "Nur NTLMv2 Antworten senden" "s" "Green"
    2werte "    " "LM & NTLM verweigern" "s" "Green"
    Leerzeile
}
function dc_SMB1 ($DC) {
    Bereichstitel "SMB-Protokoll Status:" "s"
    Leerzeile
    2werte " SMB 1:" " " "s"
    $state_smb1 = (Get-WindowsFeature -ComputerName "$DC" -Name FS-SMB1).Installstate
    if ($state_smb1 -eq "Installed") { $farbe = $F_Fehler } else { $farbe = "Green" }
    2werte " - Installationsstatus  :" $state_smb1 "s" $farbe
    $zw = Get-SmbServerConfiguration -CimSession "$DC"
    if ($false -ne $zw.EnableSMB1Protocol) { $farbe = $F_Fehler } else { $farbe = "Green" }
    if ($zw.EnableSMB1Protocol -eq $false) { $te_smb1 = "deaktiviert"} 
        else { $te_smb1 = "aktiv" }
    2werte " - Status SMB1 Protokoll:" $te_smb1 "s" $farbe
    Leerzeile
    2werte " SMB 2:" " " "s"
    if ($false -eq $zw.EnableSMB2Protocol) { $farbe = $F_Fehler } else { $farbe = "Green" }
    if ($zw.EnableSMB2Protocol -eq $false) { $te_smb2 = "deaktiviert"} 
        else { $te_smb2 = "aktiv" }
    2werte " - Status SMB2 Protokoll:" $te_smb2 "s" $farbe
    Leerzeile
}
function Power ($sys) {
    $ps_a = New-Object 'object[,]' 5,2
    $ps_a[0,0] = Invoke-Command -ComputerName $sys -ScriptBlock {Get-ExecutionPolicy -Scope MachinePolicy}
    $ps_a[1,0] = Invoke-Command -ComputerName $sys -ScriptBlock {Get-ExecutionPolicy -Scope UserPolicy}
    $ps_a[2,0] = Invoke-Command -ComputerName $sys -ScriptBlock {Get-ExecutionPolicy -Scope Process}
    $ps_a[3,0] = Invoke-Command -ComputerName $sys -ScriptBlock {Get-ExecutionPolicy -Scope CurrentUser}
    $ps_a[4,0] = Invoke-Command -ComputerName $sys -ScriptBlock {Get-ExecutionPolicy -Scope LocalMachine}
    for($i=0;$i -lt 5;$i++){
        if($ps_a[$i,0] -eq 0) { $ps_a[$i,0] = "Unrestricted" ; $ps_a[$i,1] = "Red" }
        if($ps_a[$i,0] -eq 1) { $ps_a[$i,0] = "RemoteSigned" ; $ps_a[$i,1] = "Yellow" }
        if($ps_a[$i,0] -eq 2) { $ps_a[$i,0] = "AllSigned" ; $ps_a[$i,1] = "Green" }
        if($ps_a[$i,0] -eq 3) { $ps_a[$i,0] = "Restricted" ; $ps_a[$i,1] = "Green" }
        if($ps_a[$i,0] -eq 4) { $ps_a[$i,0] = "Bypass" ; $ps_a[$i,1] = "Red" }
        if($ps_a[$i,0] -eq 5) { $ps_a[$i,0] = "Undefined" ; $ps_a[$i,1] = "Red" }
    }
    Bereichstitel "PowerShell Policy Settings:" "s"
    Leerzeile
    2werte " MachinePolicy:" $ps_a[0,0] "s" $ps_a[0,1]
    2werte " UserPolicy   :" $ps_a[1,0] "s" $ps_a[1,1]
    2werte " Process      :" $ps_a[2,0] "s" $ps_a[2,1]
    2werte " CurrentUser  :" $ps_a[3,0] "s" $ps_a[3,1]
    2werte " LocalMachine :" $ps_a[4,0] "s" $ps_a[4,1]
    Leerzeile
}
function OF_Bitlocker ($DC) {
    $ofb = Invoke-Command -ComputerName $DC -ScriptBlock {
        $ofbtp = Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -like "*BitLo*"}
        return $ofbtp
    }
    Subtitel "BitLocker Feature:" "1" "*"
    Leerzeile
    $laenge = 0
    foreach ($of in $ofb) {
        $d_laenge = ($of.FeatureName).Length
        if ($d_laenge -gt $laenge) { $laenge = $d_laenge }
    }
    $laenge = $laenge + 1
    foreach ($of in $ofb) {
        [string]$Name = $of.FeatureName
        [string]$State = $of.State
        if ($state -eq "Enabled") { $ofbfa = "Green" } else { $ofbfa = "Yellow" }
        new_2werte "1" ":" $laenge " $Name" "" "l" $State $ofbfa "l"
    }
    Leerzeile
}
function AD_Controller {
    foreach ($dcon in $dcons) {
        Bereichstitel "Zum Domain Controller: $dcon"
        Leerzeile
        Power $dcon
        dc_SMB1 $dcon
        NTLM $dcon
        dc_ldaps $dcon
        OF_Bitlocker $dcon
        if($DomCon -eq 2) {
        dcroles $dcon
        dcfeature $dcon
        dchot $dcon
        dcprog $dcon
        dcdienste $dcon
        }
      }
}
####################################################################################################
# Sicherheit Paket A: Kerberos-Angriffsflaechen (read-only LDAP-Abfragen)                          #
####################################################################################################
function chk_kerberoasting {
    # Aktivierte Benutzerkonten (keine Computer) mit SPN -> Kerberoasting-faehig. krbtgt ausgenommen.
    $spn = @(Get-ADUser -LDAPFilter '(&(servicePrincipalName=*)(!userAccountControl:1.2.840.113556.1.4.803:=2))' `
                -Properties servicePrincipalName, adminCount |
             Where-Object { $_.SamAccountName -ne 'krbtgt' })
    $anz = $spn.Count
    if ($anz -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "Benutzerkonten mit SPN:" "$anz" "s" $fa
    $priv = @($spn | Where-Object { $_.adminCount -eq 1 })
    if ($priv.Count -gt 0) { 2werte "davon privilegiert (AdminCount=1):" "$($priv.Count)" "s" $F_Fehler }
    if ($anz -gt 0) {
        Leerzeile
        Bereichstitel "Betroffene Konten:" "s"
        Leerzeile
        foreach ($u in $spn) {
            $kennz = if ($u.adminCount -eq 1) { 'privilegiert!' } else { 'Standardkonto' }
            $cfa   = if ($u.adminCount -eq 1) { $F_Fehler } else { $F_Text }
            2werte " $($u.SamAccountName)" $kennz "s" $cfa
        }
    }
}
function chk_asrep {
    # Konten ohne Kerberos-Vorauthentifizierung (DONT_REQ_PREAUTH) -> AS-REP-Roasting-faehig.
    $asr = @(Get-ADUser -LDAPFilter '(userAccountControl:1.2.840.113556.1.4.803:=4194304)' `
                -Properties userAccountControl)
    $anz = $asr.Count
    if ($anz -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "Konten ohne Vorauthentifizierung:" "$anz" "s" $fa
    if ($anz -gt 0) {
        Leerzeile
        foreach ($u in $asr) { 2werte " $($u.SamAccountName)" "DONT_REQ_PREAUTH" "s" $F_Fehler }
    }
}
function chk_delegation {
    # DC-Namen zum Ausschluss (DCs haben legitim uneingeschraenkte Delegation).
    $dcNamen = @($DCs | ForEach-Object { $_.Name })
    # 1) Uneingeschraenkte Delegation (TRUSTED_FOR_DELEGATION = 0x80000 = 524288)
    $uncon = @(Get-ADObject -LDAPFilter '(userAccountControl:1.2.840.113556.1.4.803:=524288)' `
                -Properties samAccountName | Where-Object { $dcNamen -notcontains ($_.samAccountName -replace '\$$','') })
    if ($uncon.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "Uneingeschränkte Delegation (ohne DCs):" "$($uncon.Count)" "s" $fa
    foreach ($o in $uncon) { 2werte " $($o.samAccountName)" "TrustedForDelegation" "s" $F_Fehler }
    # 2) Eingeschraenkte Delegation (msDS-AllowedToDelegateTo gesetzt)
    $con = @(Get-ADObject -LDAPFilter '(msDS-AllowedToDelegateTo=*)' -Properties samAccountName)
    if ($con.Count -gt 0) { $fa = 'Yellow' } else { $fa = 'Green' }
    2werte "Eingeschränkte Delegation:" "$($con.Count)" "s" $fa
    foreach ($o in $con) { 2werte " $($o.samAccountName)" "AllowedToDelegateTo" "s" "Yellow" }
    # 3) Ressourcenbasierte Delegation (msDS-AllowedToActOnBehalfOfOtherIdentity gesetzt)
    $rbcd = @(Get-ADObject -LDAPFilter '(msDS-AllowedToActOnBehalfOfOtherIdentity=*)' -Properties samAccountName)
    if ($rbcd.Count -gt 0) { $fa = 'Yellow' } else { $fa = 'Green' }
    2werte "Ressourcenbasierte Delegation (RBCD):" "$($rbcd.Count)" "s" $fa
    foreach ($o in $rbcd) { 2werte " $($o.samAccountName)" "AllowedToActOnBehalfOf" "s" "Yellow" }
}
function chk_kerb_enc {
    # Konten, die ausschliesslich DES verwenden (UseDESKeyOnly = 0x200000 = 2097152).
    $des = @(Get-ADObject -LDAPFilter '(userAccountControl:1.2.840.113556.1.4.803:=2097152)' `
                -Properties samAccountName)
    if ($des.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "Konten mit 'nur DES' (UseDESKeyOnly):" "$($des.Count)" "s" $fa
    foreach ($o in $des) { 2werte " $($o.samAccountName)" "UseDESKeyOnly" "s" $F_Fehler }
}
function chk_machine_quota {
    $dn  = (Get-ADDomain).DistinguishedName
    $maq = (Get-ADObject -Identity $dn -Properties 'ms-DS-MachineAccountQuota').'ms-DS-MachineAccountQuota'
    if ($null -eq $maq) { $maq = 'n/a' }
    if ("$maq" -eq '0') { $fa = 'Green' } else { $fa = $F_Fehler }
    2werte "ms-DS-MachineAccountQuota:" "$maq" "s" $fa
    2werte "Empfohlener Wert:" "0" "s"
}
####################################################################################################
# Sicherheit Paket B: Privilegien & ACLs (read-only LDAP-/ACL-Abfragen)                            #
####################################################################################################
function chk_dcsync {
    # Wer hat Replikationsrechte (Get-Changes / Get-Changes-All) am Domaenenobjekt? -> DCSync.
    $domDN = (Get-ADDomain).DistinguishedName
    $acl   = Get-Acl -Path "AD:\$domDN"
    $rids  = @{ '1131f6aa-9c07-11d1-f79f-00c04fc2dcd2' = 'Get-Changes'
                '1131f6ad-9c07-11d1-f79f-00c04fc2dcd2' = 'Get-Changes-All' }
    # Standard-Berechtigte (Namens-Suffix), die Replikation legitim besitzen:
    $ok = 'Domain Controllers','Enterprise Read-Only Domain Controllers','Administrators',
          'SYSTEM','Enterprise Domain Controllers','Domain Admins','Enterprise Admins'
    $verdaechtig = @()
    foreach ($ace in $acl.Access) {
        if ($ace.AccessControlType -ne 'Allow') { continue }
        $guid = "$($ace.ObjectType)"
        if (-not $rids.ContainsKey($guid)) { continue }
        $kurz = ("$($ace.IdentityReference)" -split '\\')[-1]
        if ($ok -notcontains $kurz) {
            $verdaechtig += [pscustomobject]@{ Wer = "$($ace.IdentityReference)"; Recht = $rids[$guid] }
        }
    }
    if ($verdaechtig.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "Nicht-Standard-Prinzipale mit Replikationsrecht:" "$($verdaechtig.Count)" "s" $fa
    foreach ($v in $verdaechtig) { 2werte " $($v.Wer)" $v.Recht "s" $F_Fehler }
    if ($verdaechtig.Count -eq 0) { 2werte " Hinweis:" "nur Standard-Berechtigte (DCs/Admins)" "s" "Green" }
}
function chk_operatoren {
    $domSID = (Get-ADDomain).DomainSID.Value
    $gruppen = @(
        @{ N = 'Account Operators'; SID = 'S-1-5-32-548' }
        @{ N = 'Server Operators';  SID = 'S-1-5-32-549' }
        @{ N = 'Print Operators';   SID = 'S-1-5-32-550' }
        @{ N = 'Backup Operators';  SID = 'S-1-5-32-551' }
        @{ N = 'Schema Admins';     SID = "$domSID-518" }
        @{ N = 'Enterprise Admins'; SID = "$domSID-519" }
    )
    foreach ($g in $gruppen) {
        try {
            $m = @(Get-ADGroupMember -Identity $g.SID -ErrorAction Stop)
            if ($m.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
            2werte " $($g.N):" "$($m.Count) Mitglieder" "s" $fa
            foreach ($mm in $m) { 2werte "   - $($mm.SamAccountName)" "$($mm.objectClass)" "s" $F_Fehler }
        } catch {
            2werte " $($g.N):" "nicht abfragbar (evtl. Child-Domain)" "s" "Yellow"
        }
    }
    try {
        $dns = @(Get-ADGroupMember -Identity 'DnsAdmins' -ErrorAction Stop)
        if ($dns.Count -gt 0) { $fa = 'Yellow' } else { $fa = 'Green' }
        2werte " DnsAdmins:" "$($dns.Count) Mitglieder" "s" $fa
        foreach ($mm in $dns) { 2werte "   - $($mm.SamAccountName)" "$($mm.objectClass)" "s" "Yellow" }
    } catch { 2werte " DnsAdmins:" "nicht vorhanden/abfragbar" "s" "Green" }
}
function chk_adminsdholder {
    $domDN = (Get-ADDomain).DistinguishedName
    $acl   = Get-Acl -Path "AD:\CN=AdminSDHolder,CN=System,$domDN"
    # Nur uebernahme-relevante Rechte flaggen; reines (oft attributgebundenes) WriteProperty/
    # ReadProperty ist haeufig ein legitimes Default-ACE (z. B. Cert Publishers, TS License
    # Servers, Azure AD Connect) und wird bewusst NICHT als Befund gewertet.
    $gefaehrlich = 'GenericAll','GenericWrite','WriteDacl','WriteOwner'
    $ok = 'SYSTEM','Domain Admins','Enterprise Admins','Administrators','CREATOR OWNER','SELF'
    $auffaellig = @()
    foreach ($ace in $acl.Access) {
        if ($ace.AccessControlType -ne 'Allow') { continue }
        $rechte = "$($ace.ActiveDirectoryRights)"
        $hat = $false
        foreach ($r in $gefaehrlich) { if ($rechte -match $r) { $hat = $true; break } }
        if (-not $hat) { continue }
        $kurz = ("$($ace.IdentityReference)" -split '\\')[-1]
        if ($ok -notcontains $kurz) {
            $auffaellig += [pscustomobject]@{ Wer = "$($ace.IdentityReference)"; Recht = $rechte }
        }
    }
    if ($auffaellig.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "Nicht-Standard-Prinzipale mit Schreibrecht auf AdminSDHolder:" "$($auffaellig.Count)" "s" $fa
    foreach ($a in $auffaellig) { 2werte " $($a.Wer)" $a.Recht "s" $F_Fehler }
    if ($auffaellig.Count -eq 0) { 2werte " Hinweis:" "nur Standard-Prinzipale" "s" "Green" }
}
function chk_protected_users {
    $domSID = (Get-ADDomain).DomainSID.Value
    try {
        $pu = @(Get-ADGroupMember -Identity "$domSID-525" -ErrorAction Stop)
        if ($pu.Count -eq 0) {
            2werte "Protected Users - Mitglieder:" "0 (Gruppe ungenutzt)" "s" "Yellow"
        } else {
            2werte "Protected Users - Mitglieder:" "$($pu.Count)" "s" "Green"
            foreach ($mm in $pu) { 2werte "   - $($mm.SamAccountName)" "$($mm.objectClass)" "s" "Green" }
        }
    } catch { 2werte "Protected Users:" "nicht abfragbar" "s" "Yellow" }
}
function chk_prewin2000 {
    try {
        $pw = @(Get-ADGroupMember -Identity 'S-1-5-32-554' -ErrorAction Stop)
        $kritSids = @('S-1-1-0','S-1-5-7')   # Everyone, Anonymous Logon
        $krit = @($pw | Where-Object { $kritSids -contains "$($_.SID)" })
        if ($krit.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
        2werte "Pre-Windows 2000 - Mitglieder:" "$($pw.Count)" "s" $fa
        foreach ($mm in $pw) {
            if ($kritSids -contains "$($mm.SID)") { $mfa = $F_Fehler } else { $mfa = $F_Text }
            2werte "   - $($mm.SamAccountName)" "$($mm.SID)" "s" $mfa
        }
    } catch { 2werte "Pre-Windows 2000 Compatible Access:" "nicht abfragbar (ggf. Spezial-Identitäten)" "s" "Yellow" }
}
####################################################################################################
# Sicherheit Paket C: AD CS / ESC (read-only AD-Objekte + certutil)                                #
####################################################################################################
function Ist-NiedrigPriv ($idref) {
    # Breite/niedrig privilegierte Prinzipale (SID-basiert, sprachunabhaengig).
    try { $sid = $idref.Translate([System.Security.Principal.SecurityIdentifier]).Value } catch { $sid = "$idref" }
    if (@('S-1-5-11','S-1-1-0','S-1-5-7') -contains $sid) { return $true }   # Auth Users, Everyone, Anonymous
    if ($sid -match '-(513|515|545)$') { return $true }                       # Domain Users, Domain Computers, Users
    return $false
}
function Get-ADCSObjekte {
    $conf = (Get-ADRootDSE).configurationNamingContext
    $pks  = "CN=Public Key Services,CN=Services,$conf"
    $tmpl = @(Get-ADObject -SearchBase "CN=Certificate Templates,$pks" -LDAPFilter '(objectClass=pKICertificateTemplate)' `
              -Properties displayName,name,'msPKI-Certificate-Name-Flag','msPKI-Enrollment-Flag','msPKI-RA-Signature',pKIExtendedKeyUsage,'msPKI-Certificate-Application-Policy')
    $cas  = @(Get-ADObject -SearchBase "CN=Enrollment Services,$pks" -LDAPFilter '(objectClass=pKIEnrollmentService)' `
              -Properties name,dNSHostName,certificateTemplates)
    $pub  = @($cas | ForEach-Object { $_.certificateTemplates } | Where-Object { $_ } | Sort-Object -Unique)
    [pscustomobject]@{ Templates = $tmpl; CAs = $cas; Published = $pub }
}
function Get-NiedrigPrivEnroller ($dn) {
    # Liefert die breiten/niedrig privilegierten Prinzipale mit Enroll-Recht (leer = keine).
    $enrollGuid = '0e10c968-78fb-11d2-90d4-00c04f79dc55'   # Certificate-Enrollment
    $autoGuid   = 'a05b8cc2-17bc-4802-a710-e7c15ab866a2'   # Auto-Enrollment
    $res = @()
    try { $acl = Get-Acl -Path "AD:\$dn" } catch { return @() }
    foreach ($ace in $acl.Access) {
        if ($ace.AccessControlType -ne 'Allow') { continue }
        $g = "$($ace.ObjectType)"
        $rechte = "$($ace.ActiveDirectoryRights)"
        $istEnroll = ($g -eq $enrollGuid -or $g -eq $autoGuid -or $rechte -match 'GenericAll')
        if ($istEnroll -and (Ist-NiedrigPriv $ace.IdentityReference)) {
            $kurz = ("$($ace.IdentityReference)" -split '\\')[-1]
            if ($res -notcontains $kurz) { $res += $kurz }
        }
    }
    return $res
}
function chk_adcs_inventory {
    $o = Get-ADCSObjekte
    2werte "Zertifizierungsstellen (CAs):" "$($o.CAs.Count)" "s"
    foreach ($ca in $o.CAs) { 2werte " - $($ca.Name)" "$($ca.dNSHostName)" "s" }
    2werte "Vorlagen gesamt / veröffentlicht:" "$($o.Templates.Count) / $($o.Published.Count)" "s"
}
function chk_esc1 {
    $o = Get-ADCSObjekte
    $authEku = '1.3.6.1.5.5.7.3.2','1.3.6.1.4.1.311.20.2.2','1.3.6.1.5.2.3.4','2.5.29.37.0'
    $treffer = @()
    foreach ($t in $o.Templates) {
        if ($o.Published -notcontains $t.name) { continue }
        if (([int]$t.'msPKI-Certificate-Name-Flag' -band 1) -eq 0) { continue }   # kein ENROLLEE_SUPPLIES_SUBJECT
        if (([int]$t.'msPKI-Enrollment-Flag' -band 2) -ne 0) { continue }          # Manager-Approval
        if ([int]$t.'msPKI-RA-Signature' -gt 0) { continue }                        # Signaturen noetig
        $ekus = @($t.pKIExtendedKeyUsage) + @($t.'msPKI-Certificate-Application-Policy') | Where-Object { $_ }
        $hatAuth = ($ekus.Count -eq 0) -or (@($ekus | Where-Object { $authEku -contains $_ }).Count -gt 0)
        if (-not $hatAuth) { continue }
        $enroller = @(Get-NiedrigPrivEnroller $t.DistinguishedName)
        if ($enroller.Count -gt 0) { $treffer += [pscustomobject]@{ Name = "$($t.displayName)"; Enroll = ($enroller -join ', ') } }
    }
    if ($treffer.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "ESC1-verdächtige Vorlagen:" "$($treffer.Count)" "s" $fa
    foreach ($x in $treffer) {
        2werte " Vorlage: $($x.Name)" "ESC1 (Subject frei + Auth-EKU, kein Approval)" "s" $F_Fehler
        2werte "   Enroll für:" $x.Enroll "s" $F_Fehler
    }
    if ($treffer.Count -eq 0) { 2werte " Hinweis:" "keine ESC1-Vorlage gefunden" "s" "Green" }
}
function chk_esc2_3 {
    $o = Get-ADCSObjekte
    $anyPurpose = '2.5.29.37.0' ; $enrollAgent = '1.3.6.1.4.1.311.20.2.1'
    $tr = @()
    foreach ($t in $o.Templates) {
        if ($o.Published -notcontains $t.name) { continue }
        $ekus = @($t.pKIExtendedKeyUsage) + @($t.'msPKI-Certificate-Application-Policy') | Where-Object { $_ }
        $isAny = ($ekus.Count -eq 0) -or ($ekus -contains $anyPurpose)
        $isAgent = $ekus -contains $enrollAgent
        if (-not ($isAny -or $isAgent)) { continue }
        $enroller = @(Get-NiedrigPrivEnroller $t.DistinguishedName)
        if ($enroller.Count -gt 0) {
            $typ = if ($isAgent) { 'ESC3 (Enrollment Agent)' } else { 'ESC2 (Any Purpose/kein EKU)' }
            $tr += [pscustomobject]@{ Name = "$($t.displayName)"; Typ = $typ; Enroll = ($enroller -join ', ') }
        }
    }
    if ($tr.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "ESC2/ESC3-verdächtige Vorlagen:" "$($tr.Count)" "s" $fa
    foreach ($x in $tr) {
        2werte " Vorlage: $($x.Name)" $x.Typ "s" $F_Fehler
        2werte "   Enroll für:" $x.Enroll "s" $F_Fehler
    }
    if ($tr.Count -eq 0) { 2werte " Hinweis:" "keine ESC2/ESC3-Vorlage gefunden" "s" "Green" }
}
function chk_esc4 {
    $o = Get-ADCSObjekte
    $gefaehr = 'GenericAll','GenericWrite','WriteDacl','WriteOwner','WriteProperty'
    $tr = @()
    foreach ($t in $o.Templates) {
        try { $acl = Get-Acl -Path "AD:\$($t.DistinguishedName)" } catch { continue }
        foreach ($ace in $acl.Access) {
            if ($ace.AccessControlType -ne 'Allow') { continue }
            $rechte = "$($ace.ActiveDirectoryRights)"
            $hat = $false ; foreach ($r in $gefaehr) { if ($rechte -match $r) { $hat = $true; break } }
            if ($hat -and (Ist-NiedrigPriv $ace.IdentityReference)) {
                $tr += [pscustomobject]@{ Name = "$($t.displayName)"; Wer = ("$($ace.IdentityReference)" -split '\\')[-1]; Recht = $rechte }
            }
        }
    }
    if ($tr.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "ESC4 (Vorlagen mit Schreibrecht für breite Gruppen):" "$($tr.Count)" "s" $fa
    foreach ($x in $tr) { 2werte " Vorlage: $($x.Name)" "Schreibrecht von $($x.Wer): $($x.Recht)" "s" $F_Fehler }
    if ($tr.Count -eq 0) { 2werte " Hinweis:" "keine ESC4-Vorlage gefunden" "s" "Green" }
}
function chk_esc6 {
    $o = Get-ADCSObjekte
    if ($o.CAs.Count -eq 0) { 2werte "ESC6 (SAN-Flag):" "keine CA gefunden" "s" "Green"; return }
    foreach ($ca in $o.CAs) {
        $cfg = "$($ca.dNSHostName)\$($ca.Name)"
        try {
            $out = certutil -config $cfg -getreg policy\EditFlags 2>&1 | Out-String
            if ($out -match 'EDITF_ATTRIBUTESUBJECTALTNAME2') { $fa = $F_Fehler; $st = 'GESETZT (gefährlich)' }
            else { $fa = 'Green'; $st = 'nicht gesetzt' }
            2werte " $cfg" $st "s" $fa
        } catch { 2werte " $cfg" "nicht abfragbar (CA/Recht)" "s" "Yellow" }
    }
}
function chk_esc8 {
    $o = Get-ADCSObjekte
    if ($o.CAs.Count -eq 0) { 2werte "ESC8 (Web Enrollment):" "keine CA gefunden" "s" "Green"; return }
    foreach ($ca in $o.CAs) {
        $h = $ca.dNSHostName
        try {
            $feat = Invoke-Command -ComputerName $h -ScriptBlock { (Get-WindowsFeature ADCS-Web-Enrollment).Installed } -ErrorAction Stop
            if ($feat) { $fa = $F_Fehler; $st = 'installiert (NTLM-Relay-Ziel)' } else { $fa = 'Green'; $st = 'nicht installiert' }
            2werte " $h" $st "s" $fa
        } catch { 2werte " $h" "nicht abfragbar (WinRM/Recht)" "s" "Yellow" }
    }
}
####################################################################################################
# Sicherheit Paket D: GPO/SYSVOL-Geheimnisse (read-only Datei-/ACL-Abfragen)                       #
####################################################################################################
function Entschluessle-GPP ($cpassword) {
    # GPP cpassword mit dem oeffentlich bekannten AES-256-Schluessel (MS14-025) entschluesseln.
    try {
        $rest = $cpassword.Length % 4
        if ($rest -ne 0) { $cpassword += ('=' * (4 - $rest)) }
        $bytes = [Convert]::FromBase64String($cpassword)
        $key = [byte[]](0x4e,0x99,0x06,0xe8,0xfc,0xb6,0x6c,0xc9,0xfa,0xf4,0x93,0x10,0x62,0x0f,0xfe,0xe8,
                        0xf4,0x96,0xe8,0x06,0xcc,0x05,0x79,0x90,0x20,0x9b,0x09,0xa4,0x33,0xb6,0x6c,0x1b)
        $aes = New-Object System.Security.Cryptography.AesManaged
        $aes.Key = $key ; $aes.IV = New-Object byte[] 16 ; $aes.Mode = 'CBC' ; $aes.Padding = 'PKCS7'
        $dec = $aes.CreateDecryptor()
        $out = $dec.TransformFinalBlock($bytes, 0, $bytes.Length)
        return [System.Text.Encoding]::Unicode.GetString($out)
    } catch { return '<nicht entschlüsselbar>' }
}
function chk_gpp_cpassword {
    $dom = (Get-ADDomain).DNSRoot
    $pol = "\\$dom\SYSVOL\$dom\Policies"
    $namen = 'Groups.xml','Services.xml','ScheduledTasks.xml','DataSources.xml','Printers.xml','Drives.xml'
    $treffer = @()
    $dateien = @(Get-ChildItem -Path $pol -Recurse -Include $namen -ErrorAction SilentlyContinue)
    foreach ($d in $dateien) {
        $inhalt = Get-Content -LiteralPath $d.FullName -Raw -ErrorAction SilentlyContinue
        if (-not $inhalt) { continue }
        foreach ($m in [regex]::Matches($inhalt, 'cpassword="([^"]+)"')) {
            $cp = $m.Groups[1].Value
            if ([string]::IsNullOrEmpty($cp)) { continue }
            $treffer += [pscustomobject]@{ Datei = $d.FullName.Replace($pol, '...'); Klar = (Entschluessle-GPP $cp) }
        }
    }
    if ($treffer.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "GPP-cpassword-Fundstellen:" "$($treffer.Count)" "s" $fa
    foreach ($x in $treffer) {
        2werte " Datei: $($x.Datei)" "cpassword gefunden" "s" $F_Fehler
        2werte "   Entschlüsselt:" $x.Klar "s" $F_Fehler
    }
    if ($treffer.Count -eq 0) { 2werte " Hinweis:" "keine GPP-Passwörter im SYSVOL" "s" "Green" }
}
function chk_sysvol_scripts {
    $dom = (Get-ADDomain).DNSRoot
    $basis = "\\$dom\SYSVOL\$dom"
    $muster = 'net use\s.+/user:', '/user:\S+\s+\S', 'password\s*[:=]', '-AsPlainText',
              'ConvertTo-SecureString', 'psexec.+\s-p\s', 'pwd\s*[:=]', 'passwd\s*[:=]'
    $treffer = @()
    $dateien = @(Get-ChildItem -Path $basis -Recurse -Include *.bat,*.cmd,*.ps1,*.vbs,*.kix -ErrorAction SilentlyContinue)
    foreach ($d in $dateien) {
        $inhalt = Get-Content -LiteralPath $d.FullName -Raw -ErrorAction SilentlyContinue
        if (-not $inhalt) { continue }
        foreach ($mu in $muster) {
            if ($inhalt -match $mu) {
                $treffer += [pscustomobject]@{ Datei = $d.FullName.Replace($basis, '...'); Muster = $mu }
                break
            }
        }
    }
    if ($treffer.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "Verdächtige Skripte (heuristisch):" "$($treffer.Count)" "s" $fa
    foreach ($x in $treffer) { 2werte " $($x.Datei)" "Muster: $($x.Muster)" "s" $F_Fehler }
    if ($treffer.Count -eq 0) { 2werte " Hinweis:" "keine auffälligen Skripte gefunden" "s" "Green" }
}
function chk_gpo_rights {
    $domDN = (Get-ADDomain).DistinguishedName
    $gpos  = @(Get-ADObject -SearchBase "CN=Policies,CN=System,$domDN" `
               -LDAPFilter '(objectClass=groupPolicyContainer)' -Properties displayName)
    $gefaehr = 'GenericAll','GenericWrite','WriteDacl','WriteOwner','WriteProperty'
    $tr = @()
    foreach ($g in $gpos) {
        try { $acl = Get-Acl -Path "AD:\$($g.DistinguishedName)" } catch { continue }
        foreach ($ace in $acl.Access) {
            if ($ace.AccessControlType -ne 'Allow') { continue }
            $rechte = "$($ace.ActiveDirectoryRights)"
            $hat = $false ; foreach ($r in $gefaehr) { if ($rechte -match $r) { $hat = $true; break } }
            if ($hat -and (Ist-NiedrigPriv $ace.IdentityReference)) {
                $tr += [pscustomobject]@{ GPO = "$($g.displayName)"; Wer = ("$($ace.IdentityReference)" -split '\\')[-1]; Recht = $rechte }
            }
        }
    }
    if ($tr.Count -gt 0) { $fa = $F_Fehler } else { $fa = 'Green' }
    2werte "GPOs mit Schreibrecht für breite Gruppen:" "$($tr.Count)" "s" $fa
    foreach ($x in $tr) { 2werte " GPO: $($x.GPO)" "Schreibrecht von $($x.Wer): $($x.Recht)" "s" $F_Fehler }
    if ($tr.Count -eq 0) { 2werte " Hinweis:" "keine GPO mit breitem Schreibrecht" "s" "Green" }
}
####################################################################################################
# Sicherheit Paket E: DC-Haertung vertieft (read-only Registry/Dienste je DC + dSHeuristics)       #
####################################################################################################
function chk_ldap_signing {
    foreach ($dc in $DCs) {
        $h = $dc.HostName
        try {
            $reg = Invoke-Command -ComputerName $h -ScriptBlock {
                $p = 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters'
                [pscustomobject]@{
                    Integrity = (Get-ItemProperty -Path $p -Name LDAPServerIntegrity -ErrorAction SilentlyContinue).LDAPServerIntegrity
                    CBT       = (Get-ItemProperty -Path $p -Name LdapEnforceChannelBinding -ErrorAction SilentlyContinue).LdapEnforceChannelBinding
                }
            } -ErrorAction Stop
            if ("$($reg.Integrity)" -eq '2') { $sig = 'erforderlich'; $sfa = 'Green' } else { $sig = 'NICHT erforderlich'; $sfa = $F_Fehler }
            switch ("$($reg.CBT)") {
                '2'     { $cbt = 'immer';             $cfa = 'Green' }
                '1'     { $cbt = 'wenn unterstützt'; $cfa = 'Yellow' }
                default { $cbt = 'aus/nicht gesetzt'; $cfa = $F_Fehler }
            }
            2werte " $($dc.Name) - LDAP-Signing:" $sig "s" $sfa
            2werte "   Channel Binding:" $cbt "s" $cfa
        } catch { 2werte " $($dc.Name):" "nicht abfragbar (WinRM/Recht)" "s" "Yellow" }
    }
}
function chk_smb_signing {
    foreach ($dc in $DCs) {
        $h = $dc.HostName
        try {
            $req = Invoke-Command -ComputerName $h -ScriptBlock {
                (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters' `
                    -Name RequireSecuritySignature -ErrorAction SilentlyContinue).RequireSecuritySignature
            } -ErrorAction Stop
            if ("$req" -eq '1') { $st = 'erforderlich'; $fa = 'Green' } else { $st = 'NICHT erforderlich'; $fa = $F_Fehler }
            2werte " $($dc.Name) - SMB-Signing (Server):" $st "s" $fa
        } catch { 2werte " $($dc.Name):" "nicht abfragbar (WinRM/Recht)" "s" "Yellow" }
    }
}
function chk_print_spooler {
    foreach ($dc in $DCs) {
        $h = $dc.HostName
        try {
            $svc = Get-Service -ComputerName $h -Name Spooler -ErrorAction Stop
            if ($svc.Status -eq 'Running') { $st = 'läuft (PrinterBug/PetitPotam-Risiko)'; $fa = $F_Fehler } else { $st = "$($svc.Status)"; $fa = 'Green' }
            2werte " $($dc.Name) - Print Spooler:" $st "s" $fa
        } catch { 2werte " $($dc.Name):" "nicht abfragbar" "s" "Yellow" }
    }
}
function chk_anon_ldap {
    try {
        $conf = (Get-ADRootDSE).configurationNamingContext
        $dsDN = "CN=Directory Service,CN=Windows NT,CN=Services,$conf"
        $heur = (Get-ADObject -Identity $dsDN -Properties dSHeuristics).dSHeuristics
        if ([string]::IsNullOrEmpty($heur)) {
            2werte "dSHeuristics:" "nicht gesetzt (Standard)" "s" "Green"
            2werte "Anonyme LDAP-Binds:" "nicht erlaubt (Standard)" "s" "Green"
            return
        }
        2werte "dSHeuristics:" "$heur" "s"
        if ($heur.Length -ge 7 -and $heur[6] -eq '2') {
            2werte "Anonyme LDAP-Binds:" "ERLAUBT (7. Zeichen = 2)" "s" $F_Fehler
        } else {
            2werte "Anonyme LDAP-Binds:" "nicht erlaubt" "s" "Green"
        }
    } catch { 2werte "dSHeuristics:" "nicht abfragbar" "s" "Yellow" }
}
####################################################################################################
## Paket F - Delta-Modus: Vergleich mit früherem JSON-Export                                      ##
####################################################################################################
function Extrahiere-Befunde ($ereignisse) {
    # Bildet aus einer Ereignisliste die Menge der auffaelligen Befunde (rot/gelb).
    # Rueckgabe: Hashtable (Schluessel = normalisierter Befundtext) fuer schnellen Mengenvergleich.
    $set = @{}
    if ($null -eq $ereignisse) { return $set }
    foreach ($e in $ereignisse) {
        if ($e.Art -ne 'Wert') { continue }
        $farbe = "$($e.Farbe)"
        if ($farbe -notmatch '^(Dark)?(Red|Yellow)$') { continue }     # nur Befund-Farben
        $name = "$($e.Name)".Trim()
        # eigene Delta-Ausgaben frueherer Laeufe nicht erneut als Befund werten:
        if ($name -match '^(\s*[+\-]\s|Neu hinzugekommen|Behoben|Vergleich mit|Befunde \(alt|Unveraendert|Hinweis:)') { continue }
        $wert = "$($e.Wert)".Trim()
        $key  = if ($wert) { "$name | $wert" } else { $name }
        $set[$key] = $true
    }
    return $set
}
function chk_delta ($altPfad) {
    if ([string]::IsNullOrWhiteSpace($altPfad)) { return }
    if (-not (Test-Path -LiteralPath $altPfad)) {
        2werte "Vergleichsdatei:" "nicht gefunden ($altPfad)" "s" "Yellow"; return
    }
    try {
        $alt = Get-Content -LiteralPath $altPfad -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        2werte "Vergleichsdatei:" "nicht lesbar (kein gültiges JSON)" "s" "Yellow"; return
    }
    $altEreignisse = if ($alt.PSObject.Properties.Name -contains 'Ereignisse') { $alt.Ereignisse } else { $alt }
    if ($R_Daten.Count -eq 0) {
        2werte "Hinweis:" "Aktueller Lauf ohne Export - Delta nicht möglich (HTML/JSON aktiv lassen)" "s" "Yellow"; return
    }
    $alteBefunde = Extrahiere-Befunde $altEreignisse
    $neueBefunde = Extrahiere-Befunde $R_Daten
    $neu     = @($neueBefunde.Keys | Where-Object { -not $alteBefunde.ContainsKey($_) } | Sort-Object)
    $behoben = @($alteBefunde.Keys | Where-Object { -not $neueBefunde.ContainsKey($_) } | Sort-Object)
    $bleibt  = @($neueBefunde.Keys | Where-Object {       $alteBefunde.ContainsKey($_) }).Count

    2werte "Vergleich mit:" (Split-Path -Leaf $altPfad) "s"
    2werte "Befunde (alt / aktuell):" "$($alteBefunde.Count) / $($neueBefunde.Count)" "s"
    2werte "Unveraendert:" "$bleibt" "s"
    Leerzeile
    2werte "Neu hinzugekommen:" "$($neu.Count)" "s" $(if ($neu.Count -gt 0) { $F_Fehler } else { "Green" })
    foreach ($x in $neu)     { 2werte "  + $x" "" "s" $F_Fehler }
    Leerzeile
    2werte "Behoben (nicht mehr vorhanden):" "$($behoben.Count)" "s" $(if ($behoben.Count -gt 0) { "Green" } else { "White" })
    foreach ($x in $behoben) { 2werte "  - $x" "" "s" "Green" }
}
####################################################################################################
####################################################################################################
##### Main                                                                                     #####
####################################################################################################
####################################################################################################
## Bereich Header                                                                                 ##
####################################################################################################
Clear-Host
Header $type $maintitel
Puffer_leeren
####################################################################################################
## Bereich Domain, Mode, FSMO                                                                     ##
####################################################################################################
if($domoco -ge 1){
    Pruefbereich "Domain, Mode, FSMO" -CheckId 'domain_allgemein' {
        Leerzeile
        dom_allgemein
    }
}
####################################################################################################
## Bereich "Central Store & Templates"                                                            ##
####################################################################################################
if($censto -ge 1 -or $sectem -ge 1){
    Pruefbereich "Central Store & Templates" -CheckId 'central_store' {
        if($censto -ge 1){ centralstore }
        if($sectem -ge 1){ sec_templates }
    }
}
####################################################################################################
## Bereich Domain Controller                                                                      ##
####################################################################################################
if ($domdcs -ge 1) {
    Pruefbereich "Domain Controller" -CheckId 'domain_controller' {
        Leerzeile
        controller
    }
}
####################################################################################################
## Bereich "Logging auf Domain Controller(n)"                                                     ##
####################################################################################################
if($loggin -ge 1){
    Pruefbereich "Logging auf Domain Controller(n)" -CheckId 'logging' {
        Leerzeile
        Event_dienst
        Auditcheck
    }
}
####################################################################################################
## Bereich "AD-Trusts - Check"                                                                    ##
####################################################################################################
if($adtchk -ge 1){
    Pruefbereich "AD-Trusts - Check" -CheckId 'trusts' {
        Leerzeile
        trusts
    }
}
####################################################################################################
## Bereich "DNS - Check"                                                                          ##
####################################################################################################
if($dnschk -ge 1){
    Pruefbereich "DNS - Check" -CheckId 'dns' {
        Leerzeile
        aging
    }
}
####################################################################################################
## Bereich "Sysvol Replication & AD-Health"                                                       ##
####################################################################################################
if($SysRep -ge 1){
    Pruefbereich "Sysvol Replication & AD-Health" -CheckId 'sysvol_health' {
        Leerzeile
        dfsr
        controller_check
    }
}
####################################################################################################
## Bereich "Administratoren und Builtin Benutzer"                                                 ##
####################################################################################################
if($admusr -ge 1){
    Pruefbereich "Administratoren und Builtin Benutzer" -CheckId 'admins' {
        Leerzeile
        Admins
        if($lokadm -eq 1) { lokale_AdmGru }
        if($AdmGri -eq 1) { dom_AdmGri }
        if($buildi -eq 1) { builtin_usr }
        if($priusr -eq 1) { AdmCount }
    }
}
####################################################################################################
## Bereich "Kerberos - Angriffsflächen" (Paket A)                                                ##
####################################################################################################
if($kerbchk -ge 1){
    Pruefbereich "Kerberos - Angriffsflächen" -CheckId 'kerberos' {
        Leerzeile
        Unterpruefung "Kerberoasting (Konten mit SPN)" 'kerberoasting' { chk_kerberoasting }
        Unterpruefung "AS-REP Roasting (ohne Vorauthentifizierung)" 'asrep' { chk_asrep }
        Unterpruefung "Delegation (Unconstrained / Constrained / RBCD)" 'delegation' { chk_delegation }
        Unterpruefung "Schwache Kerberos-Verschlüsselung" 'kerb_enc' { chk_kerb_enc }
        Unterpruefung "Computerkonten-Kontingent (MachineAccountQuota)" 'machine_quota' { chk_machine_quota }
    }
}
####################################################################################################
## Bereich "Privilegien & ACLs" (Paket B)                                                         ##
####################################################################################################
if($privchk -ge 1){
    Pruefbereich "Privilegien & ACLs" -CheckId 'privilegien' {
        Leerzeile
        Unterpruefung "DCSync-Rechte (Verzeichnis-Replikation)" 'dcsync' { chk_dcsync }
        Unterpruefung "Gefährliche Builtin-/Operatoren-Gruppen" 'operatoren' { chk_operatoren }
        Unterpruefung "AdminSDHolder-ACL" 'adminsdholder' { chk_adminsdholder }
        Unterpruefung "Protected Users (Nutzung)" 'protected_users' { chk_protected_users }
        Unterpruefung "Pre-Windows 2000 Compatible Access" 'prewin2000' { chk_prewin2000 }
    }
}
####################################################################################################
## Bereich "AD CS - Zertifikatsdienste (ESC)" (Paket C)                                            ##
####################################################################################################
if($adcschk -ge 1){
    Pruefbereich "AD CS - Zertifikatsdienste (ESC)" -CheckId 'adcs' {
        Leerzeile
        Unterpruefung "Bestand (CAs und Vorlagen)" $null { chk_adcs_inventory }
        Unterpruefung "ESC1 (Enrollee Supplies Subject + Auth-EKU)" 'esc1' { chk_esc1 }
        Unterpruefung "ESC2/ESC3 (Any Purpose / Enrollment Agent)" 'esc2_3' { chk_esc2_3 }
        Unterpruefung "ESC4 (manipulierbare Vorlagen-ACL)" 'esc4' { chk_esc4 }
        Unterpruefung "ESC6 (EDITF_ATTRIBUTESUBJECTALTNAME2 auf CA)" 'esc6' { chk_esc6 }
        Unterpruefung "ESC8 (HTTP Web Enrollment)" 'esc8' { chk_esc8 }
    }
}
####################################################################################################
## Bereich "GPO & SYSVOL - Geheimnisse" (Paket D)                                                  ##
####################################################################################################
if($sysvchk -ge 1){
    Pruefbereich "GPO & SYSVOL - Geheimnisse" -CheckId 'gpo_sysvol' {
        Leerzeile
        Unterpruefung "GPP-Passwörter (cpassword in SYSVOL)" 'gpp_cpassword' { chk_gpp_cpassword }
        Unterpruefung "Klartext-Credentials in SYSVOL-Skripten" 'sysvol_scripts' { chk_sysvol_scripts }
        Unterpruefung "GPO-Bearbeitungsrechte" 'gpo_rights' { chk_gpo_rights }
    }
}
####################################################################################################
## Bereich "DC-Härtung (vertieft)" (Paket E)                                                      ##
####################################################################################################
if($dchaert -ge 1){
    Pruefbereich "DC-Härtung (vertieft)" -CheckId 'dc_haertung' {
        Leerzeile
        Unterpruefung "LDAP-Signing und Channel Binding" 'ldap_signing' { chk_ldap_signing }
        Unterpruefung "SMB-Signing (erforderlich)" 'smb_signing' { chk_smb_signing }
        Unterpruefung "Print Spooler auf DCs" 'print_spooler' { chk_print_spooler }
        Unterpruefung "Anonyme LDAP-Binds (dSHeuristics)" 'anon_ldap' { chk_anon_ldap }
    }
}
####################################################################################################
## Bereich "Benutzer und Benutzer Accounts"                                                       ##
####################################################################################################
if ($usrchk -eq 1) {
    Pruefbereich "Benutzer und Benutzer Accounts" -CheckId 'benutzer' {
        Leerzeile
        User_chk
        if($inachk -eq 1) { inaktive_User }
        if($geschk -eq 1) { gesperrte_User }
        if($falchk -eq 1) { ou_users }
    }
}
####################################################################################################
## Bereich "Computerkonten Check"                                                                 ##
####################################################################################################
if ($syschk -eq 1) {
    Pruefbereich 'Computerkonten Check' -CheckId 'computerkonten' {
        Leerzeile
        sys_konten
    }
}
####################################################################################################
## Bereich "Client Check"                                                                         ##
####################################################################################################
if($cltchk -ge 1){
    Pruefbereich "Client Check" -CheckId 'clients' {
        Leerzeile
        clt_chk
    }
}
####################################################################################################
## Bereich "Client Check"                                                                         ##
####################################################################################################
if($srvchk -ge 1){
    Pruefbereich "Server Check" -CheckId 'server' {
        Leerzeile
        srv_chk
    }
}
####################################################################################################
## Bereich "Nicht Windows Systeme"                                                                ##
####################################################################################################
if($no_win -eq 1){
    Pruefbereich 'Nicht "Windows" Systeme' -CheckId 'nicht_windows' {
        Leerzeile
        oth_chk
    }
}
####################################################################################################
## Bereich "AD-Gruppen"                                                                           ##
####################################################################################################
if($allgru -ge 1){
    Pruefbereich "AD-Gruppen" -CheckId 'ad_gruppen' {
        Leerzeile
        ad_gruppen
    }
}
####################################################################################################
## Bereich "GPO's"                                                                                ##
####################################################################################################
if($allgpo -ge 1){
    Pruefbereich "GPO's" -CheckId 'gpos' {
        Leerzeile
        GPO_all
        Leerzeile
    }
}
####################################################################################################
## Bereich "dDP Password Settings"                                                                ##
####################################################################################################
if($dDPchk -ge 1){
    Pruefbereich "dDP Password Settings" -CheckId 'ddp_password' {
        Leerzeile
        ddomainpol
        #spezial_user
    }
}
####################################################################################################
## Bereich "fGPP - fine Grained Password Policies"                                                ##
####################################################################################################
if($fgppch -ge 1){
    Pruefbereich "fine Grained Password Policies" -CheckId 'fgpp' {
        Leerzeile
        fGPO
    }
}
####################################################################################################
## Bereich "User vs Password Policies"                                                            ##
####################################################################################################
if($userpw -ge 1){
    Pruefbereich "User vs Password Policies" -CheckId 'user_vs_pw' {
        Leerzeile
        spezial_user
    }
}
####################################################################################################
## Bereich "Organisation Units"                                                                   ##
####################################################################################################
if($OrgUni -ge 1){
    Pruefbereich "Organisation Units" -CheckId 'ous' {
        Leerzeile
        OUS
    }
}
####################################################################################################
## Bereich "DACL, Rechte-Delegierung"                                                             ##
####################################################################################################
#if($aclchk -ge 1){
#    Bereich "DACL, Rechte-Delegierung"
#    Leerzeile
#    dacls
#    Leerzeile
#}
####################################################################################################
## Bereich "Managed Service Accounts (MSA/gMSA)"                                                  ##
####################################################################################################
if($manacc -ge 1){
    Pruefbereich "Managed Service Accounts (MSA/gMSA)" -CheckId 'msa' {
        Leerzeile
        KDSR
        MSA
        gMSA
    }
}
####################################################################################################
## Bereich "Zertifizierungsstelle(n)"                                                             ##
####################################################################################################
if($caschk -ge 1){
    Pruefbereich "Zertifizierungsstelle(n)" -CheckId 'ca' {
        Leerzeile
        ca_root
        ca_sub
        #ca_templates
    }
}
####################################################################################################
## Bereich "Domain Controller"                                                                    ##
####################################################################################################
if ($DomCon -ge 1) {
    Pruefbereich "Domain Controller" -CheckId 'dc_detail' {
        Leerzeile
        AD_Controller
    }
}
####################################################################################################
## Bereich "Veränderungen seit letztem Lauf (Delta)" (Paket F)                                    ##
####################################################################################################
if ($deltchk -ge 1 -and $Vergleich) {
    Pruefbereich "Veränderungen seit letztem Lauf (Delta)" -CheckId 'delta' {
        Leerzeile
        # CheckId $null: die Doku/Begruendung kommt bereits vom Pruefbereich 'delta' -
        # so erscheint der Delta-Block nicht doppelt in Exec-Summary und Bericht.
        Unterpruefung "Vergleich mit früherem JSON-Export" $null { chk_delta $Vergleich }
    }
}
####################################################################################################
## Bereich Abschluss Script                                                                       ##
####################################################################################################
bottom
Puffer_leeren
####################################################################################################
## Zusatz-Ausgaben: HTML-Report und JSON-Export                                                   ##
####################################################################################################
if ($A_Htm -eq 1) { HTML_Report }
if ($A_Jsn -eq 1) { JSON_Export }
####################################################################################################
## Ende                                                                                           ##
####################################################################################################
