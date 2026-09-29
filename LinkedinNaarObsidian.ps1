# =============================================================
#  LinkedIn Connections -> Obsidian persoon-notities
#  Een klein Windows-venster: kies je CSV-export en je vault,
#  de notities worden aangemaakt (bestaande worden overgeslagen).
#
#  GEBRUIK:
#  1. Sla dit bestand op als:  LinkedinNaarObsidian.ps1
#  2. Klik met de rechtermuisknop op het bestand -> "Uitvoeren met PowerShell"
#     (of: rechtermuisknop -> Eigenschappen -> vink "Deblokkeren" aan als dat er staat)
#  3. Blader naar je LinkedIn-export (opslaan als CSV!) en je Obsidian-vault
#
#  Werkt op elke Windows-pc, niets te installeren.
# =============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Instellingen ----------
$NaamEmoji = [char]0xD83D + [char]0xDC68   # mannemoji ; voor vrouw: 0xD83D 0xDC69

# Map waarin gecontroleerd wordt of een persoon al bestaat:
$PersoonMap = "D:\PKM Dominiek\004 - Bronnen\Persoon"

# Nieuwe notities komen in een verse map op je bureaublad:
$UitvoerMap = Join-Path ([Environment]::GetFolderPath("Desktop")) `
    ("LinkedIn import " + (Get-Date -Format "yyyy-MM-dd HH.mm"))

function Lees-Export {
    param($pad)
    # Leest een CSV direct; een XLSX via Excel zelf (COM).
    if ($pad -like "*.csv") {
        return @(Import-Csv -Path $pad -Encoding UTF8)
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $boek = $excel.Workbooks.Open($pad)
        $blad = $boek.Worksheets.Item(1)
        $laatsteRij = $blad.Cells.Item($blad.Rows.Count, 1).End(-4162).Row   # -4162 = xlUp

        # kolomnamen zoeken (LinkedIn zet uitleg bovenaan)
        $headerRij = 0
        for ($r = 1; $r -le [Math]::Min(15, $laatsteRij); $r++) {
            $cel = "$($blad.Cells.Item($r, 1).Text)".Trim()
            if ($cel -eq "First Name") { $headerRij = $r; break }
        }
        if ($headerRij -eq 0) { throw "Kolomnamen (First Name, Last Name, ...) niet gevonden." }

        # Aantal kolommen bepalen OP DE HEADER-RIJ zelf (de uitleg-toptekst
        # bovenaan beslaat maar 1 kolom, dus daar mag je niet naar kijken)
        $laatsteKol = $blad.Cells.Item($headerRij, $blad.Columns.Count).End(-4159).Column  # -4159 = xlToLeft

        $namen = @{}
        for ($k = 1; $k -le $laatsteKol; $k++) {
            $kolomnaam = "$($blad.Cells.Item($headerRij, $k).Text)".Trim()
            if ($kolomnaam -and $kolomnaam -ne "Notes:") { $namen[$k] = $kolomnaam }
        }

        $rijen = @()
        for ($r = $headerRij + 1; $r -le $laatsteRij; $r++) {
            $obj = @{}
            foreach ($k in $namen.Keys) {
                $obj[$namen[$k]] = "$($blad.Cells.Item($r, $k).Text)".Trim()
            }
            $rijen += [PSCustomObject]$obj
        }
        return $rijen
    }
    finally {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}

function Maak-Notitie {
    param($naam, $bedrijf, $functie, $url)

    # Emoji's via tekencodes opbouwen: als je dit script opslaat zonder BOM,
    # leest Windows PowerShell letterlijke emoji's namelijk als rare tekens
    # (bv. "gr-investeren"). Met tekencodes kan dat nooit meer misgaan.
    $hek       = '```'                       # backticks: geen escape in enkele quotes
    $link      = [string][char]0xD83D + [char]0xDD17   # 🔗
    $hamer     = [string][char]0xD83D + [char]0xDD28   # 🔨
    $fotoEmoji = [string][char]0xD83D + [char]0xDC64   # 👤
    $orgEmoji  = [string][char]0xD83D + [char]0xDE9F   # 🛗 (organisatie-icoon)

    if ($bedrijf) {
        $organisatie = "[[$orgEmoji $bedrijf]]"
    } else {
        $organisatie = ""
    }
    $fotoBestand = "[[$fotoEmoji $naam.jpg]]"

    @"
---
Type: Persoon
Organisatie: $organisatie
Afdeling:
Team:
Functie: $functie
Woonplaats:
tags:
  - persoon
Verjaardag:
Foto: $fotoBestand
Kring:
---

$link LinkedIn: $url

$hek"dataview
TABLE type AS Type
FROM ""
WHERE contains(file.outlinks, this.file.link)
SORT file.mtime DESC
$hek"

## $hamer Openstaande taken

$hek" dataview
TASK
FROM ""
WHERE contains(text, this.file.name)
$hek"
"@
}

# ---------- Scherm ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "LinkedIn naar Obsidian"
$form.Size = New-Object System.Drawing.Size(640, 460)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# CSV-bestand (of xlsx)
$lblCsv = New-Object System.Windows.Forms.Label
$lblCsv.Text = "LinkedIn-export (CSV of XLSX):"
$lblCsv.Location = New-Object System.Drawing.Point(15, 18)
$lblCsv.AutoSize = $true
$form.Controls.Add($lblCsv)

$txtCsv = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = New-Object System.Drawing.Point(15, 38)
$txtCsv.Size = New-Object System.Drawing.Size(450, 22)
$form.Controls.Add($txtCsv)

$btnCsv = New-Object System.Windows.Forms.Button
$btnCsv.Text = "Bladeren..."
$btnCsv.Location = New-Object System.Drawing.Point(475, 36)
$form.Controls.Add($btnCsv)
$btnCsv.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "LinkedIn-export (*.csv;*.xlsx)|*.csv;*.xlsx|CSV-bestand (*.csv)|*.csv|Excel-bestand (*.xlsx)|*.xlsx"
    if ($dlg.ShowDialog() -eq "OK") { $txtCsv.Text = $dlg.FileName }
})

# Locatie-uitslag in het scherm
$lblMap = New-Object System.Windows.Forms.Label
$lblMap.Text = "Bestaande personen gecontroleerd in:`r`n$PersoonMap`r`n`r`nNieuwe notities komen in:`r`n$UitvoerMap"
$lblMap.Location = New-Object System.Drawing.Point(15, 75)
$lblMap.Size = New-Object System.Drawing.Size(600, 60)
$form.Controls.Add($lblMap)

# Startknop
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Notities aanmaken"
$btnStart.Location = New-Object System.Drawing.Point(15, 145)
$btnStart.Size = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($btnStart)

# Voortgang
$txtUitvoer = New-Object System.Windows.Forms.TextBox
$txtUitvoer.Location = New-Object System.Drawing.Point(15, 190)
$txtUitvoer.Size = New-Object System.Drawing.Size(595, 230)
$txtUitvoer.Multiline = $true
$txtUitvoer.ScrollBars = "Vertical"
$txtUitvoer.ReadOnly = $true
$txtUitvoer.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtUitvoer)

$btnStart.Add_Click({
    $txtUitvoer.Clear()
    $csv = $txtCsv.Text.Trim()

    if (-not $csv -or -not (Test-Path $csv)) {
        [System.Windows.Forms.MessageBox]::Show("Kies eerst een geldig CSV- of XLSX-bestand.", "Ontbreekt")
        return
    }
    if (-not (Test-Path $PersoonMap)) {
        [System.Windows.Forms.MessageBox]::Show("De controle-map bestaat niet:`r`n$PersoonMap", "Ontbreekt")
        return
    }

    # Bestand uitlezen: CSV direct, XLSX via Excel
    try {
        $rijen = Lees-Export -pad $csv
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Kon het bestand niet lezen: $($_.Exception.Message)", "Fout")
        return
    }

    # Nieuwe map op het bureaublad aanmaken
    New-Item -ItemType Directory -Path $UitvoerMap -Force | Out-Null

    $aangemaakt = 0
    $overgeslagen = 0
    $txtUitvoer.AppendText("Gevonden personen: $($rijen.Count)`r`n`r`n")

    foreach ($rij in $rijen) {
        # Naam samenstellen uit alle naamkolommen die het bestand heeft
        # (First Name, Middle Name, Last Name, Name ... in die volgorde)
        $naamdelen = @()
        foreach ($kolom in @("First Name", "Middle Name", "Last Name", "Name", "Full Name")) {
            if ($rij.PSObject.Properties.Name -contains $kolom) {
                $deel = "$($rij.$kolom)".Trim()
                if ($deel -and $deel -ne $kolom) { $naamdelen += $deel }
            }
        }
        $naam = ($naamdelen -join " ").Trim()
        if (-not $naam) { continue }   # rij zonder naam overslaan
        $bestandsnaam = "$NaamEmoji $naam.md"

        # EERST CHECKEN: bestaat de persoon al in de Persoon-map?
        $bestaand = Get-ChildItem -Path $PersoonMap -Filter $bestandsnaam -File -ErrorAction SilentlyContinue
        if ($bestaand) {
            $txtUitvoer.AppendText("OVERGESLAGEN (bestaat al): $bestandsnaam`r`n")
            $overgeslagen++
            continue
        }

        $inhoud = Maak-Notitie -naam $naam `
            -bedrijf "$($rij.'Company')".Trim() `
            -functie "$($rij.'Position')".Trim() `
            -url "$($rij.'URL')".Trim()

        $pad = Join-Path $UitvoerMap $bestandsnaam
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($pad, $inhoud, $utf8NoBom)
        $txtUitvoer.AppendText("AANGEMAAKT: $bestandsnaam`r`n")
        $aangemaakt++
    }

    $txtUitvoer.AppendText("`r`nKlaar! Aangemaakt: $aangemaakt, overgeslagen: $overgeslagen.`r`n`r`nDe nieuwe notities staan in:`r`n$UitvoerMap")
})

[void]$form.ShowDialog()