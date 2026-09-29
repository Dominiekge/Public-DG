# =====================================================================
#  Obsidian Attachment Cleaner  (v2 - robuuster)
#  Opslaan als: ObsidianAttachmentCleaner.ps1  (let op: niet .ps1.txt!)
#  Starten met dit commando in PowerShell of een .bat-bestand:
#    powershell -NoExit -ExecutionPolicy Bypass -File "C:\pad\naar\ObsidianAttachmentCleaner.ps1"
#  Door -NoExit blijft het consolevenster open, ook bij een foutmelding.
# =====================================================================

$ErrorActionPreference = 'Stop'

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic
} catch {
    Read-Host "Kon de benodigde .NET-onderdelen niet laden: $($_.Exception.Message). Druk op Enter om af te sluiten"
    exit 1
}

$ImageExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.bmp')
$script:Results = @()

# ---------------------------------------------------------------
#  Logfunctie
# ---------------------------------------------------------------
function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format 'HH:mm:ss'
    $txtLog.AppendText("[$stamp] $Message`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------
#  Filter: gebruikte bijlagen verbergen als dat is aangevinkt
# ---------------------------------------------------------------
function Apply-Filter {
    if ($null -eq $chkOnlyUnused) { return }
    $onlyUnused = $chkOnlyUnused.Checked
    foreach ($row in $grid.Rows) {
        $entry = $script:Results[[int]$row.Tag]
        $row.Visible = (-not $onlyUnused) -or (-not $entry.Used)
    }
}

# ---------------------------------------------------------------
#  Scan Vault
# ---------------------------------------------------------------
function Invoke-Scan {
    try {
        $script:Results = @()
        $grid.Rows.Clear()
        $lblStats.Text = "MD-bestanden: 0  |  Bijlagen: 0  |  Ongebruikt: 0  |  Opslag: 0 MB"

        $vaultPath = $txtVaultPath.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($vaultPath) -or -not (Test-Path -LiteralPath $vaultPath)) {
            [System.Windows.Forms.MessageBox]::Show("Vul eerst een geldig pad naar je Obsidian-vault in.", "Obsidian Attachment Cleaner", 'OK', 'Warning') | Out-Null
            Write-Log "Scan afgebroken: geen geldig vault-pad."
            return
        }

        $caseSensitive  = $chkCaseSensitive.Checked
        $searchSubfolders = $chkSubfolders.Checked
        $imageFolder = $txtImageFolder.Text.Trim()
        $pdfFolder   = $txtPdfFolder.Text.Trim()

        Write-Log "Scan gestart van: $vaultPath"
        if ($searchSubfolders) { Write-Log "Submappen doorzoeken: AAN" } else { Write-Log "Submappen doorzoeken: UIT (alle ingevulde mappen)" }
        if ($caseSensitive) { Write-Log "Hoofdlettergevoelig zoeken: AAN" } else { Write-Log "Hoofdlettergevoelig zoeken: UIT" }

        # --- 1. Alle markdownbestanden lezen (platte tekst + YAML samen) ---
        $mdFiles = @(Get-ChildItem -LiteralPath $vaultPath -Filter '*.md' -File -Recurse)
        Write-Log "Gevonden: $($mdFiles.Count) markdownbestanden."

        $mdContent = New-Object System.Text.StringBuilder
        foreach ($md in $mdFiles) {
            try {
                $raw = [System.IO.File]::ReadAllText($md.FullName)
                if (-not $caseSensitive) { $raw = $raw.ToLowerInvariant() }
                [void]$mdContent.AppendLine($raw)
            } catch {
                Write-Log "Kon niet lezen: $($md.FullName)"
            }
        }
        $allContent = $mdContent.ToString()

        # --- 2. Bijlagen verzamelen ---
        $attachments = @()
        if ($searchSubfolders) {
            $attachments = @(Get-ChildItem -LiteralPath $vaultPath -File -Recurse | Where-Object {
                ($ImageExtensions -contains $_.Extension.ToLowerInvariant()) -or ($_.Extension.ToLowerInvariant() -eq '.pdf') })
        } else {
            $folders = @((Join-Path $vaultPath $imageFolder), (Join-Path $vaultPath $pdfFolder))
            foreach ($f in $folders) {
                if (Test-Path -LiteralPath $f) {
                    $attachments += @(Get-ChildItem -LiteralPath $f -File | Where-Object {
                        ($ImageExtensions -contains $_.Extension.ToLowerInvariant()) -or ($_.Extension.ToLowerInvariant() -eq '.pdf') })
                }
            }
        }
        $attachments = @($attachments | Where-Object { $_.FullName -notmatch '[\\/]\.(obsidian|trash)[\\/]' })
        Write-Log "Gevonden: $($attachments.Count) bijlagen (afbeeldingen en PDF's)."

        # --- 3. Per bijlage checken of de bestandsnaam ergens in een notitie staat ---
        $unusedCount = 0
        $unusedBytes = [long]0
        $rootLength = $vaultPath.TrimEnd('\').Length

        foreach ($att in $attachments) {
            $name = $att.Name
            if (-not $caseSensitive) { $name = $name.ToLowerInvariant() }

            # Zoek de letterlijke bestandsnaam in alle notitietekst:
            # dekt [[Logo.png]], [[Bestand.pdf]] en losse vermeldingen zoals
            # "Negenvlaksmodel 1.png" in platte tekst of YAML.
            $isUsed = $allContent.Contains($name)

            $type = 'Afbeelding'
            if ($att.Extension.ToLowerInvariant() -eq '.pdf') { $type = 'PDF' }

            if (-not $isUsed) {
                $unusedCount++
                $unusedBytes += $att.Length
            }

            $resultIndex = $script:Results.Count
            $script:Results += [PSCustomObject]@{ FileInfo = $att; Type = $type; Used = $isUsed }

            $relPath = $att.FullName.Substring($rootLength).TrimStart('\')
            $sizeKB = [math]::Round($att.Length / 1KB, 1)
            $idx = $grid.Rows.Add($false, $type, $att.Name, $relPath, "$sizeKB KB")
            $grid.Rows[$idx].Tag = $resultIndex

            # Gebruikte bijlagen: grijs, checkbox op alleen-lezen
            if ($isUsed) {
                $grid.Rows[$idx].Cells[0].ReadOnly = $true
                $grid.Rows[$idx].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
        }

        $unusedMB = [math]::Round($unusedBytes / 1MB, 2)
        $lblStats.Text = "MD-bestanden: $($mdFiles.Count)  |  Bijlagen: $($attachments.Count)  |  Ongebruikt: $unusedCount  |  Opslag: $unusedMB MB"
        Write-Log "Scan klaar. Ongebruikte bijlagen: $unusedCount ($unusedMB MB)."
        Write-Log "Tip: aangevinkte bestanden gaan naar de prullenbak, niet definitief weg."
        Apply-Filter
    }
    catch {
        Write-Log "FOUT tijdens scan: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Er ging iets mis tijdens de scan:`r`n$($_.Exception.Message)", "Fout", 'OK', 'Error') | Out-Null
    }
}

# ---------------------------------------------------------------
#  Verwijder Geselecteerde
# ---------------------------------------------------------------
function Invoke-Remove {
    try {
        if ($grid.Rows.Count -eq 0) {
            Write-Log "Er is niets om te verwijderen. Doe eerst een scan."
            return
        }
        $selected = @($grid.Rows | Where-Object { $_.Cells[0].Value -eq $true })
        if ($selected.Count -eq 0) {
            Write-Log "Geen bestanden aangevinkt."
            return
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Je verplaatst $($selected.Count) bestand(en) naar de prullenbak. Doorgaan?",
            "Verwijderen", 'YesNo', 'Question')
        if ($answer -ne 'Yes') { Write-Log "Verwijderen geannuleerd."; return }

        $removed = 0
        foreach ($row in $selected) {
            $file = $script:Results[[int]$row.Tag].FileInfo
            try {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                    $file.FullName, 'OnlyErrorDialogs', 'SendToRecycleBin')
                Write-Log "Verwijderd: $($file.Name)"
                $removed++
            } catch {
                Write-Log "Mislukt: $($file.Name) - $($_.Exception.Message)"
            }
        }
        Write-Log "Klaar: $removed bestand(en) naar de prullenbak verplaatst."

        for ($i = $grid.Rows.Count - 1; $i -ge 0; $i--) {
            if ($grid.Rows[$i].Cells[0].Value -eq $true) { $grid.Rows.RemoveAt($i) }
        }
    }
    catch {
        Write-Log "FOUT tijdens verwijderen: $($_.Exception.Message)"
    }
}

function Set-AllChecks {
    param([bool]$Value)
    foreach ($row in $grid.Rows) {
        if (-not $row.Cells[0].ReadOnly) { $row.Cells[0].Value = $Value }
    }
}

# Verberg gebruikte rijen als "Alleen ongebruikte" is aangevinkt
# (de functie staat bovenaan in dit bestand, vóór alle aanroepen)

# ---------------------------------------------------------------
#  Het venster - eigenschappen een voor een toegewezen
# ---------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Obsidian Attachment Cleaner"
$form.Size = New-Object System.Drawing.Size(900, 720)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# ---- Blauwe kop ----
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 50
$header.BackColor = [System.Drawing.Color]::FromArgb(45, 108, 180)
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Obsidian Attachment Cleaner"
$lblTitle.ForeColor = 'White'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point -ArgumentList 12, (12)
$header.Controls.Add($lblTitle)
$form.Controls.Add($header)

$y = 60

# ---- Sectie: Instellingen ----
$lblSettings = New-Object System.Windows.Forms.Label
$lblSettings.Text = "Instellingen"
$lblSettings.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$lblSettings.AutoSize = $true
$lblSettings.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y)
$form.Controls.Add($lblSettings)
$y += 26

$lblVault = New-Object System.Windows.Forms.Label
$lblVault.Text = "Obsidian Vault Path:"
$lblVault.AutoSize = $true
$lblVault.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y + 3)
$form.Controls.Add($lblVault)

$txtVaultPath = New-Object System.Windows.Forms.TextBox
$txtVaultPath.Location = New-Object System.Drawing.Point -ArgumentList 160, ($y)
$txtVaultPath.Width = 500
$form.Controls.Add($txtVaultPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Bladeren..."
$btnBrowse.Location = New-Object System.Drawing.Point -ArgumentList 670, ($y - 2)
$btnBrowse.Width = 90
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Kies de map van je Obsidian-vault"
    if ($dlg.ShowDialog($form) -eq 'OK') { $txtVaultPath.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnBrowse)
$y += 34

$lblImg = New-Object System.Windows.Forms.Label
$lblImg.Text = "Afbeeldingen map:"
$lblImg.AutoSize = $true
$lblImg.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y + 3)
$form.Controls.Add($lblImg)

$txtImageFolder = New-Object System.Windows.Forms.TextBox
$txtImageFolder.Text = "images"
$txtImageFolder.Location = New-Object System.Drawing.Point -ArgumentList 160, ($y)
$txtImageFolder.Width = 300
$form.Controls.Add($txtImageFolder)
$y += 30

$lblPdf = New-Object System.Windows.Forms.Label
$lblPdf.Text = "PDF's map:"
$lblPdf.AutoSize = $true
$lblPdf.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y + 3)
$form.Controls.Add($lblPdf)

$txtPdfFolder = New-Object System.Windows.Forms.TextBox
$txtPdfFolder.Text = "pdfs"
$txtPdfFolder.Location = New-Object System.Drawing.Point -ArgumentList 160, ($y)
$txtPdfFolder.Width = 300
$form.Controls.Add($txtPdfFolder)
$y += 34

$chkSubfolders = New-Object System.Windows.Forms.CheckBox
$chkSubfolders.Text = "Submappen doorzoeken"
$chkSubfolders.Checked = $true
$chkSubfolders.AutoSize = $true
$chkSubfolders.Location = New-Object System.Drawing.Point -ArgumentList 160, ($y)
$form.Controls.Add($chkSubfolders)

$chkCaseSensitive = New-Object System.Windows.Forms.CheckBox
$chkCaseSensitive.Text = "Hoofdlettergevoelig"
$chkCaseSensitive.Checked = $false
$chkCaseSensitive.AutoSize = $true
$chkCaseSensitive.Location = New-Object System.Drawing.Point -ArgumentList 360, ($y)
$form.Controls.Add($chkCaseSensitive)

$chkOnlyUnused = New-Object System.Windows.Forms.CheckBox
$chkOnlyUnused.Text = "Alleen ongebruikte bijlagen tonen"
$chkOnlyUnused.Checked = $false
$chkOnlyUnused.AutoSize = $true
$chkOnlyUnused.Location = New-Object System.Drawing.Point -ArgumentList 500, ($y)
# Direct filteren bij aan/uitvinken (alleen na een scan)
$chkOnlyUnused.Add_Click({ Apply-Filter })
$form.Controls.Add($chkOnlyUnused)
$y += 36

# ---- Actieknoppen ----
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan Vault"
$btnScan.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y)
$btnScan.Width = 160
$btnScan.Height = 30
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
$btnScan.Add_Click({ Invoke-Scan })
$form.Controls.Add($btnScan)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Verwijder Geselecteerde"
$btnRemove.Location = New-Object System.Drawing.Point -ArgumentList 182, ($y)
$btnRemove.Width = 160
$btnRemove.Height = 30
$btnRemove.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
$btnRemove.Add_Click({ Invoke-Remove })
$form.Controls.Add($btnRemove)

$btnSelAll = New-Object System.Windows.Forms.Button
$btnSelAll.Text = "Selecteer Alles"
$btnSelAll.Location = New-Object System.Drawing.Point -ArgumentList 352, ($y)
$btnSelAll.Width = 160
$btnSelAll.Height = 30
$btnSelAll.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
$btnSelAll.Add_Click({ Set-AllChecks $true })
$form.Controls.Add($btnSelAll)

$btnSelNone = New-Object System.Windows.Forms.Button
$btnSelNone.Text = "Deselecteer Alles"
$btnSelNone.Location = New-Object System.Drawing.Point -ArgumentList 522, ($y)
$btnSelNone.Width = 160
$btnSelNone.Height = 30
$btnSelNone.BackColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
$btnSelNone.Add_Click({ Set-AllChecks $false })
$form.Controls.Add($btnSelNone)
$y += 40

# ---- Resultaten ----
$lblStats = New-Object System.Windows.Forms.Label
$lblStats.Text = "MD-bestanden: 0  |  Bijlagen: 0  |  Ongebruikt: 0  |  Opslag: 0 MB"
$lblStats.AutoSize = $true
$lblStats.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblStats.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y)
$form.Controls.Add($lblStats)
$y += 24

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y)
$grid.Size = New-Object System.Drawing.Size(860, 250)
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSizeColumnsMode = 'Fill'

$chkCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$chkCol.Name = 'chk'
$chkCol.HeaderText = ''
$chkCol.Width = 40
[void]$grid.Columns.Add($chkCol)

$colType = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colType.Name = 'type'; $colType.HeaderText = 'Type'; $colType.ReadOnly = $true
[void]$grid.Columns.Add($colType)

$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = 'name'; $colName.HeaderText = 'Bestandsnaam'; $colName.ReadOnly = $true
[void]$grid.Columns.Add($colName)

$colPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colPath.Name = 'path'; $colPath.HeaderText = 'Pad'; $colPath.ReadOnly = $true
[void]$grid.Columns.Add($colPath)

$colSize = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSize.Name = 'size'; $colSize.HeaderText = 'Grootte'; $colSize.ReadOnly = $true
[void]$grid.Columns.Add($colSize)

$form.Controls.Add($grid)
$y += 260

# ---- Log ----
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log"
$lblLog.AutoSize = $true
$lblLog.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$lblLog.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y)
$form.Controls.Add($lblLog)
$y += 22

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Location = New-Object System.Drawing.Point -ArgumentList 12, ($y)
$txtLog.Size = New-Object System.Drawing.Size(860, 100)
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

Write-Log "Klaar voor gebruik. Kies je vault-map en klik op 'Scan Vault'."

[void]$form.ShowDialog()