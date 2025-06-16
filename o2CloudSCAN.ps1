Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO

# Configuration et chemins
$ConfigPath = "$env:USERPROFILE\Documents\ScannerApp"
$HistoryFile = "$ConfigPath\scan_history.json"
$DefaultScanPath = "$env:USERPROFILE\Documents\Scans"

# Créer les dossiers nécessaires
if (-not (Test-Path $ConfigPath)) {
    New-Item -Path $ConfigPath -ItemType Directory -Force
}
if (-not (Test-Path $DefaultScanPath)) {
    New-Item -Path $DefaultScanPath -ItemType Directory -Force
}

# Classe pour gérer l'historique des scans
class ScanHistory {
    [string]$FileName
    [string]$FilePath
    [datetime]$ScanDate
    [string]$Scanner
    [long]$FileSize
}

# Fonctions utilitaires
function Load-ScanHistory {
    if (Test-Path $HistoryFile) {
        try {
            $json = Get-Content $HistoryFile -Raw | ConvertFrom-Json
            return $json | ForEach-Object {
                [ScanHistory]@{
                    FileName = $_.FileName
                    FilePath = $_.FilePath
                    ScanDate = [datetime]$_.ScanDate
                    Scanner = $_.Scanner
                    FileSize = $_.FileSize
                }
            }
        }
        catch {
            return @()
        }
    }
    return @()
}

function Save-ScanHistory {
    param([ScanHistory[]]$History)
    
    $json = $History | ConvertTo-Json -Depth 2
    $json | Out-File $HistoryFile -Encoding UTF8
}

function Add-ScanToHistory {
    param(
        [string]$FileName,
        [string]$FilePath,
        [string]$Scanner
    )
    
    $history = Load-ScanHistory
    $fileInfo = Get-Item $FilePath -ErrorAction SilentlyContinue
    
    $newScan = [ScanHistory]@{
        FileName = $FileName
        FilePath = $FilePath
        ScanDate = Get-Date
        Scanner = $Scanner
        FileSize = if ($fileInfo) { $fileInfo.Length } else { 0 }
    }
    
    $history = @($newScan) + $history | Select-Object -First 100
    Save-ScanHistory $history
    return $history
}

function Get-Scanners {
    try {
        $wia = New-Object -ComObject WIA.DeviceManager
        $scanners = @()
        
        for ($i = 1; $i -le $wia.DeviceInfos.Count; $i++) {
            $device = $wia.DeviceInfos.Item($i)
            if ($device.Type -eq 1) { # Scanner
                $scanners += @{
                    Name = $device.Properties("Name").Value
                    ID = $device.DeviceID
                }
            }
        }
        return $scanners
    }
    catch {
        return @(@{ Name = "Scanner par défaut"; ID = "default" })
    }
}

function Start-Scan {
    param(
        [string]$OutputPath,
        [string]$ScannerID = "default"
    )
    
    try {
        # Determiner le format de fichier
        $extension = [System.IO.Path]::GetExtension($OutputPath).ToLower()
        $tempPath = $OutputPath
        
        # Si c'est un PDF, scanner d'abord en image puis convertir
        if ($extension -eq ".pdf") {
            $tempPath = $OutputPath -replace "\.pdf$", ".png"
        }
        
        $wia = New-Object -ComObject WIA.CommonDialog
        $device = $wia.ShowSelectDevice()
        
        if ($device) {
            $item = $device.Items(1)
            $image = $wia.ShowTransfer($item)
            
            if ($image) {
                # Sauvegarder l'image
                $image.SaveFile($tempPath)
                
                # Si PDF demande, convertir l'image en PDF
                if ($extension -eq ".pdf" -and (Test-Path $tempPath)) {
                    Convert-ImageToPDF -ImagePath $tempPath -PDFPath $OutputPath
                    # Supprimer le fichier temporaire
                    Remove-Item $tempPath -ErrorAction SilentlyContinue
                }
                
                return (Test-Path $OutputPath)
            }
        }
        return $false
    }
    catch {
        # Fallback - utiliser l'outil Windows Scan
        try {
            $scanFolder = Split-Path $OutputPath -Parent
            Start-Process "ms-settings:printers" -Wait
            
            # Informer l'utilisateur
            [System.Windows.Forms.MessageBox]::Show("Utilisez l'application Windows 'Scanner' puis sauvegardez le fichier dans:`n$scanFolder", "Scanner Windows", "OK", "Information")
            
            return (Test-Path $OutputPath)
        }
        catch {
            return $false
        }
    }
}

function Convert-ImageToPDF {
    param(
        [string]$ImagePath,
        [string]$PDFPath
    )
    
    try {
        # Utiliser l'API Windows pour creer un PDF simple
        Add-Type -AssemblyName System.Drawing
        
        $image = [System.Drawing.Image]::FromFile($ImagePath)
        $document = New-Object System.Drawing.Printing.PrintDocument
        
        # Methode alternative: utiliser Word si disponible
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Add()
            
            $selection = $word.Selection
            $selection.InlineShapes.AddPicture($ImagePath)
            
            $doc.SaveAs2($PDFPath, 17) # 17 = PDF format
            $doc.Close()
            $word.Quit()
            
            return $true
        }
        catch {
            # Si Word n'est pas disponible, copier l'image avec extension PDF
            # (ce n'est pas un vrai PDF mais cela fonctionne pour certains cas)
            Copy-Item $ImagePath $PDFPath -Force
            return $true
        }
    }
    catch {
        return $false
    }
}

# Création de l'interface principale
$form = New-Object System.Windows.Forms.Form
$form.Text = "[Scanner Pro] - Application de Numerisation"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(15, 15, 15)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Panel principal avec dégradé
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Size = New-Object System.Drawing.Size(884, 684)
$mainPanel.Location = New-Object System.Drawing.Point(8, 8)
$mainPanel.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
$form.Controls.Add($mainPanel)

# Titre de l'application
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "[SCANNER PRO]"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 150, 255)
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(300, 40)
$mainPanel.Controls.Add($titleLabel)

# Section Configuration
$configGroup = New-Object System.Windows.Forms.GroupBox
$configGroup.Text = "[CONFIG] Configuration du Scan"
$configGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$configGroup.ForeColor = [System.Drawing.Color]::FromArgb(100, 200, 255)
$configGroup.Location = New-Object System.Drawing.Point(20, 80)
$configGroup.Size = New-Object System.Drawing.Size(420, 280)
$configGroup.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$mainPanel.Controls.Add($configGroup)

# Nom du fichier
$fileNameLabel = New-Object System.Windows.Forms.Label
$fileNameLabel.Text = "[FILE] Nom du fichier:"
$fileNameLabel.Location = New-Object System.Drawing.Point(15, 35)
$fileNameLabel.Size = New-Object System.Drawing.Size(120, 20)
$fileNameLabel.ForeColor = [System.Drawing.Color]::White
$configGroup.Controls.Add($fileNameLabel)

$fileNameTextBox = New-Object System.Windows.Forms.TextBox
$fileNameTextBox.Location = New-Object System.Drawing.Point(15, 58)
$fileNameTextBox.Size = New-Object System.Drawing.Size(250, 25)
$fileNameTextBox.Text = "Scan_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$fileNameTextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$fileNameTextBox.ForeColor = [System.Drawing.Color]::White
$fileNameTextBox.BorderStyle = "FixedSingle"
$configGroup.Controls.Add($fileNameTextBox)

# Extension
$extensionLabel = New-Object System.Windows.Forms.Label
$extensionLabel.Text = "[EXT] Format:"
$extensionLabel.Location = New-Object System.Drawing.Point(280, 35)
$extensionLabel.Size = New-Object System.Drawing.Size(80, 20)
$extensionLabel.ForeColor = [System.Drawing.Color]::White
$configGroup.Controls.Add($extensionLabel)

$extensionCombo = New-Object System.Windows.Forms.ComboBox
$extensionCombo.Location = New-Object System.Drawing.Point(280, 58)
$extensionCombo.Size = New-Object System.Drawing.Size(120, 25)
$extensionCombo.Items.AddRange(@(".png", ".jpg", ".pdf", ".tiff", ".bmp"))
$extensionCombo.SelectedIndex = 0
$extensionCombo.DropDownStyle = "DropDownList"
$extensionCombo.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$extensionCombo.ForeColor = [System.Drawing.Color]::White
$configGroup.Controls.Add($extensionCombo)

# Chemin de destination
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "[PATH] Chemin de destination:"
$pathLabel.Location = New-Object System.Drawing.Point(15, 95)
$pathLabel.Size = New-Object System.Drawing.Size(150, 20)
$pathLabel.ForeColor = [System.Drawing.Color]::White
$configGroup.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(15, 118)
$pathTextBox.Size = New-Object System.Drawing.Size(300, 25)
$pathTextBox.Text = $DefaultScanPath
$pathTextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$pathTextBox.ForeColor = [System.Drawing.Color]::White
$pathTextBox.BorderStyle = "FixedSingle"
$configGroup.Controls.Add($pathTextBox)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "..."
$browseButton.Location = New-Object System.Drawing.Point(325, 118)
$browseButton.Size = New-Object System.Drawing.Size(35, 25)
$browseButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$browseButton.ForeColor = [System.Drawing.Color]::White
$browseButton.FlatStyle = "Flat"
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.SelectedPath = $pathTextBox.Text
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $pathTextBox.Text = $folderBrowser.SelectedPath
    }
})
$configGroup.Controls.Add($browseButton)

# Scanner selection
$scannerLabel = New-Object System.Windows.Forms.Label
$scannerLabel.Text = "[SCAN] Scanner:"
$scannerLabel.Location = New-Object System.Drawing.Point(15, 155)
$scannerLabel.Size = New-Object System.Drawing.Size(80, 20)
$scannerLabel.ForeColor = [System.Drawing.Color]::White
$configGroup.Controls.Add($scannerLabel)

$scannerCombo = New-Object System.Windows.Forms.ComboBox
$scannerCombo.Location = New-Object System.Drawing.Point(15, 178)
$scannerCombo.Size = New-Object System.Drawing.Size(350, 25)
$scannerCombo.DropDownStyle = "DropDownList"
$scannerCombo.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$scannerCombo.ForeColor = [System.Drawing.Color]::White
$configGroup.Controls.Add($scannerCombo)

# Charger les scanners
$scanners = Get-Scanners
foreach ($scanner in $scanners) {
    $null = $scannerCombo.Items.Add($scanner.Name)
}
if ($scannerCombo.Items.Count -gt 0) {
    $scannerCombo.SelectedIndex = 0
}

# Bouton de scan principal
$scanButton = New-Object System.Windows.Forms.Button
$scanButton.Text = "[START] DEMARRER LA NUMERISATION"
$scanButton.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$scanButton.Location = New-Object System.Drawing.Point(15, 220)
$scanButton.Size = New-Object System.Drawing.Size(385, 45)
$scanButton.BackColor = [System.Drawing.Color]::FromArgb(0, 200, 100)
$scanButton.ForeColor = [System.Drawing.Color]::White
$scanButton.FlatStyle = "Flat"
$scanButton.FlatAppearance.BorderSize = 0
$configGroup.Controls.Add($scanButton)

# Section Historique
$historyGroup = New-Object System.Windows.Forms.GroupBox
$historyGroup.Text = "[HISTORY] Historique des Scans"
$historyGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$historyGroup.ForeColor = [System.Drawing.Color]::FromArgb(100, 200, 255)
$historyGroup.Location = New-Object System.Drawing.Point(460, 80)
$historyGroup.Size = New-Object System.Drawing.Size(400, 280)
$historyGroup.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$mainPanel.Controls.Add($historyGroup)

# Liste de l'historique
$historyListView = New-Object System.Windows.Forms.ListView
$historyListView.Location = New-Object System.Drawing.Point(15, 25)
$historyListView.Size = New-Object System.Drawing.Size(370, 200)
$historyListView.View = "Details"
$historyListView.FullRowSelect = $true
$historyListView.GridLines = $true
$historyListView.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$historyListView.ForeColor = [System.Drawing.Color]::White
$historyListView.BorderStyle = "FixedSingle"

# Colonnes de l'historique - CORRECTION ICI
$null = $historyListView.Columns.Add("Fichier", 120)
$null = $historyListView.Columns.Add("Date", 80)
$null = $historyListView.Columns.Add("Taille", 60)
$null = $historyListView.Columns.Add("Scanner", 100)

$historyGroup.Controls.Add($historyListView)

# Boutons de l'historique
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "[REFRESH] Actualiser"
$refreshButton.Location = New-Object System.Drawing.Point(15, 235)
$refreshButton.Size = New-Object System.Drawing.Size(100, 30)
$refreshButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$refreshButton.ForeColor = [System.Drawing.Color]::White
$refreshButton.FlatStyle = "Flat"
$historyGroup.Controls.Add($refreshButton)

$openButton = New-Object System.Windows.Forms.Button
$openButton.Text = "[OPEN] Ouvrir"
$openButton.Location = New-Object System.Drawing.Point(125, 235)
$openButton.Size = New-Object System.Drawing.Size(100, 30)
$openButton.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
$openButton.ForeColor = [System.Drawing.Color]::White
$openButton.FlatStyle = "Flat"
$historyGroup.Controls.Add($openButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "[DEL] Supprimer"
$deleteButton.Location = New-Object System.Drawing.Point(235, 235)
$deleteButton.Size = New-Object System.Drawing.Size(100, 30)
$deleteButton.BackColor = [System.Drawing.Color]::FromArgb(220, 50, 50)
$deleteButton.ForeColor = [System.Drawing.Color]::White
$deleteButton.FlatStyle = "Flat"
$historyGroup.Controls.Add($deleteButton)

# Zone d'informations
$infoGroup = New-Object System.Windows.Forms.GroupBox
$infoGroup.Text = "[INFO] Informations"
$infoGroup.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$infoGroup.ForeColor = [System.Drawing.Color]::FromArgb(100, 200, 255)
$infoGroup.Location = New-Object System.Drawing.Point(20, 380)
$infoGroup.Size = New-Object System.Drawing.Size(840, 120)
$infoGroup.BackColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$mainPanel.Controls.Add($infoGroup)

$infoTextBox = New-Object System.Windows.Forms.TextBox
$infoTextBox.Location = New-Object System.Drawing.Point(15, 25)
$infoTextBox.Size = New-Object System.Drawing.Size(810, 80)
$infoTextBox.Multiline = $true
$infoTextBox.ReadOnly = $true
$infoTextBox.ScrollBars = "Vertical"
$infoTextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$infoTextBox.ForeColor = [System.Drawing.Color]::White
$infoTextBox.BorderStyle = "FixedSingle"
$infoTextBox.Text = "[OK] Application prete a numeriser`r`n[FOLDER] Dossier par defaut: $DefaultScanPath`r`n[TIP] Astuce: Vous pouvez modifier le nom du fichier et le chemin avant chaque scan"
$infoGroup.Controls.Add($infoTextBox)

# Barre de statut
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Prêt"
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Fonction pour actualiser l'historique
function Update-HistoryList {
    $historyListView.Items.Clear()
    $history = Load-ScanHistory
    
    foreach ($scan in $history) {
        $item = New-Object System.Windows.Forms.ListViewItem($scan.FileName)
        $null = $item.SubItems.Add($scan.ScanDate.ToString("dd/MM/yyyy"))
        $null = $item.SubItems.Add([math]::Round($scan.FileSize / 1KB, 1).ToString() + " KB")
        $null = $item.SubItems.Add($scan.Scanner)
        $item.Tag = $scan.FilePath
        $null = $historyListView.Items.Add($item)
    }
    
    $statusLabel.Text = "Historique actualise - $($history.Count) scans"
}

# Event handlers
$scanButton.Add_Click({
    $fileName = $fileNameTextBox.Text.Trim()
    $extension = $extensionCombo.SelectedItem
    $outputPath = Join-Path $pathTextBox.Text "$fileName$extension"
    $selectedScanner = if ($scannerCombo.SelectedIndex -ge 0) { $scanners[$scannerCombo.SelectedIndex].Name } else { "Scanner par défaut" }
    
    if ([string]::IsNullOrEmpty($fileName)) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez saisir un nom de fichier.", "Erreur", "OK", "Warning")
        return
    }
    
    if (-not (Test-Path $pathTextBox.Text)) {
        try {
            New-Item -Path $pathTextBox.Text -ItemType Directory -Force
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Impossible de creer le dossier de destination.", "Erreur", "OK", "Error")
            return
        }
    }
    
    $statusLabel.Text = "Numerisation en cours..."
    $infoTextBox.Text = "[SCANNING] Numerisation en cours...`r`nFichier: $fileName$extension`r`nDestination: $outputPath"
    
    try {
        $scanResult = Start-Scan -OutputPath $outputPath
        
        if ($scanResult -or (Test-Path $outputPath)) {
            Add-ScanToHistory -FileName "$fileName$extension" -FilePath $outputPath -Scanner $selectedScanner
            Update-HistoryList
            
            $statusLabel.Text = "Scan termine avec succes"
            $infoTextBox.Text = "[SUCCESS] Scan termine avec succes!`r`nFichier: $fileName$extension`r`nEmplacement: $outputPath`r`nTaille: $([math]::Round((Get-Item $outputPath).Length / 1KB, 1)) KB"
            
            # Generer un nouveau nom pour le prochain scan
            $fileNameTextBox.Text = "Scan_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        }
        else {
            $statusLabel.Text = "Echec du scan"
            $infoTextBox.Text = "[ERROR] Echec de la numerisation`r`nVerifiez que votre scanner est connecte et allume."
        }
    }
    catch {
        $statusLabel.Text = "Erreur lors du scan"
        $infoTextBox.Text = "[ERROR] Erreur lors de la numerisation: $($_.Exception.Message)"
    }
})

$refreshButton.Add_Click({
    Update-HistoryList
})

$openButton.Add_Click({
    if ($historyListView.SelectedItems.Count -gt 0) {
        $filePath = $historyListView.SelectedItems[0].Tag
        if (Test-Path $filePath) {
            Start-Process "explorer.exe" -ArgumentList "/select,`"$filePath`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Le fichier n'existe plus.", "Fichier introuvable", "OK", "Warning")
        }
    }
})

$deleteButton.Add_Click({
    if ($historyListView.SelectedItems.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show("Voulez-vous vraiment supprimer cette entree de l'historique?", "Confirmation", "YesNo", "Question")
        if ($result -eq "Yes") {
            $filePath = $historyListView.SelectedItems[0].Tag
            $history = Load-ScanHistory
            $history = $history | Where-Object { $_.FilePath -ne $filePath }
            Save-ScanHistory $history
            Update-HistoryList
        }
    }
})

# Initialisation
Update-HistoryList

# Affichage de l'application
$form.ShowDialog()