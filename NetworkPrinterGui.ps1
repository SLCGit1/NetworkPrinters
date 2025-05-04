# Network Printer Installer Script
#
# Created by Jesus Ayala Sarah Lawrence College
# Modified to work with PS2EXE
#
# Supports GUI and silent operation with optional auto-install and visibility of installed printers
# Arguments:
#   -Silent         : Runs in CLI mode without GUI and installs all new printers
#   -AutoInstall    : GUI launches and automatically installs all listed printers
#   -ShowAll        : GUI includes already installed printers in the list (unchecked by default)
#   -ForceReinstall : Reinstalls even if printer is already installed
#   -PrintServer    : Optional override for the print server address (default: $PrintServer = "PrintServerFQDN")
#
# DISCLAIMER: Please test this script in a controlled environment before production use.

# Regular PowerShell parameter handling
param (
    [switch]$Silent = $false,
    [switch]$AutoInstall = $false,
    [switch]$ShowAll = $false,
    [switch]$ForceReinstall = $false,
    [string]$PrintServer = "PrintServerFQDN"
)

# PS2EXE compatibility - handle arguments when compiled as EXE
if ($MyInvocation.Line -match '\.exe') {
    # We're running as a compiled EXE - need to parse args differently
    $argList = $args
    
    # Convert positional args to named parameters
    $Silent = $argList -contains "-Silent"
    $AutoInstall = $argList -contains "-AutoInstall"
    $ShowAll = $argList -contains "-ShowAll"
    $ForceReinstall = $argList -contains "-ForceReinstall"
    
    # Handle the PrintServer parameter which requires a value
    $PrintServerIndex = [array]::IndexOf($argList, "-PrintServer")
    if ($PrintServerIndex -ge 0 -and $PrintServerIndex -lt $argList.Count - 1) {
        $PrintServer = $argList[$PrintServerIndex + 1]
    }
    # No else clause - keeps the value from the param() block
}

# Configuration
$printServerAddress = $PrintServer
$logFile = "$env:TEMP\PrinterInstallLog.txt"

# Function to log messages to file and output
function Log-Message {
    param ([string]$message)
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'u') - $message"
    Write-Output $message
}

# Log startup and parameters
Log-Message "Script started with parameters:"
Log-Message "Silent: $Silent"
Log-Message "AutoInstall: $AutoInstall"
Log-Message "ShowAll: $ShowAll"
Log-Message "ForceReinstall: $ForceReinstall"
Log-Message "PrintServer: $printServerAddress"

# Installs the provided list of printers
function Install-PrinterList {
    param ([string[]]$printerNames)

    $printServer = $printServerAddress
    # List of currently installed printer connections
    $installedConnections = Get-Printer | Where-Object { $_.ComputerName -ne $null } | ForEach-Object { "\\$($_.ComputerName)\$($_.ShareName)".Trim() }
    Log-Message "Installed Connections: $($installedConnections -join ', ')"

    $installedCount = 0
    $skippedCount = 0
    $failedCount = 0

    foreach ($printerName in $printerNames) {
        $printerPath = "\\$printServer\$printerName"
        Log-Message "CMD: rundll32.exe printui.dll,PrintUIEntry /in /n$printerPath"

        if ($installedConnections -contains $printerPath) {
            if ($ForceReinstall) {
                try {
                    Remove-Printer -Name $printerName -ErrorAction Stop
                    Log-Message "FORCE REMOVED: $printerName"
                    Start-Sleep -Seconds 1
                } catch {
                    Log-Message ("ERROR removing ${printerName}: $($_.Exception.Message)")
                    $failedCount++
                    continue
                }
            } else {
                Log-Message "SKIPPED: $printerName already installed"
                $skippedCount++
                continue
            }
        }

        try {
            $process = Start-Process -FilePath "rundll32.exe" `
                                     -ArgumentList @("printui.dll,PrintUIEntry", "/in", "/n$printerPath") `
                                     -Wait -PassThru -NoNewWindow

            if ($process.ExitCode -eq 0) {
                Log-Message "INSTALLED: $printerName"
                $installedCount++
            } else {
                throw "Installation returned exit code $($process.ExitCode)"
            }
        } catch {
            $failedCount++
            Log-Message "FAILED: ${printerName}: $($_.Exception.Message)"
        }
    }

    # Summary pop-up
    Log-Message "`nSummary:"
    Log-Message "Installed: $installedCount"
    Log-Message "Skipped : $skippedCount"
    Log-Message "Failed  : $failedCount"

    $summaryForm = New-Object System.Windows.Forms.Form
    $summaryForm.Text = "Installation Summary"
    $summaryForm.Size = New-Object System.Drawing.Size(300,150)
    $summaryForm.StartPosition = "CenterScreen"

    $summaryLabel = New-Object System.Windows.Forms.Label
    $summaryLabel.AutoSize = $true
    $summaryLabel.Location = New-Object System.Drawing.Point(20, 20)
    $summaryLabel.Text = "Installed: $installedCount`nSkipped: $skippedCount`nFailed: $failedCount"
    $summaryForm.Controls.Add($summaryLabel)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Size = New-Object System.Drawing.Size(75, 25)
    $okButton.Location = New-Object System.Drawing.Point(100, 70)
    $okButton.Add_Click({ $summaryForm.Close() })
    $summaryForm.Controls.Add($okButton)

    $summaryForm.TopMost = $true
    $summaryForm.ShowDialog()
}

# Retrieves shared printers from the print server not already installed
function Get-UniqueSharedPrinters {
    $printServer = $printServerAddress
    $allPrinters = Get-Printer -ComputerName $printServer | Where-Object { $_.Shared -eq $true -and $_.ShareName } | Sort-Object -Property ShareName -Unique
    $installedConnections = Get-Printer | Where-Object { $_.ComputerName -ne $null } | ForEach-Object { "\\$($_.ComputerName)\$($_.ShareName)".Trim() }

    $uniquePrinters = @()
    foreach ($printer in $allPrinters) {
        $printerPath = "\\$printServer\$($printer.ShareName)"
        if (-not $ShowAll -and ($installedConnections -contains $printerPath)) { continue }
        $uniquePrinters += $printer.ShareName
    }
    return $uniquePrinters
}

# Silent mode execution
if ($Silent) {
    Log-Message "Running in SILENT mode..."
    $printerList = Get-UniqueSharedPrinters
    if ($printerList.Count -eq 0) {
        Log-Message "No new printers to install."
        exit 0
    }
    Install-PrinterList -printerNames $printerList
    exit 0
}

# GUI Setup
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Network Printer Installer"
$form.Size = New-Object System.Drawing.Size(600, 560)
$form.StartPosition = "CenterScreen"
$form.Topmost = $true

# Instruction label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(580, 40)
$label.Text = if ($ShowAll) { "Select printers to install (all shown)." } else { "Select printers to install (installed printers are hidden)." }
$form.Controls.Add($label)

# Add server info label (for debugging)
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(10, 450)
$serverLabel.Size = New-Object System.Drawing.Size(580, 20)
$serverLabel.Text = "Print Server: $printServerAddress"
$form.Controls.Add($serverLabel)

# List view for available printers
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 50)
$listView.Size = New-Object System.Drawing.Size(560, 300)
$listView.View = 'Details'
$listView.CheckBoxes = $true
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Columns.Add("Printer Share Name", 250)
$listView.Columns.Add("Location", 290)
$form.Controls.Add($listView)

# Action buttons
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Location = New-Object System.Drawing.Point(10, 360)
$selectAllButton.Size = New-Object System.Drawing.Size(100, 30)
$selectAllButton.Text = "Select All"
$form.Controls.Add($selectAllButton)

$unselectAllButton = New-Object System.Windows.Forms.Button
$unselectAllButton.Location = New-Object System.Drawing.Point(120, 360)
$unselectAllButton.Size = New-Object System.Drawing.Size(100, 30)
$unselectAllButton.Text = "Unselect All"
$form.Controls.Add($unselectAllButton)

$installButton = New-Object System.Windows.Forms.Button
$installButton.Location = New-Object System.Drawing.Point(230, 360)
$installButton.Size = New-Object System.Drawing.Size(150, 30)
$installButton.Text = "Install Selected"
$form.Controls.Add($installButton)

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(390, 360)
$exitButton.Size = New-Object System.Drawing.Size(80, 30)
$exitButton.Text = "Exit"
$form.Controls.Add($exitButton)

# Progress and status
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 400)
$progressBar.Size = New-Object System.Drawing.Size(560, 20)
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 430)
$statusLabel.Size = New-Object System.Drawing.Size(560, 20)
$form.Controls.Add($statusLabel)

# Dictionary to hold printer objects
$printerMap = @{}

# Populates printer list into the GUI
function Populate-PrinterList {
    $allPrinters = Get-Printer -ComputerName $printServerAddress | Where-Object { $_.Shared -eq $true -and $_.ShareName } | Sort-Object -Property ShareName -Unique
    $installedConnections = Get-Printer | Where-Object { $_.ComputerName -ne $null } | ForEach-Object { "\\$($_.ComputerName)\$($_.ShareName)".Trim() }

    foreach ($printer in $allPrinters) {
        $shareName = $printer.ShareName
        $printerPath = "\\$printServerAddress\$shareName"
        if (-not $ShowAll -and ($installedConnections -contains $printerPath)) { continue }

        $location = $printer.Location
        $item = New-Object System.Windows.Forms.ListViewItem($shareName)
        $item.SubItems.Add($location)
        $item.Checked = -not ($installedConnections -contains $printerPath)
        $printerMap[$shareName] = $printer
        $listView.Items.Add($item)
    }
}

# Handles the selected printers and starts installation
function Install-Printers {
    param ([System.Windows.Forms.ListViewItem[]]$selectedItems)
    $printerNames = $selectedItems | ForEach-Object { $_.Text }
    Install-PrinterList -printerNames $printerNames
}

# Button actions
$selectAllButton.Add_Click({ foreach ($item in $listView.Items) { $item.Checked = $true } })
$unselectAllButton.Add_Click({ foreach ($item in $listView.Items) { $item.Checked = $false } })

$installButton.Add_Click({
    $selectedItems = @($listView.Items | Where-Object { $_.Checked })
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one printer.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $installButton.Enabled = $false
    $selectAllButton.Enabled = $false
    $unselectAllButton.Enabled = $false
    $exitButton.Enabled = $false
    $label.Text = "Installing printers... Please wait."
    $form.Refresh()

    Install-Printers -selectedItems $selectedItems

    $installButton.Enabled = $true
    $selectAllButton.Enabled = $true
    $unselectAllButton.Enabled = $true
    $exitButton.Enabled = $true
    $label.Text = "Installation complete. You can select more printers or close this window."
})

$exitButton.Add_Click({ $form.Close() })

# Initialize the form
Populate-PrinterList

if ($AutoInstall -and $listView.Items.Count -gt 0) {
    foreach ($item in $listView.Items) { $item.Checked = $true }
    Install-Printers -selectedItems @($listView.Items)
    $form.Close()
}

$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
