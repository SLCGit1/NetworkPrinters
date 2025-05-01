# NetworkPrinterGui.ps1 - Network Printer Installer Script

## 📋 Overview

`NetworkPrinterGui.ps1` is a PowerShell script that allows for interactive or silent deployment of shared printers from a central print server. It supports:

- 🖨️ Silent background installation
- 🖥️ GUI with optional auto-install
- 🔍 Filtering already installed printers
- 🔁 Forced reinstallation of printers

- Printers are compared using their UNC connection path (e.g., \\print-server\PrinterName) to determine if they are already installed. By default, only new/uninstalled printers are shown in the GUI unless the -ShowAll switch is used.

> ⚠️ **Disclaimer**: Use this script at your own risk. Test thoroughly in a staging environment before production deployment.

---

## 🔧 Arguments

| Switch            | Description                                                          |
| ----------------- | -------------------------------------------------------------------- |
| `-Silent`         | Installs all non-installed printers silently (no GUI)                |
| `-AutoInstall`    | GUI launches and auto-installs all listed printers                   |
| `-ShowAll`        | GUI includes already installed printers (unchecked by default)       |
| `-ForceReinstall` | Forces reinstallation of already installed printers                  |


---

## 🚀 Example Usage

```powershell
# Launch the script in GUI mode with default behavior (only uninstalled printers shown)
powershell.exe -File .\NetworkPrinterGui.ps1

# Run in silent mode and install all new printers
powershell.exe -File .\NetworkPrinterGui.ps1 -Silent

# Run GUI mode and auto-install printers
powershell.exe -File .\NetworkPrinterGui.ps1 -AutoInstall

# Run GUI and show all printers, including installed ones
powershell.exe -File .\NetworkPrinterGui.ps1 -ShowAll

# Force reinstall in silent mode
powershell.exe -File .\NetworkPrinterGui.ps1 -Silent -ForceReinstall

# Launch full GUI with all printers shown and reinstallation allowed
powershell.exe -File .\NetworkPrinterGui.ps1 -ShowAll -ForceReinstall
```

---

## 🧠 Code Explanation

### 🗒️ Logging
```powershell
function Log-Message {
    param ([string]$message)
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'u') - $message"
    Write-Output $message
}
```
Logs each event and step to a temporary log file and console.

### 🖨️ Getting Installed Printers
```powershell
$installedConnections = Get-Printer |
    Where-Object { $_.ComputerName -ne $null } |
    ForEach-Object { "\$($_.ComputerName)\$($_.ShareName)".Trim() }
```
Builds an array of already installed network printers using UNC paths.

### 🌐 Filtering Printers From Server
```powershell
$allPrinters = Get-Printer -ComputerName $printServer |
    Where-Object { $_.Shared -eq $true -and $_.ShareName } |
    Sort-Object -Property ShareName -Unique
```
Fetches all shared printers available on the specified print server.

### 🧩 Installing Printers
```powershell
Start-Process -FilePath "rundll32.exe" -ArgumentList @("printui.dll,PrintUIEntry", "/in", "/n$printerPath") -Wait -PassThru -NoNewWindow
```
Utilizes Windows native PrintUI to silently install printers.

### 🖱️ GUI Overview
The script uses Windows Forms for a full-featured UI with the following elements:
- `ListView`: Lists available printers with checkboxes
- `Buttons`: Select All, Unselect All, Install Selected, and Exit
- `ProgressBar` and `Label`: Visual feedback during installation

### ✅ Installation Summary
```powershell
[System.Windows.Forms.MessageBox]::Show(...)
```
Displays a summary pop-up (in the foreground) with a count of installed, skipped, and failed printers.

---

## ✅ Requirements

- ✅ Windows PowerShell 5.1+
- ✅ Administrator privileges for printer modification
- 🚫 Cannot install printers as a standard (non-admin) user — required permissions for Add-Printer, Remove-Printer, and PrintUIEntry

---

## 📄 License

MIT License *(or institutional license if required)*

---

## 👨‍💼 Maintainers

**Sarah Lawrence College – ITS Help Desk Department**

> For bugs or enhancements, please open an issue.
