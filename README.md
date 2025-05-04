# NetworkPrinterGui.ps1 - Network Printer Installer Script

## 📋 Overview

`NetworkPrinterGui.ps1` is a PowerShell script that allows for interactive or silent deployment of shared printers from a central print server. The script will check what printer the user has access to and also check if the printer has been installed already. It supports:

- 🖨️ Silent background installation
- 🖥️ GUI with optional auto-install
- 🔍 Filtering already installed printers
- 🔁 Forced reinstallation of printers
- 📦 Conversion to executable with PS2EXE

Printers are compared using their UNC connection path (e.g., `\\print-server\PrinterName`) to determine if they are already installed. By default, only new/uninstalled printers are shown in the GUI unless the `-ShowAll` switch is used.

> ⚠️ **Disclaimer**: Use this script at your own risk. Test thoroughly in a staging environment before production deployment. Depending on how many printers you have, it might take some time to show.

---

![NetworkPrinter](https://github.com/user-attachments/assets/6aa2ac6d-f9f6-484f-810e-ea4b2aef0d44)

## 🔧 Arguments

>> ⚠️ **Important**: Before using this script, you need to update the default printer server address by modifying line 17 in the script:
>> ```powershell
>> [string]$PrintServer = "PrintServerFQDN"
>> ```
>> Replace "PrintServerFQDN" with your actual print server's fully qualified domain name.

---

| Switch            | Description                                                          |
| ----------------- | -------------------------------------------------------------------- |
| `-Silent`         | Installs all non-installed printers silently (no GUI)                |
| `-AutoInstall`    | GUI launches and auto-installs all listed printers                   |
| `-ShowAll`        | GUI includes already installed printers (unchecked by default)       |
| `-ForceReinstall` | Forces reinstallation of already installed printers                  |
| `-PrintServer`    | Optional override for the print server address                       |

---

## 🚀 Example Usage (PowerShell Script)

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

# Use a custom print server address
powershell.exe -File .\NetworkPrinterGui.ps1 -PrintServer your-server.domain.com
```

## 📦 Executable Version (PS2EXE)

The script includes improved support for conversion to an executable file using PS2EXE. This allows for easier deployment and usage without requiring users to manually run PowerShell commands.

### Converting to EXE

To convert the script to an executable:

1. Install the PS2EXE module if you haven't already:
   ```powershell
   Install-Module -Name PS2EXE
   ```

2. Convert the script:
   ```powershell
   Invoke-ps2exe -InputFile .\NetworkPrinterGui.ps1 -OutputFile .\NetworkPrinterGui.exe
   ```

### 🚀 Example Usage (Executable)

#### Running from PowerShell
When using the executable version in PowerShell, you must prefix the executable with `.\` when running from the current directory:

```powershell
# Launch the EXE in GUI mode with default behavior
.\NetworkPrinterGui.exe

# Run in silent mode and install all new printers
.\NetworkPrinterGui.exe -Silent

# Run GUI mode and auto-install printers
.\NetworkPrinterGui.exe -AutoInstall

# Run GUI and show all printers, including installed ones
.\NetworkPrinterGui.exe -ShowAll

# Force reinstall in silent mode
.\NetworkPrinterGui.exe -Silent -ForceReinstall

# Specify a custom print server
.\NetworkPrinterGui.exe -PrintServer your-server.domain.com

# Combine multiple parameters
.\NetworkPrinterGui.exe -Silent -ForceReinstall -PrintServer your-server.domain.com
```

#### Running from Command Prompt (CMD)
When using the executable version from CMD, you can run it directly without the `.\` prefix:

```cmd
:: Launch the EXE in GUI mode with default behavior
NetworkPrinterGui.exe

:: Run in silent mode and install all new printers
NetworkPrinterGui.exe -Silent

:: Run GUI mode and auto-install printers
NetworkPrinterGui.exe -AutoInstall

:: Run GUI and show all printers, including installed ones
NetworkPrinterGui.exe -ShowAll

:: Force reinstall in silent mode
NetworkPrinterGui.exe -Silent -ForceReinstall

:: Specify a custom print server
NetworkPrinterGui.exe -PrintServer your-server.domain.com

:: Combine multiple parameters
NetworkPrinterGui.exe -Silent -ForceReinstall -PrintServer your-server.domain.com
```

You can also run the executable directly from File Explorer by double-clicking it, which will launch the GUI with default settings.

---

## 🔄 Deployment as Logon Script

To deploy this script as a logon-time printer installer for non-admin users:

### 🧰 Setup Requirements
- Must be deployed in a way that runs the script under the **user context**.
- Non-admin users can **enumerate printers**, but installing/removing printers **requires elevation**.
- To support installation without user prompts, **Task Scheduler** can be used.

### 🎯 Recommended Task Scheduler Setup
1. Open Task Scheduler (taskschd.msc)
2. Create a new task:
   - **Name**: `Install Network Printers`
   - **Run only when user is logged on**
   - **Run with highest privileges** (if elevation is permitted)
3. Triggers:
   - **At log on** → Specific user (or all users)
4. Actions:
   - **Start a program**:
     - For PowerShell script:
       - Program/script: `powershell.exe`
       - Add arguments: `-ExecutionPolicy Bypass -File "C:\Path\To\NetworkPrinterGui.ps1" -Silent`
     - For executable:
       - Program/script: `C:\Path\To\NetworkPrinterGui.exe`
       - Add arguments: `-Silent`
5. Conditions/Settings:
   - Disable conditions like "start only on AC power" unless needed.
   - Enable "Run task as soon as possible after a scheduled start is missed."

📌 If running without elevation, the script will list printers but **will not install them** unless run with appropriate permissions.

---

## 🧠 Code Explanation

### 🔄 Parameter Handling Improvements
The script uses a dual approach to parameter handling:

```powershell
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
```

This approach ensures proper parameter handling in both PowerShell script and executable modes:
- When running as a normal PowerShell script, it uses PowerShell's native parameter handling
- When running as a compiled executable, it detects this mode and manually parses the command line arguments
- The default print server value is defined in a single place (the param block), ensuring consistency between modes

### 📊 Enhanced GUI with Server Information Display
```powershell
# Add server info label (for debugging)
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(10, 450)
$serverLabel.Size = New-Object System.Drawing.Size(580, 20)
$serverLabel.Text = "Print Server: $printServerAddress"
$form.Controls.Add($serverLabel)
```

The GUI now includes a visible label showing which print server is being used, making it easier to verify the connection settings, especially when using the `-PrintServer` parameter to override the default.

### 🗒️ Logging
```powershell
function Log-Message {
    param ([string]$message)
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'u') - $message"
    Write-Output $message
}
```
The script logs all actions to both a temporary log file (`%TEMP%\PrinterInstallLog.txt`) and the console output, making it easier to troubleshoot issues.

### 🖨️ Getting Installed Printers
```powershell
$installedConnections = Get-Printer |
    Where-Object { $_.ComputerName -ne $null } |
    ForEach-Object { "\\$($_.ComputerName)\$($_.ShareName)".Trim() }
```
The script builds an array of already installed network printers using UNC paths to avoid duplicate installations.

### 🌐 Filtering Printers From Server
```powershell
$allPrinters = Get-Printer -ComputerName $printServer |
    Where-Object { $_.Shared -eq $true -and $_.ShareName } |
    Sort-Object -Property ShareName -Unique
```
The script fetches all shared printers available on the specified print server, then filters out printers that are already installed (unless `-ShowAll` is used).

### 🔄 PS2EXE Parameter Compatibility
The improved parameter handling ensures consistent behavior between PowerShell script and executable modes, especially for the PrintServer parameter. Key improvements include:

1. Using a single default print server value defined in the param block
2. Proper detection of execution context (PowerShell script vs. EXE)
3. No redundant default value in the EXE handling section - it uses the value from the param block if no override is provided
4. Added server information label to the GUI for better visibility and debugging

---

## ✅ Requirements

- ✅ Windows PowerShell 5.1+
- ✅ Administrator privileges for printer modification
- 🚫 Cannot install printers as a standard (non-admin) user — required permissions for Add-Printer, Remove-Printer, and PrintUIEntry
- ✅ If using Task Scheduler, configure to run at logon with appropriate permissions
- ✅ Update the default print server address in the param block: `[string]$PrintServer = "YourActualPrintServer"`
- ✅ For PS2EXE conversion: PS2EXE module (`Install-Module -Name PS2EXE`)

---

## 📝 Important Notes for Customization

To adapt this script to your environment:

1. **Update the Default Print Server**:
   - Edit line 17 in the script:
   ```powershell
   [string]$PrintServer = "PrintServerFQDN"
   ```
   - Replace "PrintServerFQDN" with your actual print server address (e.g., "print-server.domain.com")

2. **Custom Deployment Considerations**:
   - When compiling with PS2EXE, the default parameters from the param() block will be used
   - Test the executable with and without parameters to ensure correct behavior
   - The log file at `%TEMP%\PrinterInstallLog.txt` can help troubleshoot parameter issues

3. **Parameter Handling**:
   - The script now uses a consistent approach to parameter handling
   - The default PrintServer value is maintained across both script and executable modes
   - The GUI displays the active PrintServer value for verification

---

## 📄 License

MIT License *(or institutional license if required)*

---

## 👨‍💼 Maintainers

**Sarah Lawrence College – ITS Help Desk Department - Jesus Ayala**

> For bugs or enhancements, please open an issue.
