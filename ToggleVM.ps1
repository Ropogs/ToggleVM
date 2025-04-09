[CmdletBinding()]
param(
    [switch]$StartVM,
    [switch]$ShutdownVM,
    [switch]$KillVM,
    [switch]$ConnectParsec,
    [switch]$OpenMenu
)

# ----------------------------------------------------------------------------------
#                           Session & Logging Setup
# ----------------------------------------------------------------------------------

$Global:Session = [PSCustomObject]@{
    # Notice: no peerid here; we only set defaults for other keys
    DefaultConfig = @"
vmName=GPUPV
parsecExe=C:\Program Files\Parsec\parsecd.exe
parsecCloseBehavior=prompt
safetyConfirmation=true
"@
    ConfigFile         = Join-Path -Path $PSScriptRoot -ChildPath "GPUPV.config"
    VMName             = $null
    ParsecPeerId       = $null  # We'll prompt for it if missing
    ParsecExe          = $null
    VMConnectProcess   = $null
    TempConnectProcess = $null
    LogFile            = Join-Path -Path $PSScriptRoot -ChildPath "GPUPV.log"
    ParsecCloseBehavior= "prompt"
    SafetyConfirmation = $true
}

function Log-Info {
    param([string]$Message)
    Write-Host "[INFO ] $Message" -ForegroundColor White
    Out-File -FilePath $Global:Session.LogFile -Append -InputObject ("[INFO ] " + $Message)
}
function Log-Warning {
    param([string]$Message)
    Write-Host "[WARN ] $Message" -ForegroundColor Yellow
    Out-File -FilePath $Global:Session.LogFile -Append -InputObject ("[WARN ] " + $Message)
}
function Log-Error {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
    Out-File -FilePath $Global:Session.LogFile -Append -InputObject ("[ERROR] " + $Message)
}
function Log-Success {
    param([string]$Message)
    Write-Host "[ OK  ] $Message" -ForegroundColor Green
    Out-File -FilePath $Global:Session.LogFile -Append -InputObject ("[ OK  ] " + $Message)
}

function Safe-StopProcess {
    param(
        [Parameter(Mandatory)][System.Diagnostics.Process]$Process,
        [switch]$Force
    )
    try {
        if ($Process -and -not $Process.HasExited) {
            if ($Force) {
                Stop-Process -Id $Process.Id -Force -ErrorAction Stop
            } else {
                Stop-Process -Id $Process.Id -ErrorAction Stop
            }
            Log-Info "Stopped process: $($Process.Name) (PID: $($Process.Id))"
        }
    }
    catch {
        Log-Warning "Could not stop process '$($($Process?.Name))': $($_.Exception.Message)"
    }
}

# ----------------------------------------------------------------------------------
#                           Config File Setup & Validation
# ----------------------------------------------------------------------------------
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Validate-Config {
    param([string]$ConfigContent)
    # We only require these 4 keys to exist for the script to run properly.
    # `peerid` is handled separately if missing
    $keysNeeded = @("vmName", "parsecExe", "parsecCloseBehavior", "safetyConfirmation")

    $lines = $ConfigContent -split "`n"
    $configDict = @{}
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -match "^(.*)=(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $configDict[$key] = $value
        }
    }

    foreach ($key in $keysNeeded) {
        if (-not $configDict.ContainsKey($key) -or [string]::IsNullOrEmpty($configDict[$key])) {
            return $false
        }
    }
    return $true
}

function Reload-ConfigAndValidate {
    if (-not (Test-Path $Global:Session.ConfigFile)) {
        # If config doesn't exist, create from default
        $Global:Session.DefaultConfig | Out-File -FilePath $Global:Session.ConfigFile -Encoding utf8
        Log-Success "Config file created with default values (no peerid)."
    } else {
        $configContent = Get-Content $Global:Session.ConfigFile -Raw
        # Validate only the required 4 keys
        if (-not (Validate-Config -ConfigContent $configContent)) {
            Log-Error "Config file is missing required keys or incorrectly configured (vmName, parsecExe, parsecCloseBehavior, safetyConfirmation)."
            $reset = Read-Host "Do you want to reset the config to default values? (y/n)"
            if ($reset -eq "y") {
                $Global:Session.DefaultConfig | Out-File -FilePath $Global:Session.ConfigFile -Encoding utf8 -Force
                Log-Success "Config file reset to default values (no peerid)."
            }
        }
    }

    # Now parse the config
    $lines = Get-Content $Global:Session.ConfigFile
    foreach ($line in $lines) {
        if ($line -match "^(.*)=(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            switch ($key) {
                "peerid"               { $Global:Session.ParsecPeerId        = $value }
                "vmName"               { $Global:Session.VMName             = $value }
                "parsecExe"            { $Global:Session.ParsecExe          = $value }
                "parsecCloseBehavior"  { $Global:Session.ParsecCloseBehavior= $value }
                "safetyConfirmation"   {
                    $Global:Session.SafetyConfirmation = $value -eq 'true'
                }
            }
        }
    }
	
	function Set-ParsecPeerId {
    param([string]$newPeerId)
    if ([string]::IsNullOrWhiteSpace($newPeerId)) {
        Log-Error "Peer ID cannot be blank."
        return
    }
    try {
        $lines = Get-Content -Path $Global:Session.ConfigFile -ErrorAction Stop
        $found = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^\s*peerid\s*=") {
                $lines[$i] = "peerid=$newPeerId"
                $found = $true
                break
            }
        }
        if (-not $found) {
            $lines += "peerid=$newPeerId"
        }
        Set-Content -Path $Global:Session.ConfigFile -Value $lines -Force
        $Global:Session.ParsecPeerId = $newPeerId
        Log-Success "Parsec peer id updated to: $newPeerId."
    }
    catch {
        Log-Error "Error updating config file: $_"
    }
}

    # If peerid is missing or empty, prompt user to provide it
    if (-not $Global:Session.ParsecPeerId) {
        $peer = Read-Host "No Parsec peerid found! Please enter your Parsec peerid now"
        if ([string]::IsNullOrWhiteSpace($peer)) {
            Log-Error "No peerid entered! The script cannot proceed without a peerid."
            exit 1
        }
        # set & persist
        Set-ParsecPeerId -newPeerId $peer
    }
}

Reload-ConfigAndValidate

# ----------------------------------------------------------------------------------
#          Import functions to detect if the console window is active
# ----------------------------------------------------------------------------------
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
}
"@

function IsConsoleActive {
    $consoleHWND = [Win32]::GetConsoleWindow()
    $foregroundHWND = [Win32]::GetForegroundWindow()
    return ($consoleHWND -eq $foregroundHWND)
}

# ----------------------------------------------------------------------------------
#                             Common "Confirm" Helper
# ----------------------------------------------------------------------------------
function Confirm-Action {
    param(
        [string]$Message
    )

    if (-not $Global:Session.SafetyConfirmation) {
        # If safety is off, auto-confirm
        return $true
    }
    $answer = Read-Host "$Message (y/n)"
    return ($answer -eq 'y')
}

# ----------------------------------------------------------------------------------
#                             Parsec-Related Functions
# ----------------------------------------------------------------------------------
function ParsecKill {
    $killedAny = $false
    $allProcesses = Get-Process -ErrorAction SilentlyContinue

    foreach ($proc in $allProcesses) {
        try {
            $fileInfo = $proc.MainModule.FileVersionInfo
            if (($fileInfo.ProductName -match "(?i)Parsec") -or ($fileInfo.FileDescription -match "(?i)Parsec")) {
                Safe-StopProcess -Process $proc -Force
                $killedAny = $true
            }
        }
        catch {}
    }
    if ($killedAny) {
        Start-Sleep -Seconds 1
        Log-Warning "All processes with ProductName/FileDescription 'Parsec' have been terminated."
    } else {
        Log-Success "No Parsec processes found."
    }
}

function Get-ParsecStatus {
    $parsecProcesses = Get-Process -Name "parsec", "parsecd" -ErrorAction SilentlyContinue
    return ($parsecProcesses -ne $null -and $parsecProcesses.Count -gt 0)
}

function OpenParsecWindow {
    if (Test-Path $Global:Session.ParsecExe) {
        $wshShell = New-Object -ComObject WScript.Shell
        $existing = Get-Process -Name "parsec", "parsecd" -ErrorAction SilentlyContinue
        if ($existing) {
            if ($wshShell.AppActivate("Parsec")) {
                Log-Warning "Parsec is already open; brought to the front."
            } else {
                Log-Warning "Parsec is already open (could not bring to front)."
            }
        } else {
            Log-Info "Opening Parsec..."
            Start-Process $Global:Session.ParsecExe
        }
    } else {
        Log-Error "Parsec executable not found at $($Global:Session.ParsecExe)"
    }
}

function Handle-ParsecToggle {
    if (Get-ParsecStatus) {
        ParsecKill
    } else {
        OpenParsecWindow
    }
}

# ----------------------------------------------------------------------------------
#                           Hyper-V Manager & VM Functions
# ----------------------------------------------------------------------------------
function Open-HyperVManager {
    $wshShell = New-Object -ComObject WScript.Shell
    if ($wshShell.AppActivate("Hyper-V Manager")) {
        Log-Info "Hyper-V Manager is already running; bringing it to front."
    } else {
        Log-Info "Opening Hyper-V Manager..."
        Start-Process "virtmgmt.msc"
    }
}

function FactoryReset-VM {
    $vm = Get-VM -Name $Global:Session.VMName
    if ($vm.State -ne "Off") {
        Log-Error "VM must be off to perform a factory reset."
        return
    }
    if (-not (Confirm-Action "Really factory reset $($Global:Session.VMName)?")) {
        Log-Warning "Factory reset cancelled."
        return
    }

    $checkpoint = Get-VMCheckpoint -VMName $Global:Session.VMName -Name "factory reset" -ErrorAction SilentlyContinue
    if ($null -eq $checkpoint) {
        Log-Error "No 'factory reset' checkpoint found. Cannot perform factory reset."
        return
    }
    Log-Info "Applying 'factory reset' checkpoint to $($Global:Session.VMName)..."
    try {
        Restore-VMCheckpoint -VMName $Global:Session.VMName -Name "factory reset" -Confirm:$false
        Log-Success "Factory reset applied successfully."
    }
    catch {
        Log-Error "Error applying factory reset: $_"
    }
}

function VMKill {
    $vm = Get-VM -Name $Global:Session.VMName
    if ($vm.State -ne "Running") {
        Log-Warning "VM is not running."
        return
    }
    if (-not (Confirm-Action "Kill VM $($Global:Session.VMName)? (possible data loss)")) {
        Log-Warning "VM termination cancelled."
        return
    }

    if ($Global:Session.TempConnectProcess) {
        Safe-StopProcess -Process $Global:Session.TempConnectProcess -Force
        $Global:Session.TempConnectProcess = $null
        Log-Warning "Temporary Hyper-V connection closed."
    }
    if ($Global:Session.VMConnectProcess -and -not $Global:Session.VMConnectProcess.HasExited) {
        Safe-StopProcess -Process $Global:Session.VMConnectProcess -Force
        $Global:Session.VMConnectProcess = $null
        Log-Warning "Permanent Hyper-V connection closed."
    }
    try {
        Stop-VM -Name $Global:Session.VMName -TurnOff -Confirm:$false
        Log-Error "VM $($Global:Session.VMName) has been terminated (killed)."
    }
    catch {
        Log-Error "Failed to terminate VM: $_"
    }
}

function Connect-VMParsec {
    $vm = Get-VM -Name $Global:Session.VMName
    if ($vm.State -ne "Running") {
        Log-Error "VM $($Global:Session.VMName) is not running. Cannot connect via Parsec."
        return
    }
    Log-Info "Connecting to $($Global:Session.VMName) via Parsec using peer_id=$($Global:Session.ParsecPeerId)..."
    if (Test-Path $Global:Session.ParsecExe) {
        Start-Process $Global:Session.ParsecExe -ArgumentList "peer_id=$($Global:Session.ParsecPeerId)"
    } else {
        Log-Error "Parsec executable not found at $($Global:Session.ParsecExe)"
    }
}

function Disconnect-HyperVAndParsec {
    if ($Global:Session.TempConnectProcess) {
        Safe-StopProcess -Process $Global:Session.TempConnectProcess -Force
        $Global:Session.TempConnectProcess = $null
    }
    if ($Global:Session.VMConnectProcess -and -not $Global:Session.VMConnectProcess.HasExited) {
        Safe-StopProcess -Process $Global:Session.VMConnectProcess -Force
        $Global:Session.VMConnectProcess = $null
    }
    Log-Warning "Hyper-V connection(s) closed. Launching Parsec..."
    Connect-VMParsec
}

function Get-HyperVStatus {
    $tempActive = $Global:Session.TempConnectProcess -and (-not $Global:Session.TempConnectProcess.HasExited)
    $permActive = $Global:Session.VMConnectProcess -and (-not $Global:Session.VMConnectProcess.HasExited)
    if ($tempActive -and $permActive) { return "Both" }
    elseif ($permActive) { return "Permanent" }
    elseif ($tempActive) { return "Temporary" }
    else { return "None" }
}

function Start-VMCustom {
    Log-Info "Starting $($Global:Session.VMName)..."
    Start-VM -Name $Global:Session.VMName
    Log-Info "Opening temporary Hyper-V connection (25s wait)..."
    Connect-VMTemporary
}

function Start-VMWithParsec {
    Log-Info "Starting $($Global:Session.VMName) with temporary Hyper-V connection then Parsec..."
    Start-VMCustom
    Connect-VMParsec
}

function Start-VMWithPermanent {
    Log-Info "Starting $($Global:Session.VMName) with permanent Hyper-V connection..."
    Start-VM -Name $Global:Session.VMName
    Log-Info "Opening permanent Hyper-V connection..."
    Connect-VMPermanent
}

function Start-VMNoConnection {
    Log-Info "Starting $($Global:Session.VMName) without establishing a new connection..."
    Start-VM -Name $Global:Session.VMName
}

function Restart-VMCustom {
    Shutdown-VMCustom -CloseConnections:$false
    ParsecKill
    if ($Global:Session.TempConnectProcess) {
        Safe-StopProcess -Process $Global:Session.TempConnectProcess -Force
        $Global:Session.TempConnectProcess = $null
    }
    if ($Global:Session.VMConnectProcess -and -not $Global:Session.VMConnectProcess.HasExited) {
        Safe-StopProcess -Process $Global:Session.VMConnectProcess -Force
        $Global:Session.VMConnectProcess = $null
    }

    Log-Info "Connect via Parsec after restart? (y/n)"
    $ans = Get-MenuChoice
    if ($ans -eq 'y') {
        Start-VMWithParsec
    } else {
        Start-VMNoConnection
    }
}

function Connect-VMTemporary {
    Log-Info "Starting temporary Hyper-V connection (auto-closes in 25s)..."
    
    $processParams = @{
        FilePath     = "C:\Windows\System32\vmconnect.exe"
        ArgumentList = @("localhost", $Global:Session.VMName)
        PassThru     = $true
        WindowStyle  = "Normal"
    }
    $Global:Session.TempConnectProcess = Start-Process @processParams
    
    $startTime = Get-Date
    Log-Info "Press 'y' to close connection early."
    while (((Get-Date) - $startTime).TotalSeconds -lt 25) {
        Start-Sleep -Milliseconds 200
        if ([System.Console]::KeyAvailable) {
            $key = [System.Console]::ReadKey($true)
            if ($key.KeyChar -eq 'y') {
                Log-Warning "Closing temporary Hyper-V connection per user input."
                Safe-StopProcess -Process $Global:Session.TempConnectProcess -Force
                $Global:Session.TempConnectProcess = $null
                return
            }
        }
        if ($Global:Session.TempConnectProcess.HasExited) {
            $Global:Session.TempConnectProcess = $null
            return
        }
    }
    if ($Global:Session.TempConnectProcess -and -not $Global:Session.TempConnectProcess.HasExited) {
        Log-Warning "Temporary Hyper-V connection time expired. Closing connection."
        Safe-StopProcess -Process $Global:Session.TempConnectProcess -Force
        $Global:Session.TempConnectProcess = $null
    }
}

function Connect-VMPermanent {
    if ($Global:Session.VMConnectProcess -and -not $Global:Session.VMConnectProcess.HasExited) {
        Log-Warning "A permanent Hyper-V connection is already active."
    } else {
        Log-Info "Starting a permanent Hyper-V connection to $($Global:Session.VMName)..."
        $processParams = @{
            FilePath     = "C:\Windows\System32\vmconnect.exe"
            ArgumentList = @("localhost", $Global:Session.VMName)
            PassThru     = $true
            WindowStyle  = "Normal"
        }
        $Global:Session.VMConnectProcess = Start-Process @processParams
    }
}

function Disconnect-VM {
    if ($Global:Session.VMConnectProcess -and -not $Global:Session.VMConnectProcess.HasExited) {
        Safe-StopProcess -Process $Global:Session.VMConnectProcess -Force
        Log-Success "Permanent Hyper-V connection disconnected."
        $Global:Session.VMConnectProcess = $null
    } else {
        Log-Warning "No permanent Hyper-V connection active."
    }
}

# ----------------------------------------------------------------------------------
#                                Settings
# ----------------------------------------------------------------------------------

function Set-ParsecCloseBehavior {
    param([string]$newBehavior)
    try {
        $lines = Get-Content -Path $Global:Session.ConfigFile -ErrorAction Stop
        $found = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^\s*parsecCloseBehavior\s*=") {
                $lines[$i] = "parsecCloseBehavior=$newBehavior"
                $found = $true
                break
            }
        }
        if (-not $found) {
            $lines += "parsecCloseBehavior=$newBehavior"
        }
        Set-Content -Path $Global:Session.ConfigFile -Value $lines -Force
        $Global:Session.ParsecCloseBehavior = $newBehavior
        Log-Success "parsecCloseBehavior updated to: $newBehavior."
    }
    catch {
        Log-Error "Error updating config file: $_"
    }
}

function Set-SafetyConfirmation {
    param([bool]$enabled)
    try {
        $lines = Get-Content -Path $Global:Session.ConfigFile -ErrorAction Stop
        $found = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match "^\s*safetyConfirmation\s*=") {
                $lines[$i] = "safetyConfirmation=$($enabled.ToString().ToLower())"
                $found = $true
                break
            }
        }
        if (-not $found) {
            $lines += "safetyConfirmation=$($enabled.ToString().ToLower())"
        }
        Set-Content -Path $Global:Session.ConfigFile -Value $lines -Force
        $Global:Session.SafetyConfirmation = $enabled
        Log-Success "safetyConfirmation updated to: $enabled."
    }
    catch {
        Log-Error "Error updating config file: $_"
    }
}

# ----------------------------------------------------------------------------------
#                                Terminate Everything
# ----------------------------------------------------------------------------------
function Terminate-Everything {
    Log-Error "WARNING: This will terminate all connections and shut down all running VMs."
    if (-not (Confirm-Action "Are you sure you want to terminate everything?")) {
        return
    }

    ParsecKill
    
    $hvManagers = Get-Process | Where-Object { $_.MainWindowTitle -like "*Hyper-V Manager*" }
    if ($hvManagers) {
        $hvManagers | ForEach-Object {
            Safe-StopProcess -Process $_ -Force
        }
        Log-Warning "Hyper-V Manager closed."
    }

    $runningVMs = Get-VM | Where-Object { $_.State -eq "Running" }
    foreach ($vm in $runningVMs) {
        if ($vm.Name -eq $Global:Session.VMName) {
            if ($Global:Session.SafetyConfirmation) {
                $choiceKill = Read-Host "For VM '$($vm.Name)', kill (y) or shutdown gracefully (n)? (y/n)"
                if ($choiceKill -eq 'y') {
                    VMKill
                } else {
                    Log-Info "Shutting down VM: $($vm.Name) gracefully..."
                    Stop-VM -Name $vm.Name -Confirm:$false
                    Start-Sleep -Seconds 3
                    $currentState = (Get-VM -Name $vm.Name).State
                    if ($currentState -ne "Off") {
                        Log-Error "Force turning off VM: $($vm.Name)"
                        Stop-VM -Name $vm.Name -TurnOff -Confirm:$false
                    }
                }
            }
            else {
                # No safety => just kill
                VMKill
            }
        } else {
            Log-Info "Shutting down VM: $($vm.Name) gracefully..."
            Stop-VM -Name $vm.Name -Confirm:$false
            Start-Sleep -Seconds 3
            $currentState = (Get-VM -Name $vm.Name).State
            if ($currentState -ne "Off") {
                Log-Error "Force turning off VM: $($vm.Name)"
                Stop-VM -Name $vm.Name -TurnOff -Confirm:$false
            }
        }
    }

    if ($Global:Session.TempConnectProcess -and -not $Global:Session.TempConnectProcess.HasExited) {
        Safe-StopProcess -Process $Global:Session.TempConnectProcess -Force
        $Global:Session.TempConnectProcess = $null
        Log-Warning "Temporary Hyper-V connection closed."
    }
    if ($Global:Session.VMConnectProcess -and -not $Global:Session.VMConnectProcess.HasExited) {
        Safe-StopProcess -Process $Global:Session.VMConnectProcess -Force
        $Global:Session.VMConnectProcess = $null
        Log-Warning "Permanent Hyper-V connection closed."
    }

    Stop-Process -Name "vmconnect" -Force -ErrorAction SilentlyContinue
    
    Log-Error "Terminated everything."
    $final = Read-Host "Press 'y' to fully exit the script, or 'n' to stay in the menu. (y/n)"
    if ($final -eq 'y') {
        Log-Error "Exiting script now."
        exit 0
    } else {
        Log-Warning "Returning to menu..."
    }
}

# ----------------------------------------------------------------------------------
#                               Menu / UI Functions
# ----------------------------------------------------------------------------------
function Show-Header {
    param([string]$vmState)
    Write-Host ""
    Write-Host "   ~~~ VM STATUS ~~~" -ForegroundColor Yellow
    Write-Host "      Name  : $($Global:Session.VMName)" -ForegroundColor Green
    Write-Host "      State : $vmState" -ForegroundColor Red
    Write-Host "   ~~~~~~~~~~~~~~~~~" -ForegroundColor Yellow
    Write-Host ""
}

function Show-OffMenu {
    Write-Host ""
    Write-Host "   ░░░ VM CONTROL (Off) ░░░" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   [1] " -NoNewline; Write-Host "Start VM & Connect via Parsec" -ForegroundColor Green
    Write-Host "   [2] " -NoNewline; Write-Host "Start VM with Permanent Hyper-V Connection" -ForegroundColor DarkYellow
    Write-Host "   [h] " -NoNewline; Write-Host "Open Hyper-V Manager" -ForegroundColor DarkYellow
    if (Get-ParsecStatus) {
        Write-Host "   [k] " -NoNewline; Write-Host "Close Parsec" -ForegroundColor Red
    } else {
        Write-Host "   [k] " -NoNewline; Write-Host "Open Parsec" -ForegroundColor Green
    }
    Write-Host "   [r] " -NoNewline; Write-Host "Refresh Menu" -ForegroundColor White
    Write-Host "   [s] " -NoNewline; Write-Host "Settings" -ForegroundColor DarkYellow
    Write-Host "   [t] " -NoNewline; Write-Host "Terminate Everything" -ForegroundColor Red
    Write-Host "   [0/q] " -NoNewline; Write-Host "Quit" -ForegroundColor Red
    Write-Host ""
}

function Show-RunningMenu {
    $hvStatus = Get-HyperVStatus
    Write-Host ""
    Write-Host "   ░░░ VM CONTROL (Running) ░░░" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   [1] " -NoNewline; Write-Host "Shutdown VM" -ForegroundColor Green
    Write-Host "   [2] " -NoNewline; Write-Host "Restart VM" -ForegroundColor Red
    if ($hvStatus -eq "None") {
        Write-Host "   [3] " -NoNewline; Write-Host "Temporary Connect (25s Hyper-V)" -ForegroundColor DarkYellow
        Write-Host "   [4] " -NoNewline; Write-Host "Permanent Connect (Hyper-V)" -ForegroundColor DarkYellow
        Write-Host "   [p] " -NoNewline; Write-Host "Connect via Parsec" -ForegroundColor Green
    }
    elseif ($hvStatus -eq "Temporary") {
        Write-Host "   [3] " -NoNewline; Write-Host "Temporary Connection Active" -ForegroundColor Gray
        Write-Host "   [4] " -NoNewline; Write-Host "Permanent Connect (Hyper-V)" -ForegroundColor DarkYellow
        Write-Host "   [p] " -NoNewline; Write-Host "Connect via Parsec" -ForegroundColor Green
    }
    elseif ($hvStatus -eq "Permanent" -or $hvStatus -eq "Both") {
        Write-Host "   [3] " -NoNewline; Write-Host "Temporary Connect (Disabled)" -ForegroundColor Gray
        Write-Host "   [4] " -NoNewline; Write-Host "Disconnect Permanent Connect" -ForegroundColor DarkYellow
        Write-Host "   [p] " -NoNewline; Write-Host "Connect via Parsec" -ForegroundColor Green
    }
    Write-Host "   [h] " -NoNewline; Write-Host "Open Hyper-V Manager" -ForegroundColor DarkYellow
    if (Get-ParsecStatus) {
        Write-Host "   [k] " -NoNewline; Write-Host "Close Parsec" -ForegroundColor Red
    } else {
        Write-Host "   [k] " -NoNewline; Write-Host "Open Parsec" -ForegroundColor Green
    }
    Write-Host "   [r] " -NoNewline; Write-Host "Refresh Menu" -ForegroundColor White
    Write-Host "   [s] " -NoNewline; Write-Host "Settings" -ForegroundColor DarkYellow
    Write-Host "   [v] " -NoNewline; Write-Host "Kill VM" -ForegroundColor Red
    Write-Host "   [t] " -NoNewline; Write-Host "Terminate Everything" -ForegroundColor Red
    Write-Host "   [0/q] " -NoNewline; Write-Host "Quit" -ForegroundColor Red
    Write-Host ""
}

function Show-SettingsMenu {
    do {
        Clear-Host
        Write-Host ""
        Write-Host "   ░░░ SETTINGS ░░░" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "   [1] " -NoNewline; Write-Host "Set Parsec Peer ID (Current: $($Global:Session.ParsecPeerId))" -ForegroundColor Green
        Write-Host "   [2] " -NoNewline; Write-Host "Factory Reset VM" -ForegroundColor Red
        Write-Host "   [3] " -NoNewline; Write-Host "Set Parsec Close Behavior (Current: $($Global:Session.ParsecCloseBehavior))" -ForegroundColor DarkYellow
        Write-Host "   [4] " -NoNewline; Write-Host "Toggle Safety Confirmation (Current: $($Global:Session.SafetyConfirmation))" -ForegroundColor DarkYellow

        $vm = Get-VM -Name $Global:Session.VMName
        if ($vm.State -eq "Off") {
            Write-Host "   [5] " -NoNewline; Write-Host "Enter Checkpoints Menu" -ForegroundColor DarkYellow
        }
        Write-Host "   [r] " -NoNewline; Write-Host "Refresh Settings" -ForegroundColor White
        Write-Host "   [b] " -NoNewline; Write-Host "Back to Main Menu" -ForegroundColor Gray
        Write-Host ""
        Write-Host "   Please select an option:" -ForegroundColor White

        $choice = Get-MenuChoice
        switch ($choice) {
            "1" {
                Write-Host "`nCurrent Parsec Peer ID: $($Global:Session.ParsecPeerId)" -ForegroundColor White
                Write-Host "Enter new Parsec Peer ID (or press Enter to cancel):" -ForegroundColor White
                $newId = Read-Host
                if ([string]::IsNullOrWhiteSpace($newId)) {
                    Log-Warning "Peer ID update cancelled."
                } else {
                    Set-ParsecPeerId -newPeerId $newId
                }
                # no forced 2s wait here
            }
            "2" {
                if ($vm.State -ne "Off") {
                    Log-Error "Factory reset can only be performed when the VM is off."
                } else {
                    FactoryReset-VM
                }
            }
            "3" {
                # Submenu instead of typing
                Clear-Host
                Write-Host "`nPick a parsecCloseBehavior:"
                Write-Host "   [1] alwaysClose"
                Write-Host "   [2] neverClose"
                Write-Host "   [3] prompt"
                $subchoice = Read-Host
                switch ($subchoice) {
                    "1" { Set-ParsecCloseBehavior -newBehavior "alwaysClose" }
                    "2" { Set-ParsecCloseBehavior -newBehavior "neverClose" }
                    "3" { Set-ParsecCloseBehavior -newBehavior "prompt" }
                    default {
                        Log-Error "Invalid choice."
                    }
                }
            }
            "4" {
                $curr = $Global:Session.SafetyConfirmation
                [bool]$newVal = (-not $curr)
                Set-SafetyConfirmation -enabled:$newVal
            }
            "5" {
                if ($vm.State -eq "Off") {
                    Show-CheckpointsMenu
                } else {
                    Log-Error "VM must be off to access Checkpoints Menu."
                }
            }
            "r" {
                # immediate refresh
                continue
            }
            "b" {
                # Instant return - no Sleep
                return
            }
            default {
                Log-Error "Invalid choice. Please try again."
            }
        }
        # After each setting change, we don't do a forced 2-sec wait,
        # so user returns to menu quickly.
    } while ($true)
}

function Show-CheckpointsMenu {
    # Ensure the VM is off
    $vm = Get-VM -Name $Global:Session.VMName
    if ($vm.State -ne "Off") {
        Log-Error "Checkpoints Menu is only accessible when the VM is off."
        return
    }

    # Initialize selection and search variables.
    $selectedIndex = 0
    $windowStart   = 0
    $inSearchMode  = $false
    $searchString  = ""

    while ($true) {

        # Re-read the full list of checkpoints; force it into an array.
        $allCheckpoints = @(Get-VMCheckpoint -VMName $Global:Session.VMName | Sort-Object -Property CreationTime -Descending)
        # For debugging, you can enable the following line:
        # Log-Info "AllCheckpoints count: $($allCheckpoints.Count)"

        # Apply search filtering using .ToLower().Contains().
        if ([string]::IsNullOrWhiteSpace($searchString)) {
            $filtered = $allCheckpoints
        }
        else {
            $filtered = $allCheckpoints | Where-Object { $_.Name.ToLower().Contains($searchString.ToLower()) }
        }

        # Sort the filtered list (newest first).
        $filtered = $filtered | Sort-Object -Property CreationTime -Descending

        # Clamp the selection indices.
        if ($selectedIndex -ge $filtered.Count) { $selectedIndex = $filtered.Count - 1 }
        if ($selectedIndex -lt 0) { $selectedIndex = 0 }
        if ($windowStart -gt $selectedIndex) { $windowStart = $selectedIndex }

        Clear-Host
        Write-Host ""
        Write-Host "   ░░░ CHECKPOINTS MENU ░░░" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "   Use [↑/↓] to navigate, [Enter] for actions, [b] back" -ForegroundColor White
        Write-Host "   [n] New   [o] Overwrite Selected   [s] Toggle Search (ESC clears search)" -ForegroundColor White

        # No ternary operator – use a simple if/else assignment
        if ($inSearchMode) { $mode = "ON" } else { $mode = "OFF" }
        Write-Host ("   Search Mode: {0}   |   Search String: '{1}'" -f $mode, $searchString) -ForegroundColor DarkCyan
        Write-Host ("   " + ("-" * 50))
        Write-Host ""

        if ($filtered.Count -eq 0) {
            Write-Host "   No checkpoints found (your search may be too narrow)." -ForegroundColor Red
        }
        else {
            # Display up to 5 items from the filtered list.
            $visible = $filtered | Select-Object -Skip $windowStart -First 5
            for ($i = 0; $i -lt $visible.Count; $i++) {
                $index = $windowStart + $i
                $cp    = $visible[$i]
                if ($index -eq $selectedIndex) {
                    Write-Host ("   ⇒ [{0}] {1}  -  Created: {2}" -f $index, $cp.Name, $cp.CreationTime) -ForegroundColor Green
                }
                else {
                    Write-Host ("     [{0}] {1}  -  Created: {2}" -f $index, $cp.Name, $cp.CreationTime) -ForegroundColor White
                }
            }
        }

        # Read a key from user.
        $key = [System.Console]::ReadKey($true)

        if ($inSearchMode) {
            # Process keys while in search mode.
            switch ($key.Key) {
                "Escape" {
                    $inSearchMode = $false
                    $searchString = ""
                    continue
                }
                "Backspace" {
                    if ($searchString.Length -gt 0) {
                        $searchString = $searchString.Substring(0, $searchString.Length - 1)
                    }
                    continue
                }
                "Enter" {
                    # Exit search mode, keep the current search string.
                    $inSearchMode = $false
                    continue
                }
                "UpArrow" {
                    if ($filtered.Count -gt 0 -and $selectedIndex -gt 0) {
                        $selectedIndex--
                        if ($selectedIndex -lt $windowStart) { $windowStart = $selectedIndex }
                    }
                    continue
                }
                "DownArrow" {
                    if ($filtered.Count -gt 0 -and $selectedIndex -lt ($filtered.Count - 1)) {
                        $selectedIndex++
                        if ($selectedIndex -ge ($windowStart + 5)) { $windowStart++ }
                    }
                    continue
                }
                default {
                    if ($key.KeyChar) {
                        $searchString += $key.KeyChar
                    }
                    continue
                }
            }
        }
        else {
            # Not in search mode – process regular navigation and hotkeys.
            switch ($key.Key) {
                "UpArrow" {
                    if ($filtered.Count -gt 0 -and $selectedIndex -gt 0) {
                        $selectedIndex--
                        if ($selectedIndex -lt $windowStart) { $windowStart = $selectedIndex }
                    }
                    continue
                }
                "DownArrow" {
                    if ($filtered.Count -gt 0 -and $selectedIndex -lt ($filtered.Count - 1)) {
                        $selectedIndex++
                        if ($selectedIndex -ge ($windowStart + 5)) { $windowStart++ }
                    }
                    continue
                }
            }

            $char = $key.KeyChar.ToString().ToLower()
            switch ($char) {
                "n" {
                    # Create new checkpoint.
                    Clear-Host
                    Write-Host "Enter name for new checkpoint (or press Enter to cancel):" -ForegroundColor White
                    $newName = Read-Host
                    if (-not [string]::IsNullOrWhiteSpace($newName)) {
                        try {
                            Checkpoint-VM -VMName $Global:Session.VMName -SnapshotName $newName -ErrorAction Stop
                            Log-Success "Checkpoint '$newName' created."
                        }
                        catch {
                            Log-Error "Error creating checkpoint: $_"
                        }
                        Start-Sleep -Seconds 1
                        # Immediately refresh by re-reading all checkpoints and reset indices.
                        $selectedIndex = 0
                        $windowStart = 0
                    }
                }
                "o" {
                    # Overwrite the currently selected checkpoint.
                    if ($filtered.Count -gt 0) {
                        $cpName = $filtered[$selectedIndex].Name
                        Overwrite-Checkpoint -CheckpointName $cpName
                        Start-Sleep -Seconds 1
                        $selectedIndex = 0
                        $windowStart = 0
                    }
                }
                "s" {
                    # Toggle search mode on.
                    $inSearchMode = $true
                }
                "b" {
                    return
                }
                default {
                    if ($key.Key -eq "Enter") {
                        if ($filtered.Count -gt 0) {
                            $selectedCheckpoint = $filtered[$selectedIndex]
                            Show-CheckpointActionsMenu -checkpoint $selectedCheckpoint
                        }
                    }
                }
            }
        }
    }
}

##############################################################################
# Overwrite-Checkpoint: Removes an existing checkpoint, then re-creates it
# with the same name from the current VM state. Not allowed if name=Factory Reset.
##############################################################################
function Overwrite-Checkpoint {
    param(
        [Parameter(Mandatory)]
        [string]$CheckpointName
    )

    # Don't allow overwriting 'Factory Reset'
    if ($CheckpointName -eq "Factory Reset") {
        Log-Error "Cannot overwrite the 'Factory Reset' checkpoint."
        return
    }

    # Check if the checkpoint actually exists
    $existing = Get-VMCheckpoint -VMName $Global:Session.VMName -Name $CheckpointName -ErrorAction SilentlyContinue
    if (-not $existing) {
        Log-Error "No checkpoint named '$CheckpointName' was found."
        return
    }

    # Confirm (respect safetyConfirmation)
    if (-not (Confirm-Action "Overwrite checkpoint '$CheckpointName'? (remove old + create new)")) {
        Log-Warning "Overwrite cancelled."
        return
    }

    # Remove the old checkpoint
    try {
        Remove-VMCheckpoint -VMName $Global:Session.VMName -Name $CheckpointName -Confirm:$false
        Log-Success "Removed old checkpoint '$CheckpointName'."
    }
    catch {
        Log-Error "Error removing old checkpoint '$CheckpointName': $_"
        return
    }

    # Create a new checkpoint with the same name
    try {
        Checkpoint-VM -VMName $Global:Session.VMName -SnapshotName $CheckpointName -ErrorAction Stop
        Log-Success "Overwrote checkpoint '$CheckpointName' with the current VM state."
    }
    catch {
        Log-Error "Error creating the new checkpoint: $_"
    }
}




function Show-CheckpointActionsMenu {
    param($checkpoint)
    while ($true) {
        Clear-Host
        Write-Host ""
        Write-Host "   ▀▀▀ Checkpoint Actions ▀▀▀" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "   • Selected: $($checkpoint.Name)" -ForegroundColor White
        Write-Host "     Created : $($checkpoint.CreationTime)" -ForegroundColor White
        Write-Host ""
        Write-Host "   [A] " -NoNewline; Write-Host "Apply Checkpoint" -ForegroundColor Green
        Write-Host "   [R] " -NoNewline; Write-Host "Rename Checkpoint" -ForegroundColor DarkYellow
        Write-Host "   [F] " -NoNewline; Write-Host "Set as Factory Reset" -ForegroundColor Red
        Write-Host "   [D] " -NoNewline; Write-Host "Delete Checkpoint" -ForegroundColor DarkRed
        Write-Host "   [B] " -NoNewline; Write-Host "Back" -ForegroundColor Gray
        Write-Host ""
        Write-Host "   Choose an action:" -ForegroundColor White

        $actionKey = [System.Console]::ReadKey($true)
        $action = $actionKey.KeyChar.ToString().ToLower()
        switch ($action) {
            "a" {
                if (-not (Confirm-Action "Apply checkpoint '$($checkpoint.Name)'?")) {
                    Log-Warning "Apply checkpoint cancelled."
                }
                else {
                    try {
                        Restore-VMCheckpoint -VMName $Global:Session.VMName -Name $checkpoint.Name -Confirm:$false
                        Log-Success "Checkpoint applied successfully."
                    }
                    catch {
                        Log-Error "Error applying checkpoint: $_"
                    }
                }
                return
            }
            "r" {
                Write-Host "`nEnter new name for the checkpoint (cannot be 'Factory Reset'):" -ForegroundColor White
                $newName = Read-Host
                if ($newName.ToLower() -eq "factory reset") {
                    Log-Error "Cannot rename to 'Factory Reset'."
                }
                elseif ([string]::IsNullOrWhiteSpace($newName)) {
                    Log-Warning "Rename cancelled."
                }
                else {
                    try {
                        Rename-VMCheckpoint -VMName $Global:Session.VMName -Name $checkpoint.Name -NewName $newName -ErrorAction Stop
                        Log-Success "Renamed to '$newName'."
                    }
                    catch {
                        Log-Error "Error renaming: $_"
                    }
                }
            }
            "f" {
                if (-not (Confirm-Action "Set '$($checkpoint.Name)' as Factory Reset?")) {
                    Log-Warning "Operation cancelled."
                    continue
                }
                $existingFactory = Get-VMCheckpoint -VMName $Global:Session.VMName | Where-Object { $_.Name -ieq "Factory Reset" }
                if ($existingFactory -and $existingFactory.Name -ne $checkpoint.Name) {
                    if (-not (Confirm-Action "Change existing 'Factory Reset' checkpoint?")) {
                        Log-Warning "Operation cancelled."
                        continue
                    }
                    $newNameForExisting = $existingFactory.CreationTime.ToString("yyyy-MM-dd_HH-mm-ss")
                    try {
                        Rename-VMCheckpoint -VMName $Global:Session.VMName -Name $existingFactory.Name -NewName $newNameForExisting -ErrorAction Stop
                        Log-Success "Renamed existing Factory Reset to '$newNameForExisting'."
                    }
                    catch {
                        Log-Error "Error: $_"
                    }
                }
                try {
                    Rename-VMCheckpoint -VMName $Global:Session.VMName -Name $checkpoint.Name -NewName "Factory Reset" -ErrorAction Stop
                    Log-Success "Set as Factory Reset."
                }
                catch {
                    Log-Error "Error setting Factory Reset: $_"
                }
            }
            "d" {
                if (-not (Confirm-Action "Delete checkpoint '$($checkpoint.Name)'?")) {
                    Log-Warning "Deletion cancelled."
                }
                else {
                    try {
                        Remove-VMCheckpoint -VMName $Global:Session.VMName -Name $checkpoint.Name -Confirm:$false
                        Log-Success "Checkpoint deleted."
                    }
                    catch {
                        Log-Error "Error deleting: $_"
                    }
                    return
                }
            }
            "b" { return }
            default {
                Log-Error "Invalid selection. Try again."
            }
        }
    }
}

function Flush-KeyBuffer {
    while ([System.Console]::KeyAvailable) {
        [System.Console]::ReadKey($true) | Out-Null
    }
}

function Get-MenuChoice {
    while (-not [System.Console]::KeyAvailable) {
        Start-Sleep -Milliseconds 100
    }
    return [System.Console]::ReadKey($true).KeyChar
}

function Shutdown-VMCustom {
    param([bool]$CloseConnections = $true)
    Log-Info "Attempting to shut down $($Global:Session.VMName) gracefully..."
    Stop-VM -Name $Global:Session.VMName -Confirm:$false
    Log-Info "Press 'y' during shutdown to force turn off immediately."
    $maxWait = 30
    $waitTime = 0
    while ($waitTime -lt $maxWait) {
        Start-Sleep -Seconds 1
        $waitTime++
        $currentState = (Get-VM -Name $Global:Session.VMName).State
        if ($currentState -eq "Off") { break }
        if ([System.Console]::KeyAvailable) {
            $key = [System.Console]::ReadKey($true)
            if ($key.KeyChar -eq 'y') {
                Log-Warning "User requested immediate turn off."
                break
            }
        }
    }
    if ($currentState -ne "Off") {
        Log-Error "Shutdown did not complete within $maxWait seconds. Forcing off..."
        Stop-VM -Name $Global:Session.VMName -TurnOff -Confirm:$false
        Log-Error "$($Global:Session.VMName) is now turned off."
    } else {
        Log-Success "$($Global:Session.VMName) shut down gracefully."
    }

    if ($CloseConnections) {
        if ($Global:Session.TempConnectProcess) {
            Safe-StopProcess -Process $Global:Session.TempConnectProcess -Force
            $Global:Session.TempConnectProcess = $null
            Log-Warning "Temporary Hyper-V connection closed."
        }
        if ($Global:Session.VMConnectProcess -and -not $Global:Session.VMConnectProcess.HasExited) {
            Safe-StopProcess -Process $Global:Session.VMConnectProcess -Force
            $Global:Session.VMConnectProcess = $null
            Log-Warning "Permanent Hyper-V connection closed."
        }

        # parsec close behavior
        if (Get-ParsecStatus) {
            switch ($Global:Session.ParsecCloseBehavior) {
                "alwaysClose" {
                    ParsecKill
                }
                "neverClose" {
                    Log-Success "Leaving Parsec open (neverClose)."
                }
                default {
                    $killParsec = Read-Host "Close Parsec too? (y/n)"
                    if ($killParsec -eq "y") {
                        ParsecKill
                    }
                }
            }
        } else {
            Log-Success "No Parsec processes running."
        }
    }
}

# ----------------------------------------------------------------------------------
#                              Main Logic / Menu Loop
# ----------------------------------------------------------------------------------

if ($StartVM -or $ShutdownVM -or $KillVM -or $ConnectParsec) {
    if ($StartVM) {
        Start-VMWithParsec
        return
    }
    if ($ShutdownVM) {
        Shutdown-VMCustom -CloseConnections:$true
        return
    }
    if ($KillVM) {
        VMKill
        return
    }
    if ($ConnectParsec) {
        Connect-VMParsec
        return
    }
}

if (-not $PSBoundParameters.ContainsKey('OpenMenu') -and ($StartVM -or $ShutdownVM -or $KillVM -or $ConnectParsec)) {
    return
}

do {
    if (-not (IsConsoleActive)) {
        while (-not (IsConsoleActive)) {
            Start-Sleep -Milliseconds 500
        }
        Clear-Host
    } else {
        Clear-Host
    }
    Flush-KeyBuffer

    if ($Global:Session.TempConnectProcess) {
        try { $Global:Session.TempConnectProcess.Refresh() } catch {}
        if ($Global:Session.TempConnectProcess.HasExited) {
            $Global:Session.TempConnectProcess = $null
        }
    }
    if ($Global:Session.VMConnectProcess) {
        try { $Global:Session.VMConnectProcess.Refresh() } catch {}
        if ($Global:Session.VMConnectProcess.HasExited) {
            $Global:Session.VMConnectProcess = $null
        }
    }
    
    $vm = Get-VM -Name $Global:Session.VMName
    $vmState = $vm.State
    Show-Header -vmState $vmState

    if ($vmState -eq "Running") {
        Show-RunningMenu
    } else {
        Show-OffMenu
    }

    $choice = Get-MenuChoice
    switch ($choice) {
        "1" {
            if ($vmState -eq "Running") {
                Shutdown-VMCustom -CloseConnections:$true
            } else {
                Start-VMWithParsec
            }
        }
        "2" {
            if ($vmState -eq "Running") {
                Restart-VMCustom
            } else {
                if ($Global:Session.TempConnectProcess -or $Global:Session.VMConnectProcess) {
                    Log-Warning "Connection active. Disconnect it first."
                } else {
                    Start-VMWithPermanent
                }
            }
        }
        "3" {
            if ($vmState -eq "Running") {
                if ($Global:Session.TempConnectProcess -or $Global:Session.VMConnectProcess) {
                    Log-Warning "Temporary Hyper-V connection not available while connection is active."
                } else {
                    Connect-VMTemporary
                }
            }
        }
        "4" {
            if ($vmState -eq "Running") {
                if ($Global:Session.VMConnectProcess) {
                    Disconnect-VM
                } else {
                    Connect-VMPermanent
                }
            }
        }
        "p" {
            if ($vmState -eq "Running") {
                Connect-VMParsec
            }
        }
        "k" {
            Handle-ParsecToggle
        }
        "h" {
            Open-HyperVManager
        }
        "s" {
            Show-SettingsMenu
        }
        "t" {
            Terminate-Everything
        }
        "v" {
            if ($vmState -eq "Running") {
                VMKill
            }
        }
        "q" { $choice = "0" }
        "r" { }  # refresh
        "0" {
            Log-Error "Exiting menu."
        }
        default {
            Log-Error "Invalid choice. Please try again."
        }
    }
} while ($choice -ne "0")
exit 0
