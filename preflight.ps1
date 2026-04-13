# VIGIL Pre-Flight Checklist
# Run on target Windows machine to verify environment before building VIGIL.
# Usage:  powershell -ExecutionPolicy Bypass -File .\preflight.ps1

$ErrorActionPreference = 'Continue'
$results = [ordered]@{}

function Write-Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $status = if ($Ok) { '[ OK ]' } else { '[FAIL]' }
    $color  = if ($Ok) { 'Green' }  else { 'Red'   }
    Write-Host ("{0} {1}" -f $status, $Name) -ForegroundColor $color
    if ($Detail) { Write-Host "       $Detail" -ForegroundColor DarkGray }
    $script:results[$Name] = [pscustomobject]@{ Ok = $Ok; Detail = $Detail }
}

Write-Host ""
Write-Host "VIGIL Pre-Flight Checklist" -ForegroundColor Cyan
Write-Host "==========================" -ForegroundColor Cyan
Write-Host ""

# 1. WPF availability
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName PresentationCore      -ErrorAction Stop
    Add-Type -AssemblyName WindowsBase            -ErrorAction Stop
    [void][System.Windows.MessageBox]::Show('WPF OK', 'VIGIL Pre-Check', 'OK', 'Information')
    Write-Check 'WPF assemblies loadable' $true 'PresentationFramework + PresentationCore + WindowsBase'
} catch {
    Write-Check 'WPF assemblies loadable' $false $_.Exception.Message
}

# 2. PowerShell version (target 5.1+)
$ver = $PSVersionTable.PSVersion
$psOk = ($ver.Major -ge 5)
Write-Check 'PowerShell 5.1+' $psOk "Detected $ver"

# 3. Outlook COM
try {
    $ol  = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns  = $ol.GetNamespace('MAPI')
    $cal = $ns.GetDefaultFolder(9)   # olFolderCalendar
    $inb = $ns.GetDefaultFolder(6)   # olFolderInbox
    $tks = $ns.GetDefaultFolder(13)  # olFolderTasks
    $detail = "Calendar=$($cal.Items.Count)  Inbox=$($inb.Items.Count)  Tasks=$($tks.Items.Count)"
    Write-Check 'Outlook COM (MAPI + folders 9/6/13)' $true $detail
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol) | Out-Null
} catch {
    Write-Check 'Outlook COM (MAPI + folders 9/6/13)' $false $_.Exception.Message
}

# 4. user32.dll P/Invoke (RegisterHotKey / UnregisterHotKey)
try {
    if (-not ([System.Management.Automation.PSTypeName]'VigilHotkeyTest').Type) {
        Add-Type -ErrorAction Stop -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class VigilHotkeyTest {
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
    }
    $fg = [VigilHotkeyTest]::GetForegroundWindow()
    Write-Check 'user32.dll P/Invoke (RegisterHotKey et al.)' $true "Foreground hWnd=$fg"
} catch {
    Write-Check 'user32.dll P/Invoke (RegisterHotKey et al.)' $false $_.Exception.Message
}

# 5. Startup folder access (for auto-start .lnk)
try {
    $startup = [Environment]::GetFolderPath('Startup')
    $exists  = Test-Path $startup
    Write-Check 'Startup folder accessible' $exists $startup
} catch {
    Write-Check 'Startup folder accessible' $false $_.Exception.Message
}

# 6. ~/.vigil writable
try {
    $vigilDir = Join-Path $env:USERPROFILE '.vigil'
    if (-not (Test-Path $vigilDir)) { New-Item -ItemType Directory -Path $vigilDir -Force | Out-Null }
    $probe = Join-Path $vigilDir '.write-probe'
    Set-Content -Path $probe -Value 'ok' -Encoding UTF8
    Remove-Item $probe -Force
    Write-Check '~/.vigil writable' $true $vigilDir
} catch {
    Write-Check '~/.vigil writable' $false $_.Exception.Message
}

# Summary
$total  = $results.Count
$passed = ($results.Values | Where-Object { $_.Ok }).Count
Write-Host ""
Write-Host ("Result: {0}/{1} checks passed" -f $passed, $total) -ForegroundColor Cyan
if ($passed -eq $total) {
    Write-Host "Ready to build VIGIL." -ForegroundColor Green
    exit 0
} else {
    Write-Host "Resolve failing checks before building." -ForegroundColor Yellow
    exit 1
}
