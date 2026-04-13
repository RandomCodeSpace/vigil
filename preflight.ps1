# VIGIL Pre-Flight Checklist
# Run on target Windows machine to verify environment before building VIGIL.
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\preflight.ps1
#   powershell -ExecutionPolicy Bypass -File .\preflight.ps1 -TenantId <guid>
#
# At the end of the run, a single compact result string is printed (line
# starting with "VIGIL:v1:"). Copy ONLY that line back for remote triage —
# it encodes pass/fail for every check as a bitmap.

[CmdletBinding()]
param(
    [string]$TenantId = '',     # Optional: expected Azure AD tenant GUID
    [switch]$Quiet               # Suppress per-check output, only emit result string
)

$ErrorActionPreference = 'Continue'
$results = [ordered]@{}
$script:CheckOrder = New-Object System.Collections.Generic.List[string]

function Write-Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $status = if ($Ok) { '[ OK ]' } else { '[FAIL]' }
    $color  = if ($Ok) { 'Green' }  else { 'Red'   }
    if (-not $Quiet) {
        Write-Host ("{0} {1}" -f $status, $Name) -ForegroundColor $color
        if ($Detail) { Write-Host "       $Detail" -ForegroundColor DarkGray }
    }
    $script:results[$Name] = [pscustomobject]@{ Ok = $Ok; Detail = $Detail }
    [void]$script:CheckOrder.Add($Name)
}

Write-Host ""
Write-Host "VIGIL Pre-Flight Checklist" -ForegroundColor Cyan
Write-Host "==========================" -ForegroundColor Cyan
Write-Host ""

# 1. WPF availability — silent load + type construction (no MessageBox popup)
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName PresentationCore      -ErrorAction Stop
    Add-Type -AssemblyName WindowsBase            -ErrorAction Stop
    # Construct a WPF type without rendering anything: proves the stack works.
    $probeWin = New-Object System.Windows.Window
    $probeWin.Width = 1; $probeWin.Height = 1
    $probeWin = $null
    Write-Check 'WPF assemblies loadable' $true 'PresentationFramework + PresentationCore + WindowsBase'
} catch {
    Write-Check 'WPF assemblies loadable' $false $_.Exception.Message
}

# 2. PowerShell version (target 5.1+)
$ver = $PSVersionTable.PSVersion
$psOk = ($ver.Major -ge 5)
Write-Check 'PowerShell 5.1+' $psOk "Detected $ver"

# 3. Outlook COM — fully released in reverse order (no leaked RCWs)
$ol = $ns = $cal = $inb = $tks = $calItems = $inbItems = $tksItems = $null
try {
    $ol  = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns  = $ol.GetNamespace('MAPI')
    $cal = $ns.GetDefaultFolder(9)   # olFolderCalendar
    $inb = $ns.GetDefaultFolder(6)   # olFolderInbox
    $tks = $ns.GetDefaultFolder(13)  # olFolderTasks
    $calItems = $cal.Items
    $inbItems = $inb.Items
    $tksItems = $tks.Items
    $detail = "Calendar=$($calItems.Count)  Inbox=$($inbItems.Count)  Tasks=$($tksItems.Count)"
    Write-Check 'Outlook COM (MAPI + folders 9/6/13)' $true $detail
} catch {
    Write-Check 'Outlook COM (MAPI + folders 9/6/13)' $false $_.Exception.Message
} finally {
    foreach ($o in @($calItems, $inbItems, $tksItems, $cal, $inb, $tks, $ns, $ol)) {
        if ($null -ne $o) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {}
        }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
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

# 7. Atomic rename primitive: [System.IO.File]::Replace
#    Eng review §2A: Move-Item -Force is NOT atomic on Windows. Replace is.
try {
    $a = Join-Path $vigilDir '.replace-a'
    $b = Join-Path $vigilDir '.replace-b'
    $c = Join-Path $vigilDir '.replace-c'
    [System.IO.File]::WriteAllText($a, 'new', [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllText($b, 'old', [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::Replace($a, $b, $c)
    $ok = ((Get-Content $b -Raw) -eq 'new') -and ((Get-Content $c -Raw) -eq 'old')
    Remove-Item $b, $c -ErrorAction SilentlyContinue
    Write-Check '[IO.File]::Replace atomic rename + backup' $ok 'Required for crash-safe tasks.json writes'
} catch {
    Write-Check '[IO.File]::Replace atomic rename + backup' $false $_.Exception.Message
}

# 8. UTF-8 without BOM writing
#    Eng review §2B: BOM breaks many JSON parsers. Verify BOM-less write works.
try {
    $noBomFile = Join-Path $vigilDir '.nobom-probe.json'
    [System.IO.File]::WriteAllText($noBomFile, '{"ok":true}', [System.Text.UTF8Encoding]::new($false))
    $bytes = [System.IO.File]::ReadAllBytes($noBomFile)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    Remove-Item $noBomFile -ErrorAction SilentlyContinue
    Write-Check 'UTF-8 without BOM write' (-not $hasBom) "BOM present: $hasBom"
} catch {
    Write-Check 'UTF-8 without BOM write' $false $_.Exception.Message
}

# 9. Marshal.ReleaseComObject available (Outlook COM lifecycle, §1A)
try {
    $t = [System.Runtime.InteropServices.Marshal]
    $hasRelease = ($null -ne $t.GetMethod('ReleaseComObject'))
    Write-Check 'Marshal.ReleaseComObject available' $hasRelease 'Required to avoid Outlook.exe zombie process'
} catch {
    Write-Check 'Marshal.ReleaseComObject available' $false $_.Exception.Message
}

# 10. Calendar Sort-before-Restrict pattern (Outlook COM gotcha, §1B)
#     Verifies IncludeRecurrences + Sort + Restrict flow actually returns items.
try {
    $ol  = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns  = $ol.GetNamespace('MAPI')
    $cal = $ns.GetDefaultFolder(9)
    $items = $cal.Items
    $items.IncludeRecurrences = $true
    $items.Sort('[Start]')
    $start = (Get-Date).ToString('g')
    $end   = (Get-Date).AddHours(24).ToString('g')
    $filter = "[Start] >= '$start' AND [Start] <= '$end'"
    $restricted = $items.Restrict($filter)
    $count = $restricted.Count
    Write-Check 'Calendar Sort-before-Restrict returns items' $true "Next 24h meetings: $count"
    foreach ($o in @($restricted, $items, $cal, $ns, $ol)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
} catch {
    Write-Check 'Calendar Sort-before-Restrict returns items' $false $_.Exception.Message
}

# 11. Outlook EntryID readable on a flagged email (dedup key, §2E)
try {
    $ol  = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns  = $ol.GetNamespace('MAPI')
    $inb = $ns.GetDefaultFolder(6)
    $flagged = $inb.Items.Restrict("[FlagStatus] = 2")
    $hasId = $false
    if ($flagged.Count -gt 0) {
        $first = $flagged.Item(1)
        $hasId = -not [string]::IsNullOrEmpty($first.EntryID)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($first)
    } else {
        $hasId = $true  # no flagged items is not a failure
    }
    foreach ($o in @($flagged, $inb, $ns, $ol)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Check 'Outlook EntryID readable on flagged items' $hasId 'Used as stable dedup key'
} catch {
    Write-Check 'Outlook EntryID readable on flagged items' $false $_.Exception.Message
}

# 12. Named mutex create/release (single-instance, §1D)
try {
    $createdNew = $false
    $mx = New-Object System.Threading.Mutex($true, 'Global\VIGIL_PreflightProbe', [ref]$createdNew)
    if ($createdNew) {
        $mx.ReleaseMutex()
        $mx.Dispose()
        Write-Check 'Global named mutex (single-instance)' $true 'Global\VIGIL_PreflightProbe'
    } else {
        $mx.Dispose()
        Write-Check 'Global named mutex (single-instance)' $false 'Mutex already held'
    }
} catch {
    Write-Check 'Global named mutex (single-instance)' $false $_.Exception.Message
}

# 13. FindWindow / SetForegroundWindow P/Invoke (activate existing instance, §1D)
try {
    if (-not ([System.Management.Automation.PSTypeName]'VigilWin32').Type) {
        Add-Type -ErrorAction Stop -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class VigilWin32 {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
    }
    [void][VigilWin32]::FindWindow($null, 'Program Manager')
    Write-Check 'FindWindow / SetForegroundWindow P/Invoke' $true 'Used to activate existing VIGIL window'
} catch {
    Write-Check 'FindWindow / SetForegroundWindow P/Invoke' $false $_.Exception.Message
}

# 14. Clipboard read via WPF (used by Ctrl+Win+A flow, §7.3)
try {
    $clipOk = $true
    try { [void][System.Windows.Clipboard]::GetText() } catch { $clipOk = $false }
    Write-Check 'WPF clipboard read' $clipOk 'Quick-Add auto-fill source'
} catch {
    Write-Check 'WPF clipboard read' $false $_.Exception.Message
}

# 15. Pester available (tests in plan review)
try {
    $pester = Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 1
    $ok = $null -ne $pester
    $detail = if ($ok) { "Pester $($pester.Version)" } else { 'Not installed — needed for Phase 1 test bar' }
    Write-Check 'Pester test framework' $ok $detail
} catch {
    Write-Check 'Pester test framework' $false $_.Exception.Message
}

# 16. Working area / off-screen clamp data (WinForms, §14 risk row)
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $detail = "Primary working area: $($wa.Width)x$($wa.Height) @ ($($wa.X),$($wa.Y))"
    Write-Check 'Screen.WorkingArea available (clamp on-screen)' $true $detail
} catch {
    Write-Check 'Screen.WorkingArea available (clamp on-screen)' $false $_.Exception.Message
}

# 17. DispatcherTimer available (15-min Outlook sync, §6.3)
try {
    $dt = New-Object System.Windows.Threading.DispatcherTimer
    $dt.Interval = [TimeSpan]::FromMinutes(15)
    Write-Check 'DispatcherTimer available' $true '15-min sync scheduler'
} catch {
    Write-Check 'DispatcherTimer available' $false $_.Exception.Message
}

# 18. WScript.Shell for Startup shortcut creation (§9.2)
try {
    $wsh = New-Object -ComObject WScript.Shell -ErrorAction Stop
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh)
    Write-Check 'WScript.Shell COM (shortcut creation)' $true 'For VIGIL.lnk in Startup folder'
} catch {
    Write-Check 'WScript.Shell COM (shortcut creation)' $false $_.Exception.Message
}

# 19. ExecutionPolicy check (documentation, §13 risk row)
try {
    $ep = Get-ExecutionPolicy -Scope CurrentUser
    $restricted = ($ep -eq 'Restricted' -or $ep -eq 'AllSigned')
    $detail = "CurrentUser scope: $ep"
    if ($restricted) { $detail += ' — launch VIGIL with -ExecutionPolicy Bypass' }
    Write-Check 'ExecutionPolicy allows script run' (-not $restricted) $detail
} catch {
    Write-Check 'ExecutionPolicy allows script run' $false $_.Exception.Message
}

# --- Corporate lockdown checks (firewall / GPO / AppLocker / WDAC) ---

# 20. Constrained Language Mode — the #1 corp-lockdown killer for VIGIL.
#     CLM blocks Add-Type, New-Object COM, and all P/Invoke. If this fails,
#     VIGIL cannot run on this machine at all and needs IT intervention.
try {
    $mode = $ExecutionContext.SessionState.LanguageMode
    $ok = ($mode -eq 'FullLanguage')
    $detail = "LanguageMode = $mode"
    if (-not $ok) { $detail += ' — blocks Add-Type, COM, P/Invoke. VIGIL cannot run.' }
    Write-Check 'PowerShell FullLanguage mode (not CLM)' $ok $detail
} catch {
    Write-Check 'PowerShell FullLanguage mode (not CLM)' $false $_.Exception.Message
}

# 21. AppLocker script rules — would block running VIGIL.ps1 from user profile.
try {
    $testPath = Join-Path $env:USERPROFILE '.vigil\VIGIL.ps1'
    if (-not (Test-Path (Split-Path $testPath))) {
        New-Item -ItemType Directory -Path (Split-Path $testPath) -Force | Out-Null
    }
    if (-not (Test-Path $testPath)) { Set-Content -Path $testPath -Value '# probe' -Encoding UTF8 }
    $applockerCmd = Get-Command Test-AppLockerPolicy -ErrorAction SilentlyContinue
    if ($applockerCmd) {
        $r = Test-AppLockerPolicy -Path $testPath -User $env:USERNAME -ErrorAction Stop
        $allowed = ($r.PolicyDecision -eq 'Allowed' -or $r.PolicyDecision -eq 'AllowedByDefault')
        Write-Check 'AppLocker allows .ps1 from user profile' $allowed "PolicyDecision = $($r.PolicyDecision)"
    } else {
        Write-Check 'AppLocker allows .ps1 from user profile' $true 'Test-AppLockerPolicy not available — assuming no AppLocker'
    }
} catch {
    Write-Check 'AppLocker allows .ps1 from user profile' $false $_.Exception.Message
}

# 22. AMSI / EDR does not flag inline C# Add-Type (used for P/Invoke).
#     If aggressive EDR blocks Add-Type, VIGIL's hotkey + window activation die.
#     Guard against "type already exists" on repeat runs in same PS session.
try {
    if (-not ([System.Management.Automation.PSTypeName]'VigilAmsiProbe').Type) {
        Add-Type -ErrorAction Stop -TypeDefinition @"
public class VigilAmsiProbe { public static int Answer() { return 42; } }
"@
    }
    $ok = ([VigilAmsiProbe]::Answer() -eq 42)
    Write-Check 'Add-Type inline C# not blocked by AMSI/EDR' $ok 'Required for hotkey + FindWindow P/Invoke'
} catch {
    Write-Check 'Add-Type inline C# not blocked by AMSI/EDR' $false $_.Exception.Message
}

# 23. Office COM automation permitted by GPO.
#     Some corp policies disable ProgID registration for Outlook.Application.
try {
    $progId = [Type]::GetTypeFromProgID('Outlook.Application')
    $ok = ($null -ne $progId)
    Write-Check 'Outlook.Application ProgID registered' $ok 'GPO may block Office COM automation'
} catch {
    Write-Check 'Outlook.Application ProgID registered' $false $_.Exception.Message
}

# 24. Windows Defender real-time scan not blocking ~/.vigil
#     Some EDRs quarantine scripts under %USERPROFILE% as untrusted.
try {
    $probe = Join-Path $env:USERPROFILE '.vigil\.defender-probe.ps1'
    Set-Content -Path $probe -Value '# harmless probe' -Encoding UTF8
    Start-Sleep -Milliseconds 200
    $stillThere = Test-Path $probe
    if ($stillThere) { Remove-Item $probe -Force }
    Write-Check 'Defender/EDR does not quarantine ~/.vigil scripts' $stillThere 'Probe file survived write+read cycle'
} catch {
    Write-Check 'Defender/EDR does not quarantine ~/.vigil scripts' $false $_.Exception.Message
}

# 25. Cascadia Mono / Consolas font available (§5.1)
try {
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $installed = (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }
    $hasCascadia = $installed -contains 'Cascadia Mono'
    $hasConsolas = $installed -contains 'Consolas'
    $ok = $hasCascadia -or $hasConsolas
    $detail = "Cascadia Mono: $hasCascadia, Consolas: $hasConsolas"
    Write-Check 'Monospace font available (Cascadia/Consolas)' $ok $detail
} catch {
    Write-Check 'Monospace font available (Cascadia/Consolas)' $false $_.Exception.Message
}

# --- Environment / Azure AD / identity checks ---

# 26. Machine Azure AD / Hybrid join state via dsregcmd (local, no network)
$script:DetectedTenantId = ''
try {
    $dsreg = & dsregcmd /status 2>$null | Out-String
    $aadJoined = ($dsreg -match 'AzureAdJoined\s*:\s*YES')
    $domJoined = ($dsreg -match 'DomainJoined\s*:\s*YES')
    if ($dsreg -match 'TenantId\s*:\s*([0-9a-fA-F\-]{36})') { $script:DetectedTenantId = $Matches[1] }
    $kind = @()
    if ($aadJoined) { $kind += 'AzureAD' }
    if ($domJoined) { $kind += 'Domain'  }
    if (-not $kind)  { $kind += 'Workgroup' }
    $detail = "Join = $($kind -join '+')"
    if ($script:DetectedTenantId) { $detail += "  TenantId = $script:DetectedTenantId" }
    Write-Check 'Device join state (dsregcmd)' $true $detail
} catch {
    Write-Check 'Device join state (dsregcmd)' $false $_.Exception.Message
}

# 27. Tenant ID matches expected (only runs if -TenantId was passed)
try {
    if ([string]::IsNullOrWhiteSpace($TenantId)) {
        Write-Check 'Tenant ID matches expected' $true 'No -TenantId supplied — skipped'
    } else {
        $match = ($script:DetectedTenantId -and
                  ($script:DetectedTenantId.ToLower() -eq $TenantId.ToLower()))
        $detail = "Expected $TenantId, got $script:DetectedTenantId"
        Write-Check 'Tenant ID matches expected' $match $detail
    }
} catch {
    Write-Check 'Tenant ID matches expected' $false $_.Exception.Message
}

# 28. Outlook profile is configured for a mailbox (MAPI store present)
try {
    $ol  = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns  = $ol.GetNamespace('MAPI')
    $stores = $ns.Stores
    $count = 0
    foreach ($s in $stores) { $count++; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($s) }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($stores)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Check 'Outlook profile has at least one mail store' ($count -ge 1) "$count store(s)"
} catch {
    Write-Check 'Outlook profile has at least one mail store' $false $_.Exception.Message
}

# 29. .NET Framework >= 4.7.2 (WPF transparency + backdrop features)
try {
    $release = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction Stop).Release
    $ok = ($release -ge 461808)
    Write-Check '.NET Framework >= 4.7.2' $ok "Release key = $release"
} catch {
    Write-Check '.NET Framework >= 4.7.2' $false $_.Exception.Message
}

# 30. Task Scheduler COM (Schedule.Service) — fallback for auto-start if
#     Startup folder .lnk is blocked by GPO.
try {
    $sch = New-Object -ComObject Schedule.Service -ErrorAction Stop
    $sch.Connect()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sch)
    Write-Check 'Task Scheduler COM (Schedule.Service)' $true 'Fallback auto-start path'
} catch {
    Write-Check 'Task Scheduler COM (Schedule.Service)' $false $_.Exception.Message
}

# 31. BitLocker on system drive (signals whether plaintext tasks.json is acceptable)
try {
    $vol = Get-CimInstance -Namespace 'root/cimv2/security/microsoftvolumeencryption' `
                           -ClassName Win32_EncryptableVolume `
                           -Filter "DriveLetter='$env:SystemDrive'" -ErrorAction Stop
    $on = ($null -ne $vol -and $vol.ProtectionStatus -eq 1)
    Write-Check 'BitLocker on system drive' $on 'If off, consider DPAPI encryption of tasks.json'
} catch {
    Write-Check 'BitLocker on system drive' $false $_.Exception.Message
}

# 32. Windows Firewall profile (informational — VIGIL makes no network calls
#     so blocked outbound is fine, but we record the state for triage).
try {
    $fw = Get-NetFirewallProfile -ErrorAction Stop | Select-Object Name,Enabled
    $active = ($fw | Where-Object { $_.Enabled -eq $true } | ForEach-Object Name) -join ','
    Write-Check 'Windows Firewall enabled profiles' $true "Active: $active"
} catch {
    # Older Win10 builds or restricted shells may lack Get-NetFirewallProfile
    Write-Check 'Windows Firewall enabled profiles' $true 'Get-NetFirewallProfile unavailable — skipped'
}

# 33. Proxy configuration (informational — zero-network app, but useful triage)
try {
    $proxy = [System.Net.WebRequest]::DefaultWebProxy
    $viaProxy = $proxy.GetProxy('https://login.microsoftonline.com').AbsoluteUri
    Write-Check 'System proxy visible' $true "Resolved: $viaProxy"
} catch {
    Write-Check 'System proxy visible' $true 'No proxy resolver — direct connection assumed'
}

# 34. Microsoft.Graph PowerShell module (informational — NOT required by VIGIL,
#     but if present tells us Graph fallback is feasible in Phase 5).
try {
    $graph = Get-Module -ListAvailable -Name Microsoft.Graph* | Select-Object -First 1
    $ok = $null -ne $graph
    $detail = if ($ok) { "Microsoft.Graph $($graph.Version) installed (Phase 5 option)" }
              else     { 'Not installed — Phase 5 Graph fallback would need IT approval' }
    Write-Check 'Microsoft.Graph module (informational)' $true $detail
} catch {
    Write-Check 'Microsoft.Graph module (informational)' $true $_.Exception.Message
}

# 35. Script running as non-admin (VIGIL must not need elevation)
try {
    $id  = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $pri = New-Object System.Security.Principal.WindowsPrincipal($id)
    $isAdmin = $pri.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-Check 'Running as non-admin (as intended)' (-not $isAdmin) "IsAdmin = $isAdmin"
} catch {
    Write-Check 'Running as non-admin (as intended)' $false $_.Exception.Message
}

# --- Extended environment inspection (v2 additions, checks 36-55) ---

# 36. Windows version + build — gates Mica (22000+) and PerMonitorV2 DPI
try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $build = [int]($os.BuildNumber)
    $caption = $os.Caption
    $isWin11 = ($build -ge 22000)
    $detail = "$caption build $build" + $(if ($isWin11) { ' (Win11 — Mica OK)' } else { ' (Win10 — flat glass fallback)' })
    Write-Check 'Windows version + build' $true $detail
} catch {
    Write-Check 'Windows version + build' $false $_.Exception.Message
}

# 37. Outlook version + bitness (x64 vs x86 affects COM marshalling choices)
try {
    $ol = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ver = $ol.Version
    $exe = $ol.Path
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Check 'Outlook version + install path' $true "v$ver at $exe"
} catch {
    Write-Check 'Outlook version + install path' $false $_.Exception.Message
}

# 38. Outlook currently running (affects first-sync latency strategy)
try {
    $running = $null -ne (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    $detail = if ($running) { 'Running — GetActiveObject path' } else { 'Not running — startup sync will launch Outlook' }
    Write-Check 'Outlook process currently running' $true $detail
} catch {
    Write-Check 'Outlook process currently running' $true $_.Exception.Message
}

# 39. Windows theme (light/dark) — auto-theme for Apple reskin
try {
    $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
    $appsLight = (Get-ItemProperty -Path $k -Name AppsUseLightTheme -ErrorAction Stop).AppsUseLightTheme
    $theme = if ($appsLight -eq 0) { 'dark' } else { 'light' }
    Write-Check 'Windows app theme (light/dark)' $true "Theme = $theme"
} catch {
    Write-Check 'Windows app theme (light/dark)' $true 'Key absent — defaulting to dark'
}

# 40. Transparency effects enabled — if OFF, skip glass/blur entirely
try {
    $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize'
    $tOn = (Get-ItemProperty -Path $k -Name EnableTransparency -ErrorAction Stop).EnableTransparency
    $ok = ($tOn -eq 1)
    Write-Check 'Transparency effects enabled' $ok "EnableTransparency = $tOn"
} catch {
    Write-Check 'Transparency effects enabled' $true 'Key absent — assume on'
}

# 41. System accent color (informational — VIGIL ignores, but good triage)
try {
    $k = 'HKCU:\Software\Microsoft\Windows\DWM'
    $accent = (Get-ItemProperty -Path $k -Name ColorizationColor -ErrorAction Stop).ColorizationColor
    $hex = ('#{0:X8}' -f $accent)
    Write-Check 'System accent color' $true "DWM ColorizationColor = $hex"
} catch {
    Write-Check 'System accent color' $true 'Not readable — default accent'
}

# 42. Display count + primary resolution (multi-monitor quick-add popup)
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $primary = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    Write-Check 'Display layout' $true "$($screens.Count) display(s), primary $($primary.Width)x$($primary.Height)"
} catch {
    Write-Check 'Display layout' $false $_.Exception.Message
}

# 43. System DPI scale factor (WPF widget pixel sizing)
try {
    if (-not ([System.Management.Automation.PSTypeName]'VigilDpi').Type) {
        Add-Type -ErrorAction Stop -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class VigilDpi {
    [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
    [DllImport("gdi32.dll")]  public static extern int GetDeviceCaps(IntPtr hDC, int nIndex);
}
"@
    }
    $dc = [VigilDpi]::GetDC([IntPtr]::Zero)
    $dpi = [VigilDpi]::GetDeviceCaps($dc, 88)  # LOGPIXELSX
    [void][VigilDpi]::ReleaseDC([IntPtr]::Zero, $dc)
    $scale = [math]::Round($dpi / 96.0, 2)
    Write-Check 'System DPI + scale factor' $true "$dpi DPI (${scale}x scale)"
} catch {
    Write-Check 'System DPI + scale factor' $false $_.Exception.Message
}

# 44. High contrast mode (accessibility — disables translucency)
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $hc = [System.Windows.Forms.SystemInformation]::HighContrast
    Write-Check 'High contrast mode OFF' (-not $hc) "HighContrast = $hc"
} catch {
    Write-Check 'High contrast mode OFF' $true $_.Exception.Message
}

# 45. Reduced motion / client animations enabled (accessibility)
try {
    $k = 'HKCU:\Control Panel\Desktop\WindowMetrics'
    $anim = (Get-ItemProperty -Path $k -Name MinAnimate -ErrorAction Stop).MinAnimate
    $on = ($anim -eq '1')
    Write-Check 'Window animations enabled' $on "MinAnimate = $anim"
} catch {
    Write-Check 'Window animations enabled' $true 'Key absent — assume on'
}

# 46. WPF render tier (0=software, 1=partial hw, 2=full hw)
try {
    Add-Type -AssemblyName PresentationCore -ErrorAction Stop
    $tier = [System.Windows.Media.RenderCapability]::Tier -shr 16
    $label = @('software','partial','full')[[math]::Min(2, $tier)]
    Write-Check 'WPF render tier' $true "Tier $tier ($label)"
} catch {
    Write-Check 'WPF render tier' $false $_.Exception.Message
}

# 47. RDP / Citrix / virtual session detection (widget strategy differs)
try {
    $sessionName = $env:SESSIONNAME
    $isRemote = ($sessionName -like 'RDP-*') -or ($sessionName -like 'ICA-*')
    $detail = "SESSIONNAME = $sessionName" + $(if ($isRemote) { ' (remote session)' } else { ' (console)' })
    Write-Check 'Console (non-remote) session' (-not $isRemote) $detail
} catch {
    Write-Check 'Console (non-remote) session' $true $_.Exception.Message
}

# 48. Power state (battery vs plugged in — informational)
try {
    $batt = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
    if ($null -eq $batt) {
        Write-Check 'Power state' $true 'Desktop (no battery)'
    } else {
        $pct = $batt.EstimatedChargeRemaining
        $onAc = ($batt.BatteryStatus -eq 2)
        $state = if ($onAc) { 'on AC' } else { 'on battery' }
        Write-Check 'Power state' $true "Laptop, $state, $pct% charge"
    }
} catch {
    Write-Check 'Power state' $true $_.Exception.Message
}

# 49. Locale + time zone (due-date parsing)
try {
    $cult = (Get-Culture).Name
    $tz = (Get-TimeZone).Id
    Write-Check 'Locale + time zone' $true "$cult / $tz"
} catch {
    Write-Check 'Locale + time zone' $false $_.Exception.Message
}

# 50. PowerShell host bitness (x64 vs x86 affects COM marshalling)
try {
    $is64 = [Environment]::Is64BitProcess
    $osIs64 = [Environment]::Is64BitOperatingSystem
    $detail = "Process: $(if ($is64) {'x64'} else {'x86'}), OS: $(if ($osIs64) {'x64'} else {'x86'})"
    Write-Check 'PowerShell host bitness' $true $detail
} catch {
    Write-Check 'PowerShell host bitness' $false $_.Exception.Message
}

# 51. Long path support enabled (affects deep %USERPROFILE% paths)
try {
    $k = 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem'
    $lp = (Get-ItemProperty -Path $k -Name LongPathsEnabled -ErrorAction SilentlyContinue).LongPathsEnabled
    $ok = ($lp -eq 1)
    Write-Check 'Long path support' $ok "LongPathsEnabled = $lp"
} catch {
    Write-Check 'Long path support' $true 'Key absent — assume off (260 char limit applies)'
}

# 52. Free disk space on user profile drive (>100 MB for safety margin)
try {
    $drive = (Get-Item $env:USERPROFILE).PSDrive.Name
    $free = (Get-PSDrive $drive).Free
    $mb = [math]::Round($free / 1MB)
    $ok = ($mb -ge 100)
    Write-Check "Free space on $drive`: drive" $ok "$mb MB free"
} catch {
    Write-Check 'Free space on user profile drive' $false $_.Exception.Message
}

# 53. TEMP directory writable (fallback if ~/.vigil is locked)
try {
    $probe = Join-Path $env:TEMP ".vigil-temp-probe-$([Guid]::NewGuid().ToString('N'))"
    Set-Content -Path $probe -Value 'ok' -Encoding UTF8
    Remove-Item $probe -Force
    Write-Check 'TEMP directory writable' $true $env:TEMP
} catch {
    Write-Check 'TEMP directory writable' $false $_.Exception.Message
}

# 54. Focus Assist state (informational — affects future toast feature)
try {
    $k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings'
    $fa = (Get-ItemProperty -Path $k -Name NOC_GLOBAL_SETTING_TOASTS_ENABLED -ErrorAction SilentlyContinue).NOC_GLOBAL_SETTING_TOASTS_ENABLED
    $toastsOn = ($fa -ne 0)
    Write-Check 'Toast notifications allowed (global)' $toastsOn "NOC_GLOBAL_SETTING_TOASTS_ENABLED = $fa"
} catch {
    Write-Check 'Toast notifications allowed (global)' $true 'Key absent — assume on'
}

# 55. Existing VIGIL install detection (prior version?)
try {
    $lnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'VIGIL.lnk'
    $ps1 = Join-Path $env:USERPROFILE '.vigil\VIGIL.ps1'
    $tasks = Join-Path $env:USERPROFILE '.vigil\tasks.json'
    $has = (Test-Path $lnk) -or (Test-Path $ps1) -or (Test-Path $tasks)
    $detail = if ($has) { 'Prior install found — upgrade path' } else { 'Clean machine' }
    Write-Check 'Prior VIGIL install state' $true $detail
} catch {
    Write-Check 'Prior VIGIL install state' $true $_.Exception.Message
}

# --- Email-specific inspection (checks 56-60) ---

# 56. Flagged email count + due-date presence sample.
#     Tells me how much data Phase 3 flag-sync will handle and whether flags
#     routinely carry TaskDueDate (affects dueDate mapping fallback).
try {
    $ol  = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns  = $ol.GetNamespace('MAPI')
    $inb = $ns.GetDefaultFolder(6)
    $flagged = $inb.Items.Restrict('[FlagStatus] = 2')
    $count = $flagged.Count
    $withDue = 0
    $sampleMax = [math]::Min(10, $count)
    for ($i = 1; $i -le $sampleMax; $i++) {
        $it = $flagged.Item($i)
        if ($it.TaskDueDate -and $it.TaskDueDate.Year -gt 2000) { $withDue++ }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($it)
    }
    foreach ($o in @($flagged, $inb, $ns, $ol)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Check 'Flagged emails + due-date sample' $true "$count flagged; $withDue/$sampleMax sampled have TaskDueDate"
} catch {
    Write-Check 'Flagged emails + due-date sample' $false $_.Exception.Message
}

# 57. Cached Exchange Mode enabled (affects sync latency + offline behavior)
try {
    $found = $false; $enabled = $false; $detail = 'No Outlook profile keys found'
    $officeVersions = '16.0','15.0','14.0'
    foreach ($v in $officeVersions) {
        $base = "HKCU:\Software\Microsoft\Office\$v\Outlook\Cached Mode"
        if (Test-Path $base) {
            $found = $true
            $val = (Get-ItemProperty -Path $base -Name Enable -ErrorAction SilentlyContinue).Enable
            if ($val -eq 1) { $enabled = $true }
            $detail = "Office $v, Cached Mode Enable = $val"
            break
        }
    }
    if (-not $found) {
        Write-Check 'Cached Exchange Mode' $true 'No cached-mode registry keys (online-only or non-Exchange)'
    } else {
        Write-Check 'Cached Exchange Mode' $enabled $detail
    }
} catch {
    Write-Check 'Cached Exchange Mode' $true $_.Exception.Message
}

# 58. Primary SMTP address of the current mailbox (identity for task attribution)
try {
    $ol = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns = $ol.GetNamespace('MAPI')
    $smtp = ''
    try {
        $cu = $ns.CurrentUser
        $ae = $cu.AddressEntry
        if ($ae -and $ae.Type -eq 'EX') {
            $eu = $ae.GetExchangeUser()
            if ($eu) { $smtp = $eu.PrimarySmtpAddress; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($eu) }
        } elseif ($ae) {
            $smtp = $ae.Address
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ae)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($cu)
    } catch {}
    foreach ($o in @($ns, $ol)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $ok = -not [string]::IsNullOrWhiteSpace($smtp)
    $detail = if ($ok) { "Primary SMTP: $smtp" } else { 'Could not resolve primary SMTP' }
    Write-Check 'Primary SMTP address resolvable' $ok $detail
} catch {
    Write-Check 'Primary SMTP address resolvable' $false $_.Exception.Message
}

# 59. Mailbox type (Exchange / Exchange Online / IMAP / POP / PST)
#     Detected from Stores.ExchangeStoreType where available.
try {
    $ol = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns = $ol.GetNamespace('MAPI')
    $stores = $ns.Stores
    $types = @()
    foreach ($s in $stores) {
        $kind = 'Unknown'
        try {
            switch ($s.ExchangeStoreType) {
                0 { $kind = 'PrimaryExchange' }
                1 { $kind = 'DelegateExchange' }
                2 { $kind = 'PublicFolder' }
                3 { $kind = 'NotExchange' }
                default { $kind = "Type$($s.ExchangeStoreType)" }
            }
        } catch { $kind = 'NotExchange' }
        $types += "$($s.DisplayName)=$kind"
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($s)
    }
    foreach ($o in @($stores, $ns, $ol)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Check 'Mailbox store types' $true ($types -join '; ')
} catch {
    Write-Check 'Mailbox store types' $false $_.Exception.Message
}

# 60. Unread inbox count (informational — sync volume sizing)
try {
    $ol  = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns  = $ol.GetNamespace('MAPI')
    $inb = $ns.GetDefaultFolder(6)
    $unread = $inb.UnReadItemCount
    $total  = $inb.Items.Count
    foreach ($o in @($inb, $ns, $ol)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Write-Check 'Inbox size + unread count' $true "$unread unread / $total total"
} catch {
    Write-Check 'Inbox size + unread count' $false $_.Exception.Message
}

# --- Summary + compact result string ---

$total  = $results.Count
$passed = ($results.Values | Where-Object { $_.Ok }).Count
$failed = $total - $passed

if (-not $Quiet) {
    Write-Host ""
    Write-Host ("Result: {0}/{1} checks passed" -f $passed, $total) -ForegroundColor Cyan
    if ($failed -gt 0) {
        Write-Host ""
        Write-Host "Failed checks:" -ForegroundColor Yellow
        $i = 0
        foreach ($name in $script:CheckOrder) {
            $i++
            if (-not $results[$name].Ok) {
                Write-Host ("  #{0,-3} {1}" -f $i, $name) -ForegroundColor Red
            }
        }
    }
}

# Build compact result string.
#   Format: VIGIL:v1:<count>:<hexBitmap>:P<pass>:F<fail>:T<tenantTag>
#   Bitmap: bit 0 (LSB) = check #1, bit N-1 = check #N. 1 = pass, 0 = fail.
#   tenantTag: "none" | "match" | "mismatch" | "unknown" (for quick triage)
$bits = [System.Numerics.BigInteger]::Zero
for ($idx = 0; $idx -lt $script:CheckOrder.Count; $idx++) {
    if ($results[$script:CheckOrder[$idx]].Ok) {
        $bits = $bits -bor ([System.Numerics.BigInteger]::One -shl $idx)
    }
}
# BigInteger.ToString('X') may add a leading 0 for sign-safety; strip it.
$hex = $bits.ToString('X').TrimStart('0')
if (-not $hex) { $hex = '0' }

$tenantTag = 'none'
if ($TenantId) {
    if ($results['Tenant ID matches expected'].Ok) { $tenantTag = 'match' }
    else                                            { $tenantTag = 'mismatch' }
} elseif ($script:DetectedTenantId) {
    $tenantTag = 'detected'
}

$resultString = "VIGIL:v2:{0}:{1}:P{2}:F{3}:T{4}" -f $total, $hex, $passed, $failed, $tenantTag

Write-Host ""
Write-Host "Paste this single line back for remote triage:" -ForegroundColor Cyan
Write-Host $resultString -ForegroundColor White

if ($passed -eq $total) { exit 0 } else { exit 1 }
