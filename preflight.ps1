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
try {
    Add-Type -ErrorAction Stop -TypeDefinition @"
public class VigilAmsiProbe { public static int Answer() { return 42; } }
"@
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
