#!/usr/bin/env pwsh
# VIGIL core logic smoke tests.
# Runs on Windows PowerShell 5.1 AND cross-platform pwsh 6+.
# Exercises: task schema, sort modes, filter modes, JSON round-trip,
# null-entry healing, due-date formatting.
#
# Usage:
#   pwsh -File Test-Vigil.ps1
#   powershell -File Test-Vigil.ps1


$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vigilPath = Join-Path $scriptDir 'VIGIL.ps1'

if (-not (Test-Path $vigilPath)) {
    Write-Host "VIGIL.ps1 not found at $vigilPath" -ForegroundColor Red
    exit 1
}

# Dot-source VIGIL in NoUI mode - defines all core functions, returns before UI
. $vigilPath -NoUI

$script:Passed = 0
$script:Failed = 0
$script:TestLog = @()

function Assert-That {
    param([bool]$Condition, [string]$Label, [object]$Expected = $null, [object]$Actual = $null)
    if ($Condition) {
        Write-Host "  PASS  $Label" -ForegroundColor Green
        $script:Passed++
    } else {
        Write-Host "  FAIL  $Label" -ForegroundColor Red
        if ($null -ne $Expected -or $null -ne $Actual) {
            Write-Host "        expected: $Expected" -ForegroundColor DarkGray
            Write-Host "        actual:   $Actual" -ForegroundColor DarkGray
        }
        $script:Failed++
    }
}

function Assert-Eq { param($Expected, $Actual, [string]$Label)
    Assert-That (($Expected -eq $Actual) -or ($null -eq $Expected -and $null -eq $Actual)) $Label $Expected $Actual
}

function Assert-True { param([bool]$Cond, [string]$Label)  Assert-That $Cond $Label }

function Assert-Count { param([int]$Expected, $Collection, [string]$Label)
    $arr = @($Collection)
    Assert-That ($arr.Count -eq $Expected) $Label $Expected $arr.Count
}

function Section { param([string]$Name)
    Write-Host ""
    Write-Host "== $Name ==" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "VIGIL Test Suite" -ForegroundColor Cyan
Write-Host ('Version: ' + $script:VigilVersion) -ForegroundColor DarkGray
Write-Host ('Host:    ' + $(if ($script:IsWindowsHost) { 'Windows' } else { 'Linux/macOS (DPAPI passthrough)' })) -ForegroundColor DarkGray

# --------------------------------------------------------------------------
Section 'New-VigilTask schema'
# --------------------------------------------------------------------------

$t = New-VigilTask -Title 'Test task' -Priority 'high'
Assert-Eq 'Test task' $t.title    'title set'
Assert-Eq 'high'      $t.priority 'priority set'
Assert-Eq 'manual'    $t.source   'default source is manual'
Assert-Eq $false      $t.done     'starts not done'
Assert-Eq ''          $t.dueDate  'empty dueDate by default'
Assert-True ($null -ne $t.id -and $t.id.Length -gt 0) 'id is non-empty GUID'
Assert-True ($null -ne $t.createdAt -and $t.createdAt.Length -gt 0) 'createdAt is set'

$t2 = New-VigilTask -Title 'With due' -Priority 'critical' -DueDate ([datetime]'2026-12-25T10:00:00')
Assert-True ($t2.dueDate -match '2026-12-25') 'dueDate set via param'

# --------------------------------------------------------------------------
Section 'Sort-VigilTasks - smart mode'
# --------------------------------------------------------------------------

$tasks = @(
    (New-VigilTask -Title 'low'      -Priority 'low')
    (New-VigilTask -Title 'crit'     -Priority 'critical')
    (New-VigilTask -Title 'normal'   -Priority 'normal')
    (New-VigilTask -Title 'high'     -Priority 'high')
)
$sorted = @(Sort-VigilTasks -tasks $tasks -mode 'smart')
Assert-Eq 'crit'   $sorted[0].title 'critical first'
Assert-Eq 'high'   $sorted[1].title 'high second'
Assert-Eq 'normal' $sorted[2].title 'normal third'
Assert-Eq 'low'    $sorted[3].title 'low last'

# --------------------------------------------------------------------------
Section 'Sort-VigilTasks - due date mode (overdue floats to top)'
# --------------------------------------------------------------------------

$yesterday = (Get-Date).AddDays(-1)
$tomorrow  = (Get-Date).AddDays(1)
$nextWeek  = (Get-Date).AddDays(7)

$tasksWithDue = @(
    (New-VigilTask -Title 'next-week' -Priority 'normal' -DueDate $nextWeek)
    (New-VigilTask -Title 'tomorrow'  -Priority 'normal' -DueDate $tomorrow)
    (New-VigilTask -Title 'overdue'   -Priority 'normal' -DueDate $yesterday)
)
$sortedDue = @(Sort-VigilTasks -tasks $tasksWithDue -mode 'smart')
Assert-Eq 'overdue'   $sortedDue[0].title 'overdue first in smart mode'
Assert-Eq 'tomorrow'  $sortedDue[1].title 'tomorrow second'
Assert-Eq 'next-week' $sortedDue[2].title 'next-week last'

# --------------------------------------------------------------------------
Section 'Sort-VigilTasks - newest first'
# --------------------------------------------------------------------------

Start-Sleep -Milliseconds 20
$t_old = New-VigilTask -Title 'old' -Priority 'normal'
Start-Sleep -Milliseconds 20
$t_mid = New-VigilTask -Title 'mid' -Priority 'normal'
Start-Sleep -Milliseconds 20
$t_new = New-VigilTask -Title 'new' -Priority 'normal'
$sortedNew = @(Sort-VigilTasks -tasks @($t_old, $t_mid, $t_new) -mode 'added')
Assert-Eq 'new' $sortedNew[0].title 'newest first in added mode'
Assert-Eq 'old' $sortedNew[2].title 'oldest last in added mode'

# --------------------------------------------------------------------------
Section 'Sort-VigilTasks - null priority safety'
# --------------------------------------------------------------------------

$badTask = New-VigilTask -Title 'bad' -Priority 'normal'
$badTask.priority = $null
$sortedBad = @(Sort-VigilTasks -tasks @((New-VigilTask -Title 'ok' -Priority 'high'), $badTask) -mode 'smart')
Assert-Eq 2 $sortedBad.Count 'null priority does not throw'

# --------------------------------------------------------------------------
Section 'Filter-VigilTasks'
# --------------------------------------------------------------------------

$filterFixture = @(
    (New-VigilTask -Title 'm1' -Priority 'high'     -Source 'manual')
    (New-VigilTask -Title 'm2' -Priority 'low'      -Source 'manual')
    (New-VigilTask -Title 'c1' -Priority 'critical' -Source 'outlook-cal')
    (New-VigilTask -Title 'f1' -Priority 'normal'   -Source 'outlook-flag')
    (New-VigilTask -Title 't1' -Priority 'high'     -Source 'outlook-task')
)

Assert-Count 5 (Filter-VigilTasks -tasks $filterFixture -mode 'all')    'all keeps everything'
Assert-Count 2 (Filter-VigilTasks -tasks $filterFixture -mode 'manual') 'manual = 2'
Assert-Count 3 (Filter-VigilTasks -tasks $filterFixture -mode 'outlook') 'outlook = 3'
Assert-Count 3 (Filter-VigilTasks -tasks $filterFixture -mode 'urgent')  'urgent (high+critical) = 3'

# --------------------------------------------------------------------------
Section 'Format-DueLabel'
# --------------------------------------------------------------------------

$todayAt5 = (Get-Date).Date.AddHours(17)
$label1 = Format-DueLabel $todayAt5.ToString('o')
Assert-True ($label1 -like 'Today*') "today-at-5pm -> $label1"

$tomorrowAt9 = (Get-Date).Date.AddDays(1).AddHours(9)
$label2 = Format-DueLabel $tomorrowAt9.ToString('o')
Assert-True ($label2 -like 'Tomorrow*') "tomorrow-9am -> $label2"

$past = (Get-Date).AddHours(-2)
$label3 = Format-DueLabel $past.ToString('o')
Assert-True ($label3 -like 'Overdue*') "past date -> $label3"

$label4 = Format-DueLabel ''
Assert-Eq '' $label4 'empty input -> empty string'

# --------------------------------------------------------------------------
Section 'JSON round-trip via Load/Save (uses temp path)'
# --------------------------------------------------------------------------

$tmpBase = [IO.Path]::GetTempPath()
$runId = [guid]::NewGuid().ToString('N')
$script:TasksPath    = Join-Path $tmpBase ("vigil-test-$runId-tasks.json")
$script:BackupPath   = Join-Path $tmpBase ("vigil-test-$runId-tasks.backup.json")
$script:TmpPath      = Join-Path $tmpBase ("vigil-test-$runId-tasks.tmp.json")
$script:SettingsPath = Join-Path $tmpBase ("vigil-test-$runId-settings.json")

$writeMe = @(
    (New-VigilTask -Title 'round1' -Priority 'high')
    (New-VigilTask -Title 'round2' -Priority 'normal' -DueDate (Get-Date).AddHours(3))
    (New-VigilTask -Title 'round3' -Priority 'low')
)

Save-VigilTasks $writeMe
Assert-True (Test-Path $script:TasksPath) 'save created file on disk'

$loaded = @(Load-VigilTasks)
Assert-Count 3 $loaded 'round-trip preserves count'
Assert-Eq 'round1' $loaded[0].title    'round-trip preserves title[0]'
Assert-Eq 'high'   $loaded[0].priority 'round-trip preserves priority[0]'
Assert-Eq 'low'    $loaded[2].priority 'round-trip preserves priority[2]'

# --------------------------------------------------------------------------
Section 'Null-entry healing (simulates build25-era corruption)'
# --------------------------------------------------------------------------

# Build an array with a null element, save it, verify our foreach-filter works
$corruptArray = @($null, (New-VigilTask -Title 'survivor' -Priority 'normal'))
$cleaned = @()
foreach ($t in $corruptArray) { if ($null -ne $t -and $t.id) { $cleaned += $t } }
Assert-Count 1 $cleaned 'null entries filtered out'
Assert-Eq 'survivor' $cleaned[0].title 'valid entry survived'

# --------------------------------------------------------------------------
Section 'Settings merge-with-defaults'
# --------------------------------------------------------------------------

# Remove any test settings file so we get defaults
if (Test-Path $script:SettingsPath) { Remove-Item $script:SettingsPath -Force }
$settings = Load-VigilSettings
Assert-True ($null -ne $settings.sortMode) 'default settings has sortMode'
Assert-Eq 'smart' $settings.sortMode 'default sortMode is smart'
Assert-True ($null -ne $settings.activeFilter) 'default settings has activeFilter'

# Mutate and save
$settings.sortMode = 'priority'
Save-VigilSettings $settings
$settings2 = Load-VigilSettings
Assert-Eq 'priority' $settings2.sortMode 'settings persists across load'

# Verify defaults still merge in if file is missing keys (manual corruption)
$partial = @{ posX = 9999 } | ConvertTo-Json
[IO.File]::WriteAllText($script:SettingsPath, $partial)
$settings3 = Load-VigilSettings
Assert-Eq 9999 $settings3.posX 'partial settings keeps posX from file'
Assert-Eq 'smart' $settings3.sortMode 'missing sortMode defaults from merge'

# --------------------------------------------------------------------------
Section 'Update-VigilTask - mutate individual fields'
# --------------------------------------------------------------------------

$addMe = New-VigilTask -Title 'original' -Priority 'low'
Save-VigilTasks @($addMe)

$ok = Update-VigilTask -Id $addMe.id -Title 'updated'
Assert-True $ok 'Update returns true when task found'
$reloaded = Get-VigilTaskById $addMe.id
Assert-Eq 'updated' $reloaded.title 'title updated via Update-VigilTask'

[void](Update-VigilTask -Id $addMe.id -Priority 'critical')
$reloaded = Get-VigilTaskById $addMe.id
Assert-Eq 'critical' $reloaded.priority 'priority updated'

[void](Update-VigilTask -Id $addMe.id -Priority 'garbage')
$reloaded = Get-VigilTaskById $addMe.id
Assert-Eq 'critical' $reloaded.priority 'invalid priority rejected'

[void](Update-VigilTask -Id $addMe.id -Notes 'remember to test')
$reloaded = Get-VigilTaskById $addMe.id
Assert-Eq 'remember to test' $reloaded.notes 'notes updated'

[void](Update-VigilTask -Id $addMe.id -Done $true)
$reloaded = Get-VigilTaskById $addMe.id
Assert-Eq $true $reloaded.done 'done flag set'
Assert-True ($reloaded.doneAt -ne '') 'doneAt timestamp set'

$missing = Update-VigilTask -Id 'nonexistent' -Title 'nope'
Assert-Eq $false $missing 'Update returns false for missing id'

# --------------------------------------------------------------------------
Section 'Add-VigilTask / Remove-VigilTask'
# --------------------------------------------------------------------------

Save-VigilTasks @()
$t1 = Add-VigilTask (New-VigilTask -Title 'add1' -Priority 'normal')
$t2 = Add-VigilTask (New-VigilTask -Title 'add2' -Priority 'high')
$all = @(Load-VigilTasks)
Assert-Count 2 $all 'Add-VigilTask appends to disk'

$removed = Remove-VigilTask $t1.id
Assert-True $removed 'Remove returns true when removed'
$all = @(Load-VigilTasks)
Assert-Count 1 $all 'Remove-VigilTask removes by id'
Assert-Eq 'add2' $all[0].title 'correct task remains'

$removed2 = Remove-VigilTask 'nope'
Assert-Eq $false $removed2 'Remove returns false when not found'

# --------------------------------------------------------------------------
Section 'Get-VigilOverdueTasks'
# --------------------------------------------------------------------------

$tnow = Get-Date
$overFix = @(
    (New-VigilTask -Title 'past'   -Priority 'normal' -DueDate $tnow.AddHours(-2))
    (New-VigilTask -Title 'future' -Priority 'normal' -DueDate $tnow.AddHours(4))
    (New-VigilTask -Title 'no-due' -Priority 'normal')
)
$doneOverdue = New-VigilTask -Title 'past-done' -Priority 'normal' -DueDate $tnow.AddHours(-3)
$doneOverdue.done = $true
$overFix += $doneOverdue

$overdueList = @(Get-VigilOverdueTasks -tasks $overFix)
Assert-Count 1 $overdueList 'Overdue returns only active past-due'
Assert-Eq 'past' $overdueList[0].title 'correct overdue task'

# --------------------------------------------------------------------------
Section 'Export-VigilMarkdown'
# --------------------------------------------------------------------------

$exportFixture = @(
    (New-VigilTask -Title 'active-high' -Priority 'high')
    (New-VigilTask -Title 'overdue-it'  -Priority 'critical' -DueDate $tnow.AddHours(-1))
)
$dTask = New-VigilTask -Title 'done-task' -Priority 'normal'
$dTask.done = $true
$exportFixture += $dTask

$md = Export-VigilMarkdown -tasks $exportFixture
Assert-True ($md -like '*# VIGIL Tasks*')     'header present'
Assert-True ($md -like '*## Overdue*')        'overdue section present'
Assert-True ($md -like '*## Active*')         'active section present'
Assert-True ($md -like '*## Completed*')      'completed section present'
Assert-True ($md -like '*overdue-it*')        'overdue task listed'
Assert-True ($md -like '*active-high*')       'active task listed'
Assert-True ($md.Contains('- [x] done-task'))   'done task uses [x] checkbox'
Assert-True ($md.Contains('- [ ] active-high')) 'active task uses [ ] checkbox'

# --------------------------------------------------------------------------
Section 'Edge cases - empty inputs'
# --------------------------------------------------------------------------

$emptySort = @(Sort-VigilTasks -tasks @() -mode 'smart')
Assert-Count 0 $emptySort 'Sort on empty returns empty'

$emptyFilter = @(Filter-VigilTasks -tasks @() -mode 'urgent')
Assert-Count 0 $emptyFilter 'Filter on empty returns empty'

$emptyOverdue = @(Get-VigilOverdueTasks -tasks @())
Assert-Count 0 $emptyOverdue 'Overdue on empty returns empty'

$emptyMd = Export-VigilMarkdown -tasks @()
Assert-True ($emptyMd.Contains('# VIGIL Tasks')) 'Export on empty still has header'
Assert-True (-not $emptyMd.Contains('## Overdue')) 'Export on empty omits overdue section'
Assert-True (-not $emptyMd.Contains('## Active'))  'Export on empty omits active section'

Save-VigilTasks @()
$loadedEmpty = @(Load-VigilTasks)
Assert-Count 0 $loadedEmpty 'Round-trip empty array'

# --------------------------------------------------------------------------
Section 'Edge cases - malformed data resilience'
# --------------------------------------------------------------------------

$garbage = New-VigilTask -Title 'garbage due' -Priority 'normal'
$garbage.dueDate = 'not-a-date'
$mixedTasks = @((New-VigilTask -Title 'good' -Priority 'high'), $garbage)

$sortedGarbage = @(Sort-VigilTasks -tasks $mixedTasks -mode 'smart')
Assert-Count 2 $sortedGarbage 'Sort tolerates unparseable dueDate'

$overdueGarbage = @(Get-VigilOverdueTasks -tasks $mixedTasks)
Assert-Count 0 $overdueGarbage 'Unparseable dueDate not flagged as overdue'

$labelGarbage = Format-DueLabel 'not-a-date'
Assert-Eq '' $labelGarbage 'Format-DueLabel returns empty on garbage'

$labelNull = Format-DueLabel $null
Assert-Eq '' $labelNull 'Format-DueLabel returns empty on null'

# --------------------------------------------------------------------------
Section 'Edge cases - Format-DueLabel boundaries'
# --------------------------------------------------------------------------

$edgeNow = Get-Date
$thisWeek = $edgeNow.AddDays(3).Date.AddHours(10)   # 3 days from now
$label = Format-DueLabel $thisWeek.ToString('o')
Assert-True ($label -notlike 'Overdue*') 'Within-week date not marked overdue'
Assert-True ($label -notlike 'Today*')   'Within-week != Today'

$farFuture = $edgeNow.AddDays(30).Date
$labelFar = Format-DueLabel $farFuture.ToString('o')
Assert-True ($labelFar.Length -gt 0) 'Far-future date formats to non-empty string'
Assert-True ($labelFar -notlike 'Overdue*') 'Far-future not overdue'

# Task due today but in the past -> Overdue HH:MM (not Today)
$pastToday = $edgeNow.Date.AddHours(6)  # 6am today
if ($pastToday -lt $edgeNow) {
    $labelPT = Format-DueLabel $pastToday.ToString('o')
    Assert-True ($labelPT -like 'Overdue*') 'Past-today shows Overdue, not Today'
}

# --------------------------------------------------------------------------
Section 'Edge cases - sort / filter fall-through'
# --------------------------------------------------------------------------

$fixture4 = @(
    (New-VigilTask -Title 'a' -Priority 'normal')
    (New-VigilTask -Title 'b' -Priority 'high')
)
$unknownSort = @(Sort-VigilTasks -tasks $fixture4 -mode 'bogus')
Assert-Count 2 $unknownSort 'Unknown sort mode falls through to smart'
Assert-Eq 'b' $unknownSort[0].title 'Smart fallback: high before normal'

$unknownFilter = @(Filter-VigilTasks -tasks $fixture4 -mode 'bogus')
Assert-Count 2 $unknownFilter 'Unknown filter mode returns all'

$allMode = @(Filter-VigilTasks -tasks $fixture4 -mode 'all')
Assert-Count 2 $allMode 'Explicit "all" returns everything'

# --------------------------------------------------------------------------
Section 'Edge cases - Update-VigilTask no-op + empty values'
# --------------------------------------------------------------------------

Save-VigilTasks @()
$noop = New-VigilTask -Title 'noop' -Priority 'normal'
Save-VigilTasks @($noop)

# Update with no fields specified = just Id + no mutation
$noopResult = Update-VigilTask -Id $noop.id
Assert-True $noopResult 'Update with no fields still returns true for found id'

# Empty title should be rejected (title is required)
[void](Update-VigilTask -Id $noop.id -Title '')
$reloaded = Get-VigilTaskById $noop.id
Assert-Eq 'noop' $reloaded.title 'Empty title rejected - original preserved'

# DueDate can be explicitly cleared
[void](Update-VigilTask -Id $noop.id -DueDate '')
$reloaded = Get-VigilTaskById $noop.id
Assert-Eq '' $reloaded.dueDate 'DueDate cleared via empty string'

# Done = false explicit
[void](Update-VigilTask -Id $noop.id -Done $true)
[void](Update-VigilTask -Id $noop.id -Done $false)
$reloaded = Get-VigilTaskById $noop.id
Assert-Eq $false $reloaded.done 'Done can be toggled back to false'
Assert-Eq '' $reloaded.doneAt  'doneAt cleared when un-done'

# --------------------------------------------------------------------------
Section 'Edge cases - very long title + special chars'
# --------------------------------------------------------------------------

$longTitle = 'x' * 500
$longTask = New-VigilTask -Title $longTitle -Priority 'normal'
Save-VigilTasks @($longTask)
$reloaded = @(Load-VigilTasks)
Assert-Eq $longTitle $reloaded[0].title 'Long title round-trips via JSON'

$special = 'Task with "quotes" and \ backslash and `ticks and $vars'
$specTask = New-VigilTask -Title $special -Priority 'low'
Save-VigilTasks @($specTask)
$reloaded = @(Load-VigilTasks)
Assert-Eq $special $reloaded[0].title 'Special chars round-trip safely'

# --------------------------------------------------------------------------
Section 'Edge cases - settings file corruption'
# --------------------------------------------------------------------------

# Invalid JSON in settings file -> fall back to defaults
[IO.File]::WriteAllText($script:SettingsPath, 'this is { not valid: json')
$recovered = Load-VigilSettings
Assert-Eq 'smart' $recovered.sortMode 'Corrupt settings falls back to default sortMode'
Assert-True ($null -ne $recovered.posX) 'Corrupt settings still returns object'

# Empty settings file
[IO.File]::WriteAllText($script:SettingsPath, '')
$emptyS = Load-VigilSettings
Assert-Eq 'smart' $emptyS.sortMode 'Empty settings file -> defaults'

# Delete settings file entirely
Remove-Item $script:SettingsPath -ErrorAction SilentlyContinue
$goneS = Load-VigilSettings
Assert-Eq 'smart' $goneS.sortMode 'Missing settings file -> defaults'

# --------------------------------------------------------------------------
Section 'Edge cases - overdue sort order (priority within overdue)'
# --------------------------------------------------------------------------

$pastBase = (Get-Date).AddHours(-4)
$overdueMix = @(
    (New-VigilTask -Title 'overdue-low'      -Priority 'low'      -DueDate $pastBase)
    (New-VigilTask -Title 'overdue-critical' -Priority 'critical' -DueDate $pastBase)
    (New-VigilTask -Title 'future-critical'  -Priority 'critical' -DueDate (Get-Date).AddHours(4))
)
$sortedMix = @(Sort-VigilTasks -tasks $overdueMix -mode 'smart')
Assert-Eq 'overdue-critical' $sortedMix[0].title 'Overdue beats future even at same priority'
Assert-Eq 'overdue-low'      $sortedMix[1].title 'Within overdue, critical still beats low'
Assert-Eq 'future-critical'  $sortedMix[2].title 'Future items come last'

# --------------------------------------------------------------------------
Section 'Edge cases - null-entry healing in Sort/Filter/Overdue'
# --------------------------------------------------------------------------

$withNulls = @($null, (New-VigilTask -Title 'survivor' -Priority 'normal'), $null)
# Sort skips null entries (no throw)
$sortedN = @(Sort-VigilTasks -tasks $withNulls -mode 'smart')
Assert-True ($sortedN.Count -ge 1) 'Sort with null entries does not throw'

# Filter with null entries
$filteredN = @(Filter-VigilTasks -tasks $withNulls -mode 'all')
Assert-True ($filteredN.Count -ge 1) 'Filter with null entries does not throw'

# Overdue with null entries
$overdueN = @(Get-VigilOverdueTasks -tasks $withNulls)
Assert-Count 0 $overdueN 'Overdue ignores null entries'

# Export with null entries
$mdN = Export-VigilMarkdown -tasks $withNulls
Assert-True ($mdN.Contains('survivor')) 'Export skips null entries but keeps real ones'

# --------------------------------------------------------------------------
# Cleanup
# --------------------------------------------------------------------------

Remove-Item -ErrorAction SilentlyContinue $script:TasksPath, $script:BackupPath, $script:TmpPath, $script:SettingsPath

Write-Host ""
Write-Host ('=' * 48) -ForegroundColor DarkGray
$color = if ($script:Failed -eq 0) { 'Green' } else { 'Red' }
Write-Host ("Passed: {0}   Failed: {1}" -f $script:Passed, $script:Failed) -ForegroundColor $color
Write-Host ('=' * 48) -ForegroundColor DarkGray
if ($script:Failed -gt 0) { exit 1 } else { exit 0 }
