# VIGIL — Personal Task Command Center
# Phase 1: widget + data layer + Apple-styled UI (reduce-motion variant)
#
# Environment requirements verified by preflight.ps1 (schema v2, 58/60):
#   - #31 BitLocker OFF  → tasks.json is DPAPI-wrapped (CurrentUser scope)
#   - #45 MinAnimate OFF → all WPF Storyboards removed (static UI)
#
# Usage: powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File .\VIGIL.ps1

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.Windows.Forms

# --- Single instance -------------------------------------------------------
$script:Mutex = New-Object System.Threading.Mutex($false, 'Global\VIGIL_TaskTracker')
if (-not $script:Mutex.WaitOne(0, $false)) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class VigilWin32 {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@ -ErrorAction SilentlyContinue
    $h = [VigilWin32]::FindWindow($null, 'VIGIL')
    if ($h -ne [IntPtr]::Zero) {
        [VigilWin32]::ShowWindow($h, 9) | Out-Null
        [VigilWin32]::SetForegroundWindow($h) | Out-Null
    }
    exit 0
}

# --- Paths -----------------------------------------------------------------
$script:VigilDir     = Join-Path $env:USERPROFILE '.vigil'
$script:TasksPath    = Join-Path $script:VigilDir 'tasks.json'
$script:BackupPath   = Join-Path $script:VigilDir 'tasks.backup.json'
$script:TmpPath      = Join-Path $script:VigilDir 'tasks.tmp.json'
$script:SettingsPath = Join-Path $script:VigilDir 'settings.json'
$script:LogPath      = Join-Path $script:VigilDir 'vigil.log'
if (-not (Test-Path $script:VigilDir)) { New-Item -ItemType Directory -Path $script:VigilDir -Force | Out-Null }

function Write-VigilLog([string]$msg) {
    $ts = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'
    "$ts  $msg" | Add-Content -Path $script:LogPath -Encoding UTF8
}

# Log rotation: keep last 500 lines
try {
    if ((Test-Path $script:LogPath) -and (Get-Content $script:LogPath).Count -gt 500) {
        $tail = Get-Content $script:LogPath -Tail 500
        [System.IO.File]::WriteAllLines($script:LogPath, $tail, [System.Text.UTF8Encoding]::new($false))
    }
} catch { }

# --- DPAPI wrap / unwrap (required because BitLocker is OFF, preflight #31)
function Protect-VigilBytes([byte[]]$plain) {
    [System.Security.Cryptography.ProtectedData]::Protect(
        $plain, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
}
function Unprotect-VigilBytes([byte[]]$cipher) {
    [System.Security.Cryptography.ProtectedData]::Unprotect(
        $cipher, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
}

# --- Data layer ------------------------------------------------------------
function New-VigilTask {
    param(
        [string]$Title,
        [ValidateSet('low','normal','high','critical')][string]$Priority = 'normal',
        [datetime]$DueDate = [datetime]::MinValue,
        [string]$Source = 'manual',
        [string]$Notes = ''
    )
    [pscustomobject]@{
        id        = [guid]::NewGuid().ToString()
        title     = $Title
        priority  = $Priority
        dueDate   = $(if ($DueDate -gt [datetime]::MinValue) { $DueDate.ToString('o') } else { '' })
        source    = $Source
        category  = ''
        notes     = $Notes
        done      = $false
        createdAt = (Get-Date).ToString('o')
        doneAt    = ''
    }
}

function Load-VigilTasks {
    if (-not (Test-Path $script:TasksPath)) { return @() }
    foreach ($path in @($script:TasksPath, $script:BackupPath)) {
        try {
            $cipher = [System.IO.File]::ReadAllBytes($path)
            if ($cipher.Length -eq 0) { return @() }
            $plain = Unprotect-VigilBytes $cipher
            $json = [System.Text.Encoding]::UTF8.GetString($plain)
            if ([string]::IsNullOrWhiteSpace($json)) { return @() }
            return @(ConvertFrom-Json $json)
        } catch {
            Write-VigilLog "Load failed for $path : $($_.Exception.Message)"
            continue
        }
    }
    return @()
}

function Save-VigilTasks([object[]]$tasks) {
    $json = ConvertTo-Json -InputObject @($tasks) -Depth 6 -Compress
    $plain = [System.Text.Encoding]::UTF8.GetBytes($json)
    $cipher = Protect-VigilBytes $plain
    [System.IO.File]::WriteAllBytes($script:TmpPath, $cipher)
    if (Test-Path $script:TasksPath) {
        [System.IO.File]::Replace($script:TmpPath, $script:TasksPath, $script:BackupPath)
    } else {
        [System.IO.File]::Move($script:TmpPath, $script:TasksPath)
    }
}

# --- Settings --------------------------------------------------------------
function Load-VigilSettings {
    $default = [pscustomobject]@{
        posX = 1200; posY = 400; collapsed = $false; showCompleted = $false
        outlookSync = $false; syncIntervalMin = 15; opacity = 1.0
        lastSyncTime = ''; activeFilter = 'all'; autoStartInstalled = $false
    }
    if (-not (Test-Path $script:SettingsPath)) { return $default }
    try {
        $raw = [System.IO.File]::ReadAllText($script:SettingsPath, [System.Text.UTF8Encoding]::new($false))
        return (ConvertFrom-Json $raw)
    } catch {
        Write-VigilLog "Settings load failed: $($_.Exception.Message)"
        return $default
    }
}

function Save-VigilSettings($settings) {
    $json = ConvertTo-Json -InputObject $settings -Depth 4
    $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($json)
    $tmp = "$script:SettingsPath.tmp"
    [System.IO.File]::WriteAllBytes($tmp, $bytes)
    if (Test-Path $script:SettingsPath) {
        [System.IO.File]::Replace($tmp, $script:SettingsPath, $null)
    } else {
        [System.IO.File]::Move($tmp, $script:SettingsPath)
    }
}

# --- Sort + priority helpers -----------------------------------------------
$script:PriorityRank = @{ critical = 0; high = 1; normal = 2; low = 3 }

function Sort-VigilTasks([object[]]$tasks) {
    $now = Get-Date
    $tasks | Sort-Object @{
        Expression = {
            $due = if ($_.dueDate) { [datetime]::Parse($_.dueDate) } else { [datetime]::MaxValue }
            $overdue = ($due -lt $now -and -not $_.done)
            if ($overdue) { 0 } else { 1 }
        }
    }, @{
        Expression = { $script:PriorityRank[$_.priority] }
    }, @{
        Expression = {
            if ($_.dueDate) { [datetime]::Parse($_.dueDate) } else { [datetime]::MaxValue }
        }
    }
}

function Format-DueLabel([string]$iso) {
    if ([string]::IsNullOrWhiteSpace($iso)) { return '' }
    try {
        $d = [datetime]::Parse($iso)
        $now = Get-Date
        $today = $now.Date
        $tomorrow = $today.AddDays(1)
        if ($d.Date -eq $today)    { return "Today $($d.ToString('h:mm tt'))" }
        if ($d.Date -eq $tomorrow) { return "Tomorrow $($d.ToString('h:mm tt'))" }
        if ($d -lt $now)           { return "Overdue $($d.ToString('MMM d'))" }
        if (($d - $now).Days -lt 7) { return $d.ToString('dddd h:mm tt') }
        return $d.ToString('MMM d')
    } catch { return '' }
}

# --- XAML (Apple skin, static, no storyboards) -----------------------------
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="VIGIL"
        Width="340" Height="460"
        WindowStyle="None" ResizeMode="NoResize"
        AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False"
        TextOptions.TextFormattingMode="Ideal"
        TextOptions.TextRenderingMode="ClearType"
        UseLayoutRounding="True"
        FontFamily="Segoe UI Variable Text, Segoe UI">
  <Window.Resources>
    <SolidColorBrush x:Key="SurfaceBase"   Color="#000000"/>
    <SolidColorBrush x:Key="SurfaceGlass"  Color="#CC000000"/>
    <SolidColorBrush x:Key="SurfaceCard"   Color="#1D1D1F"/>
    <SolidColorBrush x:Key="SurfaceCardHi" Color="#272729"/>
    <SolidColorBrush x:Key="Divider"       Color="#14FFFFFF"/>
    <SolidColorBrush x:Key="TextPrimary"   Color="#FFFFFF"/>
    <SolidColorBrush x:Key="TextSecondary" Color="#8EFFFFFF"/>
    <SolidColorBrush x:Key="TextTertiary"  Color="#52FFFFFF"/>
    <SolidColorBrush x:Key="Accent"        Color="#2997FF"/>
    <SolidColorBrush x:Key="Destructive"   Color="#FF453A"/>
    <SolidColorBrush x:Key="Success"       Color="#30D158"/>

    <Style x:Key="TitleBarButton" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Width" Value="28"/>
      <Setter Property="Height" Value="28"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="5">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#1AFFFFFF"/>
          <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <Border CornerRadius="12" Background="{StaticResource SurfaceBase}">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="30" ShadowDepth="5" Opacity="0.5"/>
    </Border.Effect>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="44"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="32"/>
      </Grid.RowDefinitions>

      <!-- Title bar -->
      <Border Grid.Row="0" Background="{StaticResource SurfaceGlass}"
              CornerRadius="12,12,0,0" x:Name="TitleBar">
        <Border BorderBrush="{StaticResource Divider}" BorderThickness="0,0,0,1">
          <Grid Margin="14,0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="VIGIL" FontSize="17" FontWeight="SemiBold"
                       Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
            <Border Grid.Column="1" x:Name="CountBadge" Background="{StaticResource SurfaceCard}"
                    CornerRadius="10" Padding="8,2" HorizontalAlignment="Left" Margin="10,0,0,0"
                    VerticalAlignment="Center">
              <TextBlock x:Name="CountText" Text="0" FontSize="11" FontWeight="SemiBold"
                         Foreground="{StaticResource TextPrimary}"/>
            </Border>
            <StackPanel Grid.Column="2" Orientation="Horizontal">
              <Button x:Name="BtnCollapse" Content="—" Style="{StaticResource TitleBarButton}"/>
              <Button x:Name="BtnClose"    Content="✕" Style="{StaticResource TitleBarButton}"/>
            </StackPanel>
          </Grid>
        </Border>
      </Border>

      <!-- Task list -->
      <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Disabled" Padding="0,4">
        <ItemsControl x:Name="TaskList">
          <ItemsControl.ItemsPanel>
            <ItemsPanelTemplate>
              <StackPanel/>
            </ItemsPanelTemplate>
          </ItemsControl.ItemsPanel>
        </ItemsControl>
      </ScrollViewer>

      <!-- Inline add -->
      <Border Grid.Row="2" BorderBrush="{StaticResource Divider}" BorderThickness="0,1,0,0"
              Padding="14,10">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBox x:Name="AddInput" Grid.Column="0"
                   Background="{StaticResource SurfaceCard}" Foreground="{StaticResource TextPrimary}"
                   BorderThickness="0" Padding="10,8" FontSize="13" CaretBrush="{StaticResource Accent}">
            <TextBox.Resources>
              <Style TargetType="Border">
                <Setter Property="CornerRadius" Value="8"/>
              </Style>
            </TextBox.Resources>
          </TextBox>
          <Button x:Name="BtnAdd" Grid.Column="1" Content="Add" Margin="8,0,0,0"
                  Background="{StaticResource Accent}" Foreground="White" BorderThickness="0"
                  Padding="14,8" FontSize="13" FontWeight="Medium" Cursor="Hand">
            <Button.Resources>
              <Style TargetType="Border">
                <Setter Property="CornerRadius" Value="8"/>
              </Style>
            </Button.Resources>
          </Button>
        </Grid>
      </Border>

      <!-- Status bar -->
      <Border Grid.Row="3" BorderBrush="{StaticResource Divider}" BorderThickness="0,1,0,0">
        <Grid Margin="14,0">
          <TextBlock x:Name="StatusLeft" Text="" FontSize="11"
                     Foreground="{StaticResource TextTertiary}"
                     VerticalAlignment="Center" HorizontalAlignment="Left"/>
          <TextBlock x:Name="StatusRight" Text="" FontSize="11"
                     Foreground="{StaticResource TextTertiary}"
                     VerticalAlignment="Center" HorizontalAlignment="Right"/>
        </Grid>
      </Border>
    </Grid>
  </Border>
</Window>
'@

# --- Load XAML -------------------------------------------------------------
[xml]$xml = $xaml
$reader = New-Object System.Xml.XmlNodeReader $xml
$window = [Windows.Markup.XamlReader]::Load($reader)

$TitleBar    = $window.FindName('TitleBar')
$BtnCollapse = $window.FindName('BtnCollapse')
$BtnClose    = $window.FindName('BtnClose')
$TaskList    = $window.FindName('TaskList')
$AddInput    = $window.FindName('AddInput')
$BtnAdd      = $window.FindName('BtnAdd')
$CountText   = $window.FindName('CountText')
$CountBadge  = $window.FindName('CountBadge')
$StatusLeft  = $window.FindName('StatusLeft')
$StatusRight = $window.FindName('StatusRight')

# --- State -----------------------------------------------------------------
$script:Tasks    = @(Load-VigilTasks)
$script:Settings = Load-VigilSettings

# Restore position (clamped to working area)
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$window.Left = [math]::Max($wa.X, [math]::Min($script:Settings.posX, $wa.Right  - 340))
$window.Top  = [math]::Max($wa.Y, [math]::Min($script:Settings.posY, $wa.Bottom - 460))

# --- Rendering -------------------------------------------------------------
function Build-TaskCard($task) {
    $border = New-Object System.Windows.Controls.Border
    $border.Margin = New-Object System.Windows.Thickness(10,4,10,4)
    $border.Padding = New-Object System.Windows.Thickness(12,10,12,10)
    $border.CornerRadius = New-Object System.Windows.CornerRadius(8)
    $border.Background = [System.Windows.Media.Brushes]::Transparent
    $border.Cursor = [System.Windows.Input.Cursors]::Hand

    $isOverdue = $false
    if ($task.dueDate) {
        try { $isOverdue = ([datetime]::Parse($task.dueDate) -lt (Get-Date)) -and -not $task.done } catch {}
    }
    if ($task.priority -eq 'critical' -or $isOverdue) {
        $border.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255,69,58))
        $border.BorderThickness = New-Object System.Windows.Thickness(2,0,0,0)
    }

    $grid = New-Object System.Windows.Controls.Grid
    $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::Auto
    $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = New-Object System.Windows.GridLength(1, 'Star')
    $grid.ColumnDefinitions.Add($c1); $grid.ColumnDefinitions.Add($c2)

    $check = New-Object System.Windows.Controls.CheckBox
    $check.IsChecked = $task.done
    $check.Margin = New-Object System.Windows.Thickness(0,2,10,0)
    $check.VerticalAlignment = 'Top'
    $check.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(140,255,255,255))
    [System.Windows.Controls.Grid]::SetColumn($check, 0)

    $stack = New-Object System.Windows.Controls.StackPanel
    [System.Windows.Controls.Grid]::SetColumn($stack, 1)

    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = $task.title
    $title.FontSize = 15
    $title.TextWrapping = 'Wrap'
    switch ($task.priority) {
        'critical' { $title.FontWeight = 'Bold';     $title.Foreground = [System.Windows.Media.Brushes]::White }
        'high'     { $title.FontWeight = 'SemiBold'; $title.Foreground = [System.Windows.Media.Brushes]::White }
        'normal'   { $title.FontWeight = 'Normal';   $title.Foreground = [System.Windows.Media.Brushes]::White }
        'low'      { $title.FontWeight = 'Normal';   $title.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(140,255,255,255)) }
    }
    if ($task.done) {
        $title.TextDecorations = [System.Windows.TextDecorations]::Strikethrough
        $title.Opacity = 0.4
    }
    $stack.Children.Add($title) | Out-Null

    $metaText = @()
    $due = Format-DueLabel $task.dueDate
    if ($due) { $metaText += $due }
    if ($task.source -ne 'manual') { $metaText += $task.source }
    if ($metaText.Count -gt 0) {
        $meta = New-Object System.Windows.Controls.TextBlock
        $meta.Text = ($metaText -join ' · ')
        $meta.FontSize = 12
        $meta.Margin = New-Object System.Windows.Thickness(0,3,0,0)
        if ($isOverdue) {
            $meta.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255,69,58))
        } else {
            $meta.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(140,255,255,255))
        }
        $stack.Children.Add($meta) | Out-Null
    }

    $grid.Children.Add($check) | Out-Null
    $grid.Children.Add($stack) | Out-Null
    $border.Child = $grid

    # Context menu
    $menu = New-Object System.Windows.Controls.ContextMenu
    foreach ($p in 'low','normal','high','critical') {
        $mi = New-Object System.Windows.Controls.MenuItem
        $mi.Header = "Priority: $p"
        $mi.Tag = @{ id = $task.id; action = 'priority'; value = $p }
        $mi.Add_Click({ Handle-ContextAction $this.Tag })
        $menu.Items.Add($mi) | Out-Null
    }
    $sep = New-Object System.Windows.Controls.Separator
    $menu.Items.Add($sep) | Out-Null
    $delItem = New-Object System.Windows.Controls.MenuItem
    $delItem.Header = 'Delete'
    $delItem.Tag = @{ id = $task.id; action = 'delete' }
    $delItem.Add_Click({ Handle-ContextAction $this.Tag })
    $menu.Items.Add($delItem) | Out-Null
    $border.ContextMenu = $menu

    # Check toggle — static, no fade
    $check.Add_Checked({
        param($s, $e)
        $card = $s.Parent.Parent
        $id = $card.Tag
        Toggle-Done $id $true
    }.GetNewClosure())
    $check.Add_Unchecked({
        param($s, $e)
        $card = $s.Parent.Parent
        $id = $card.Tag
        Toggle-Done $id $false
    }.GetNewClosure())

    $border.Tag = $task.id
    return $border
}

function Refresh-Render {
    $TaskList.Items.Clear()
    $visible = if ($script:Settings.showCompleted) { $script:Tasks } else { @($script:Tasks | Where-Object { -not $_.done }) }
    $sorted = Sort-VigilTasks -tasks $visible
    foreach ($t in $sorted) {
        $card = Build-TaskCard $t
        $TaskList.Items.Add($card) | Out-Null
    }
    $active = @($script:Tasks | Where-Object { -not $_.done }).Count
    $CountText.Text = "$active"
    $CountBadge.Visibility = if ($active -gt 0) { 'Visible' } else { 'Collapsed' }
    $StatusRight.Text = "$active active"
    $StatusLeft.Text  = "$(Get-Date -Format 'h:mm tt')"
}

function Toggle-Done([string]$id, [bool]$done) {
    $t = $script:Tasks | Where-Object { $_.id -eq $id }
    if (-not $t) { return }
    $t.done = $done
    $t.doneAt = if ($done) { (Get-Date).ToString('o') } else { '' }
    Save-VigilTasks $script:Tasks
    Refresh-Render
}

function Handle-ContextAction($tag) {
    $id = $tag.id; $action = $tag.action
    if ($action -eq 'delete') {
        $script:Tasks = @($script:Tasks | Where-Object { $_.id -ne $id })
    } elseif ($action -eq 'priority') {
        $t = $script:Tasks | Where-Object { $_.id -eq $id }
        if ($t) { $t.priority = $tag.value }
    }
    Save-VigilTasks $script:Tasks
    Refresh-Render
}

# --- Event wiring ----------------------------------------------------------
$TitleBar.Add_MouseLeftButtonDown({ $window.DragMove() })

$BtnClose.Add_Click({
    $script:Settings.posX = [int]$window.Left
    $script:Settings.posY = [int]$window.Top
    Save-VigilSettings $script:Settings
    $window.Close()
})

$BtnCollapse.Add_Click({
    if ($window.Height -gt 44) {
        $window.Height = 44
    } else {
        $window.Height = 460
    }
})

$AddFn = {
    $txt = $AddInput.Text.Trim()
    if (-not $txt) { return }
    $new = New-VigilTask -Title $txt -Priority 'normal'
    $script:Tasks = @($script:Tasks) + @($new)
    Save-VigilTasks $script:Tasks
    $AddInput.Text = ''
    Refresh-Render
}
$BtnAdd.Add_Click($AddFn)
$AddInput.Add_KeyDown({
    param($s,$e)
    if ($e.Key -eq 'Return') { & $AddFn; $e.Handled = $true }
})

$window.Add_Closing({
    $script:Settings.posX = [int]$window.Left
    $script:Settings.posY = [int]$window.Top
    Save-VigilSettings $script:Settings
    try { $script:Mutex.ReleaseMutex() } catch {}
    $script:Mutex.Dispose()
})

# --- Go --------------------------------------------------------------------
Write-VigilLog "VIGIL started. $($script:Tasks.Count) tasks loaded."
Refresh-Render
[void]$window.ShowDialog()
