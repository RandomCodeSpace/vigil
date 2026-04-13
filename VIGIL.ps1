# VIGIL - Personal Task Command Center
# Phase 1: widget + data layer + Apple-styled UI (reduce-motion variant)
#
# Environment requirements verified by preflight.ps1 (schema v2, 58/60):
#   - #31 BitLocker OFF  -> tasks.json is DPAPI-wrapped (CurrentUser scope)
#   - #45 MinAnimate OFF -> all WPF Storyboards removed (static UI)
#
# Usage: powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File .\VIGIL.ps1

[CmdletBinding()]
param()

# Build stamp - bumped on every commit. Visible in status bar + vigil.log.
# Format: YYYY-MM-DD HH:MM (UTC)  buildN
$script:VigilVersion = '2026-04-13 22:30 UTC  build21 ascii-clean'

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
        $logUtf8 = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllLines($script:LogPath, $tail, $logUtf8)
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
    $dueIso = ''
    if ($DueDate -gt [datetime]::MinValue) { $dueIso = $DueDate.ToString('o') }
    [pscustomobject]@{
        id        = [guid]::NewGuid().ToString()
        title     = $Title
        priority  = $Priority
        dueDate   = $dueIso
        source    = $Source
        category  = ''
        notes     = $Notes
        done      = $false
        createdAt = (Get-Date).ToString('o')
        doneAt    = ''
    }
}

function Load-VigilTasks {
    if (-not (Test-Path $script:TasksPath)) {
        Write-VigilLog "Load: no tasks file at $script:TasksPath"
        return @()
    }
    foreach ($path in @($script:TasksPath, $script:BackupPath)) {
        if (-not (Test-Path $path)) { continue }
        try {
            $cipher = [System.IO.File]::ReadAllBytes($path)
            if ($cipher.Length -eq 0) {
                Write-VigilLog "Load: $path is empty"
                continue
            }
            $plain = Unprotect-VigilBytes $cipher
            $json = [System.Text.Encoding]::UTF8.GetString($plain)
            if ([string]::IsNullOrWhiteSpace($json)) {
                Write-VigilLog "Load: $path decrypted but JSON is empty"
                continue
            }
            $parsed = ConvertFrom-Json $json
            $arr = @($parsed)
            $lmsg = 'Load: read {0} tasks from {1}' -f $arr.Count, $path
            Write-VigilLog $lmsg
            return $arr
        } catch {
            $lfmsg = 'Load FAILED for {0} : {1}' -f $path, $_.Exception.Message
            Write-VigilLog $lfmsg
            continue
        }
    }
    Write-VigilLog 'Load: both primary and backup failed - returning empty'
    return @()
}

function Save-VigilTasks([object[]]$tasks) {
    try {
        $arr = @($tasks)
        $json = ConvertTo-Json -InputObject $arr -Depth 6 -Compress
        if ([string]::IsNullOrWhiteSpace($json)) { $json = '[]' }
        $plain = [System.Text.Encoding]::UTF8.GetBytes($json)
        $cipher = Protect-VigilBytes $plain
        [System.IO.File]::WriteAllBytes($script:TmpPath, $cipher)
        if (Test-Path $script:TasksPath) {
            [System.IO.File]::Replace($script:TmpPath, $script:TasksPath, $script:BackupPath)
        } else {
            [System.IO.File]::Move($script:TmpPath, $script:TasksPath)
        }
        $msg = 'Save: wrote {0} tasks ({1} bytes cipher)' -f $arr.Count, $cipher.Length
        Write-VigilLog $msg
    } catch {
        $emsg = 'Save FAILED: {0}' -f $_.Exception.Message
        Write-VigilLog $emsg
        throw
    }
}

# --- Settings --------------------------------------------------------------
function Load-VigilSettings {
    $default = @{
        posX = 1200; posY = 400; collapsed = $false; showCompleted = $false
        outlookSync = $false; syncIntervalMin = 15; opacity = 1.0
        lastSyncTime = ''; activeFilter = 'all'; autoStartInstalled = $false
        sortMode = 'smart'
    }
    $loaded = @{}
    if (Test-Path $script:SettingsPath) {
        try {
            $utf8 = New-Object System.Text.UTF8Encoding($false)
            $raw = [System.IO.File]::ReadAllText($script:SettingsPath, $utf8)
            $parsed = ConvertFrom-Json $raw
            foreach ($p in $parsed.PSObject.Properties) { $loaded[$p.Name] = $p.Value }
        } catch {
            $smsg = 'Settings load failed: {0}' -f $_.Exception.Message
            Write-VigilLog $smsg
        }
    }
    # Merge loaded over defaults - guarantees every key exists on the returned object
    $merged = @{}
    foreach ($k in $default.Keys) { $merged[$k] = $default[$k] }
    foreach ($k in $loaded.Keys)  { $merged[$k] = $loaded[$k]  }
    [pscustomobject]$merged
}

function Save-VigilSettings($settings) {
    try {
        $json = ConvertTo-Json -InputObject $settings -Depth 4
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $bytes = $utf8.GetBytes($json)
        $tmp = $script:SettingsPath + '.tmp'
        [System.IO.File]::WriteAllBytes($tmp, $bytes)
        if (Test-Path $script:SettingsPath) {
            [System.IO.File]::Delete($script:SettingsPath)
        }
        [System.IO.File]::Move($tmp, $script:SettingsPath)
    } catch {
        $m = 'Save-VigilSettings FAILED: {0}' -f $_.Exception.Message
        Write-VigilLog $m
    }
}

# --- Sort + priority helpers -----------------------------------------------
$script:PriorityRank = @{ critical = 0; high = 1; normal = 2; low = 3 }

function Sort-VigilTasks([object[]]$tasks, [string]$mode = 'smart') {
    $now = Get-Date
    $annotated = foreach ($t in $tasks) {
        $due = [datetime]::MaxValue
        if ($t.dueDate) {
            try { $due = [datetime]::Parse($t.dueDate) } catch {}
        }
        $overdue = 1
        if (($due -lt $now) -and (-not $t.done)) { $overdue = 0 }
        $prank = 9
        $pkey = 'normal'
        if ($t.priority) { $pkey = [string]$t.priority }
        if ($script:PriorityRank.ContainsKey($pkey)) {
            $prank = $script:PriorityRank[$pkey]
        }
        $created = [datetime]::MinValue
        if ($t.createdAt) {
            try { $created = [datetime]::Parse($t.createdAt) } catch {}
        }
        [pscustomobject]@{
            _task     = $t
            _overdue  = $overdue
            _priority = $prank
            _due      = $due
            _created  = $created
        }
    }
    $annotatedArr = @($annotated)
    if ($mode -eq 'added') {
        $sorted = @($annotatedArr | Sort-Object -Property _created -Descending)
    } elseif ($mode -eq 'priority') {
        $sorted = @($annotatedArr | Sort-Object -Property _priority, _due)
    } elseif ($mode -eq 'due') {
        $sorted = @($annotatedArr | Sort-Object -Property _due, _priority)
    } else {
        $sorted = @($annotatedArr | Sort-Object -Property _overdue, _priority, _due)
    }
    $out = @()
    foreach ($row in $sorted) { $out += $row._task }
    return $out
}

function Format-DueLabel([string]$iso) {
    if ([string]::IsNullOrWhiteSpace($iso)) { return '' }
    try {
        $d = [datetime]::Parse($iso)
        $now = Get-Date
        $today = $now.Date
        $tomorrow = $today.AddDays(1)
        if ($d.Date -eq $today)    { return ('Today {0}'    -f $d.ToString('h:mm tt')) }
        if ($d.Date -eq $tomorrow) { return ('Tomorrow {0}' -f $d.ToString('h:mm tt')) }
        if ($d -lt $now)           { return ('Overdue {0}'  -f $d.ToString('MMM d'))  }
        if (($d - $now).Days -lt 7) { return $d.ToString('dddd h:mm tt') }
        return $d.ToString('MMM d')
    } catch { return '' }
}

# --- XAML (custom dark theme, designed from scratch, reduce-motion) --------
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="VIGIL"
        Width="360" SizeToContent="Height"
        WindowStyle="None" ResizeMode="NoResize"
        AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False"
        TextOptions.TextFormattingMode="Ideal"
        TextOptions.TextRenderingMode="Grayscale"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True"
        FontFamily="Segoe UI">
  <Window.Resources>
    <!-- Custom dark theme - designed from scratch, no external design system -->
    <SolidColorBrush x:Key="SurfaceBase"    Color="#0B0E14"/>
    <SolidColorBrush x:Key="SurfaceElev1"   Color="#141922"/>
    <SolidColorBrush x:Key="SurfaceElev2"   Color="#1C2230"/>
    <SolidColorBrush x:Key="SurfaceHover"   Color="#222A3A"/>
    <SolidColorBrush x:Key="Divider"        Color="#1E2431"/>
    <SolidColorBrush x:Key="BorderSubtle"   Color="#2A3245"/>
    <SolidColorBrush x:Key="TextPrimary"    Color="#F0F2F7"/>
    <SolidColorBrush x:Key="TextSecondary"  Color="#8A94A8"/>
    <SolidColorBrush x:Key="TextTertiary"   Color="#555E70"/>
    <SolidColorBrush x:Key="Accent"         Color="#7C5CFF"/>
    <SolidColorBrush x:Key="AccentSoft"     Color="#A996FF"/>
    <SolidColorBrush x:Key="Urgent"         Color="#FF4D6D"/>
    <SolidColorBrush x:Key="UrgentSoft"     Color="#FF7A90"/>
    <SolidColorBrush x:Key="Success"        Color="#2DD4BF"/>

    <!-- Icon button (close, minimize) -->
    <Style x:Key="IconButton" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Width" Value="30"/>
      <Setter Property="Height" Value="30"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bg" Background="{TemplateBinding Background}" CornerRadius="7">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="{StaticResource SurfaceHover}"/>
          <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Close button with red hover -->
    <Style x:Key="CloseButton" TargetType="Button" BasedOn="{StaticResource IconButton}">
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#3A1820"/>
          <Setter Property="Foreground" Value="{StaticResource Urgent}"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Primary filled button (accent violet) -->
    <Style x:Key="PrimaryButton" TargetType="Button">
      <Setter Property="Background" Value="{StaticResource Accent}"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="18,10"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="8"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#9278FF"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Ghost button (used for Sort dropdown trigger) -->
    <Style x:Key="GhostButton" TargetType="Button">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderSubtle}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="7"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="{StaticResource SurfaceElev1}"/>
          <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
          <Setter Property="BorderBrush" Value="{StaticResource Accent}"/>
        </Trigger>
      </Style.Triggers>
    </Style>

  </Window.Resources>

  <Border CornerRadius="16" Background="{StaticResource SurfaceBase}"
          BorderBrush="{StaticResource BorderSubtle}" BorderThickness="1"
          Margin="14">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="36" ShadowDepth="0" Opacity="0.65"/>
    </Border.Effect>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Title bar -->
      <Border Grid.Row="0" x:Name="TitleBar" Background="{StaticResource SurfaceElev1}"
              CornerRadius="16,16,0,0" Padding="16,0" Height="48"
              BorderBrush="{StaticResource Divider}" BorderThickness="0,0,0,1">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>

          <TextBlock Grid.Column="0" Text="VIGIL" FontSize="12" FontWeight="Bold"
                     Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"
                     Margin="0,0,0,0">
            <TextBlock.RenderTransform>
              <TranslateTransform X="0" Y="0"/>
            </TextBlock.RenderTransform>
          </TextBlock>

          <Border Grid.Column="1" x:Name="CountBadge" Background="{StaticResource Accent}"
                  CornerRadius="8" Padding="6,1" Margin="8,0,0,0"
                  VerticalAlignment="Center" MinWidth="18" Height="17">
            <TextBlock x:Name="CountText" Text="0" FontSize="10" FontWeight="Bold"
                       Foreground="#FFFFFF" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>

          <Button Grid.Column="3" x:Name="BtnSort" Content="Smart"
                  Style="{StaticResource GhostButton}" Margin="0,0,8,0"
                  ToolTip="Change sort order"/>

          <StackPanel Grid.Column="4" Orientation="Horizontal" VerticalAlignment="Center">
            <Button x:Name="BtnCollapse" Style="{StaticResource IconButton}" ToolTip="Minimize">
              <Path Data="M0,0 L8,0" Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                    StrokeThickness="1.3" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
            </Button>
            <Button x:Name="BtnClose" Style="{StaticResource CloseButton}" ToolTip="Close" Margin="2,0,0,0">
              <Path Data="M0,0 L8,8 M8,0 L0,8" Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                    StrokeThickness="1.3" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
            </Button>
          </StackPanel>
        </Grid>
      </Border>

      <!-- Task list area -->
      <Border Grid.Row="1" x:Name="TaskArea" Background="{StaticResource SurfaceBase}">
        <ScrollViewer MaxHeight="380" VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled" Padding="10,10,10,6">
          <ItemsControl x:Name="TaskList">
            <ItemsControl.ItemsPanel>
              <ItemsPanelTemplate>
                <StackPanel/>
              </ItemsPanelTemplate>
            </ItemsControl.ItemsPanel>
          </ItemsControl>
        </ScrollViewer>
      </Border>

      <!-- Inline add -->
      <Border Grid.Row="2" x:Name="AddArea" Background="{StaticResource SurfaceElev1}"
              BorderBrush="{StaticResource Divider}" BorderThickness="0,1,0,0" Padding="14,12">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Border Grid.Column="0" Background="{StaticResource SurfaceElev2}"
                  BorderBrush="{StaticResource BorderSubtle}" BorderThickness="1"
                  CornerRadius="8">
            <TextBox x:Name="AddInput" Background="Transparent"
                     Foreground="{StaticResource TextPrimary}"
                     BorderThickness="0" Padding="12,9" FontSize="13"
                     CaretBrush="{StaticResource Accent}"
                     VerticalContentAlignment="Center"/>
          </Border>
          <Button Grid.Column="1" x:Name="BtnAdd" Content="Add"
                  Style="{StaticResource PrimaryButton}" Margin="10,0,0,0"/>
        </Grid>
      </Border>

      <!-- Status bar -->
      <Border Grid.Row="3" x:Name="StatusArea" Background="{StaticResource SurfaceElev1}"
              BorderBrush="{StaticResource Divider}" BorderThickness="0,1,0,0"
              CornerRadius="0,0,16,16" Padding="14,8">
        <Grid>
          <TextBlock x:Name="StatusLeft" Text="" FontSize="10"
                     Foreground="{StaticResource TextTertiary}"
                     VerticalAlignment="Center" HorizontalAlignment="Left"/>
          <TextBlock x:Name="StatusRight" Text="" FontSize="10"
                     Foreground="{StaticResource TextTertiary}"
                     VerticalAlignment="Center" HorizontalAlignment="Right"
                     FontWeight="SemiBold"/>
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
$BtnSort     = $window.FindName('BtnSort')
$TaskArea    = $window.FindName('TaskArea')
$TaskList    = $window.FindName('TaskList')
$AddArea     = $window.FindName('AddArea')
$AddInput    = $window.FindName('AddInput')
$BtnAdd      = $window.FindName('BtnAdd')
$CountText   = $window.FindName('CountText')
$CountBadge  = $window.FindName('CountBadge')
$StatusArea  = $window.FindName('StatusArea')
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
    $border.Margin = New-Object System.Windows.Thickness(0,0,0,6)
    $border.Padding = New-Object System.Windows.Thickness(14,11,12,11)
    $border.CornerRadius = New-Object System.Windows.CornerRadius(10)
    $elev1 = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(20,25,34))
    $borderSubtle = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(42,50,69))
    $border.Background = $elev1
    $border.BorderBrush = $borderSubtle
    $border.BorderThickness = New-Object System.Windows.Thickness(1)
    $border.Cursor = [System.Windows.Input.Cursors]::Hand

    $isOverdue = $false
    if ($task.dueDate) {
        try { $isOverdue = ([datetime]::Parse($task.dueDate) -lt (Get-Date)) -and -not $task.done } catch {}
    }
    if ($task.priority -eq 'critical' -or $isOverdue) {
        $urgent = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255,77,109))
        $border.BorderBrush = $urgent
        $border.BorderThickness = New-Object System.Windows.Thickness(3,1,1,1)
    } elseif ($task.priority -eq 'high') {
        $accentSoft = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(169,150,255))
        $border.BorderBrush = $accentSoft
        $border.BorderThickness = New-Object System.Windows.Thickness(3,1,1,1)
    }

    $grid = New-Object System.Windows.Controls.Grid
    $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::Auto
    $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = New-Object System.Windows.GridLength(1, 'Star')
    $grid.ColumnDefinitions.Add($c1); $grid.ColumnDefinitions.Add($c2)

    $check = New-Object System.Windows.Controls.CheckBox
    $check.IsChecked = $task.done
    $check.Margin = New-Object System.Windows.Thickness(0,2,12,0)
    $check.VerticalAlignment = 'Top'
    $check.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(138,148,168))
    [System.Windows.Controls.Grid]::SetColumn($check, 0)

    $stack = New-Object System.Windows.Controls.StackPanel
    [System.Windows.Controls.Grid]::SetColumn($stack, 1)

    $primaryText   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(240,242,247))
    $secondaryText = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(138,148,168))
    $tertiaryText  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(85,94,112))
    $urgentText    = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255,122,144))

    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = $task.title
    $title.FontSize = 14
    $title.LineHeight = 19
    $title.TextWrapping = 'Wrap'
    switch ($task.priority) {
        'critical' { $title.FontWeight = 'Bold';     $title.Foreground = $primaryText }
        'high'     { $title.FontWeight = 'SemiBold'; $title.Foreground = $primaryText }
        'normal'   { $title.FontWeight = 'Normal';   $title.Foreground = $primaryText }
        'low'      { $title.FontWeight = 'Normal';   $title.Foreground = $secondaryText }
    }
    if ($task.done) {
        $title.TextDecorations = [System.Windows.TextDecorations]::Strikethrough
        $title.Opacity = 0.45
    }
    $stack.Children.Add($title) | Out-Null

    $metaText = @()
    $due = Format-DueLabel $task.dueDate
    if ($due) { $metaText += $due }
    if ($task.source -ne 'manual') { $metaText += $task.source }
    if ($metaText.Count -gt 0) {
        $meta = New-Object System.Windows.Controls.TextBlock
        $meta.Text = ($metaText -join '   ')
        $meta.FontSize = 11
        $meta.Margin = New-Object System.Windows.Thickness(0,4,0,0)
        if ($isOverdue) { $meta.Foreground = $urgentText; $meta.FontWeight = 'SemiBold' }
        else            { $meta.Foreground = $tertiaryText }
        $stack.Children.Add($meta) | Out-Null
    }

    $grid.Children.Add($check) | Out-Null
    $grid.Children.Add($stack) | Out-Null
    $border.Child = $grid

    # Context menu - priority, due date, delete
    $menu = New-Object System.Windows.Controls.ContextMenu

    $prioRoot = New-Object System.Windows.Controls.MenuItem
    $prioRoot.Header = 'Priority'
    foreach ($p in 'low','normal','high','critical') {
        $mi = New-Object System.Windows.Controls.MenuItem
        $mi.Header = $p
        $mi.Tag = @{ id = $task.id; action = 'priority'; value = $p }
        $mi.Add_Click({ Handle-ContextAction $this.Tag })
        $prioRoot.Items.Add($mi) | Out-Null
    }
    $menu.Items.Add($prioRoot) | Out-Null

    $dueRoot = New-Object System.Windows.Controls.MenuItem
    $dueRoot.Header = 'Due'

    $nowDate = (Get-Date).Date
    $daysToFri = (([int][System.DayOfWeek]::Friday) - [int]$nowDate.DayOfWeek + 7) % 7
    if ($daysToFri -eq 0) { $daysToFri = 7 }
    $fridayIso = $nowDate.AddDays($daysToFri).AddHours(17).ToString('o')

    $daysToMon = (([int][System.DayOfWeek]::Monday) - [int]$nowDate.DayOfWeek + 7) % 7
    if ($daysToMon -eq 0) { $daysToMon = 7 }
    $mondayIso = $nowDate.AddDays($daysToMon).AddHours(9).ToString('o')

    $dueOptions = @(
        @{ label = 'None';                 when = '' }
        @{ label = 'Today 5 PM';           when = $nowDate.AddHours(17).ToString('o') }
        @{ label = 'Tomorrow 9 AM';        when = $nowDate.AddDays(1).AddHours(9).ToString('o') }
        @{ label = 'This week (Fri 5 PM)'; when = $fridayIso }
        @{ label = 'Next Monday 9 AM';     when = $mondayIso }
    )
    foreach ($opt in $dueOptions) {
        $mi = New-Object System.Windows.Controls.MenuItem
        $mi.Header = $opt.label
        $mi.Tag = @{ id = $task.id; action = 'due'; value = $opt.when }
        $mi.Add_Click({ Handle-ContextAction $this.Tag })
        $dueRoot.Items.Add($mi) | Out-Null
    }
    $menu.Items.Add($dueRoot) | Out-Null

    $menu.Items.Add((New-Object System.Windows.Controls.Separator)) | Out-Null

    $delItem = New-Object System.Windows.Controls.MenuItem
    $delItem.Header = 'Delete'
    $delItem.Tag = @{ id = $task.id; action = 'delete' }
    $delItem.Add_Click({ Handle-ContextAction $this.Tag })
    $menu.Items.Add($delItem) | Out-Null
    $border.ContextMenu = $menu

    # Check toggle - static, no fade
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

$script:SortLabels = @{
    smart    = 'Smart'
    priority = 'Priority'
    due      = 'Due date'
    added    = 'Newest'
}

function Refresh-Render {
    $TaskList.Items.Clear()
    $visible = if ($script:Settings.showCompleted) { $script:Tasks } else { @($script:Tasks | Where-Object { -not $_.done }) }
    $mode = 'smart'
    if ($script:Settings.sortMode) { $mode = [string]$script:Settings.sortMode }
    $sorted = Sort-VigilTasks -tasks $visible -mode $mode
    foreach ($t in $sorted) {
        $card = Build-TaskCard $t
        $TaskList.Items.Add($card) | Out-Null
    }
    $active = @($script:Tasks | Where-Object { -not $_.done }).Count
    $CountText.Text = [string]$active
    $CountBadge.Visibility = if ($active -gt 0) { 'Visible' } else { 'Collapsed' }
    $StatusRight.Text = ('{0} active' -f $active)
    $StatusLeft.Text  = $script:VigilVersion

    $label = $script:SortLabels[$mode]
    if (-not $label) { $label = 'Smart' }
    $BtnSort.Content = ($label + ' ' + [string][char]0x25BE)
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
    } elseif ($action -eq 'due') {
        $t = $script:Tasks | Where-Object { $_.id -eq $id }
        if ($t) { $t.dueDate = $tag.value }
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

$script:IsCollapsed = $false
$BtnCollapse.Add_Click({
    if ($script:IsCollapsed) {
        $TaskArea.Visibility   = 'Visible'
        $AddArea.Visibility    = 'Visible'
        $StatusArea.Visibility = 'Visible'
        $BtnSort.Visibility    = 'Visible'
        $script:IsCollapsed = $false
    } else {
        $TaskArea.Visibility   = 'Collapsed'
        $AddArea.Visibility    = 'Collapsed'
        $StatusArea.Visibility = 'Collapsed'
        $BtnSort.Visibility    = 'Collapsed'
        $script:IsCollapsed = $true
    }
})

$sortMenu = New-Object System.Windows.Controls.ContextMenu
$sortModes = @(
    @{ key = 'smart';    label = 'Smart (overdue + priority)' }
    @{ key = 'priority'; label = 'Priority' }
    @{ key = 'due';      label = 'Due date' }
    @{ key = 'added';    label = 'Newest first' }
)
foreach ($opt in $sortModes) {
    $mi = New-Object System.Windows.Controls.MenuItem
    $mi.Header = $opt.label
    $mi.Tag = $opt.key
    $mi.Add_Click({
        $k = $this.Tag
        $script:Settings.sortMode = $k
        Save-VigilSettings $script:Settings
        Refresh-Render
    })
    $sortMenu.Items.Add($mi) | Out-Null
}
$BtnSort.Add_Click({
    $sortMenu.PlacementTarget = $BtnSort
    $sortMenu.Placement = 'Bottom'
    $sortMenu.IsOpen = $true
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
$startMsg = 'VIGIL started. version={0}  tasks={1}' -f $script:VigilVersion, $script:Tasks.Count
Write-VigilLog $startMsg
Refresh-Render
[void]$window.ShowDialog()
