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
$script:VigilVersion = '2026-04-14 00:15 UTC  build24 hotkey-csharp-bridge'

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.Windows.Forms

# --- Win32 P/Invoke (foreground tracking + window activation) --------------
if (-not ([System.Management.Automation.PSTypeName]'VigilWin32').Type) {
    Add-Type -ErrorAction Stop -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class VigilWin32 {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
"@
}

# --- Hotkey helper (C# bridge so PS avoids HwndSourceHook ref-delegate cast)
if (-not ([System.Management.Automation.PSTypeName]'VigilHotkey').Type) {
    Add-Type -ErrorAction Stop -ReferencedAssemblies PresentationCore, WindowsBase -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;
public static class VigilHotkey {
    public static event Action HotkeyPressed;
    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID = 9001;
    private static HwndSource _src;

    public static bool Register(IntPtr hwnd, uint mods, uint vk) {
        _src = HwndSource.FromHwnd(hwnd);
        if (_src == null) return false;
        _src.AddHook(WndProc);
        return RegisterHotKey(hwnd, HOTKEY_ID, mods, vk);
    }

    public static void Unregister(IntPtr hwnd) {
        try { UnregisterHotKey(hwnd, HOTKEY_ID); } catch { }
        if (_src != null) { try { _src.RemoveHook(WndProc); } catch { } _src = null; }
    }

    private static IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
        if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID) {
            var h = HotkeyPressed;
            if (h != null) h();
        }
        return IntPtr.Zero;
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
"@
}

# --- Single instance -------------------------------------------------------
$script:Mutex = New-Object System.Threading.Mutex($false, 'Global\VIGIL_TaskTracker')
if (-not $script:Mutex.WaitOne(0, $false)) {
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
        $mi.Add_Click({ param($s, $e) Handle-ContextAction $s.Tag })
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
        $mi.Add_Click({ param($s, $e) Handle-ContextAction $s.Tag })
        $dueRoot.Items.Add($mi) | Out-Null
    }
    $menu.Items.Add($dueRoot) | Out-Null

    $menu.Items.Add((New-Object System.Windows.Controls.Separator)) | Out-Null

    $delItem = New-Object System.Windows.Controls.MenuItem
    $delItem.Header = 'Delete'
    $delItem.Tag = @{ id = $task.id; action = 'delete' }
    $delItem.Add_Click({ param($s, $e) Handle-ContextAction $s.Tag })
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
        param($s, $e)
        $k = [string]$s.Tag
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

# --- Phase 2: Quick-Add popup + global hotkey -----------------------------

$quickAddXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="VIGIL - New task"
        Width="460" SizeToContent="Height"
        WindowStyle="None" ResizeMode="NoResize"
        AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False"
        TextOptions.TextFormattingMode="Ideal"
        TextOptions.TextRenderingMode="Grayscale"
        UseLayoutRounding="True" SnapsToDevicePixels="True"
        FontFamily="Segoe UI"
        WindowStartupLocation="Manual">
  <Border CornerRadius="16" Background="#0B0E14"
          BorderBrush="#2A3245" BorderThickness="1" Margin="14">
    <Border.Effect>
      <DropShadowEffect Color="#000000" BlurRadius="40" ShadowDepth="0" Opacity="0.75"/>
    </Border.Effect>
    <StackPanel Margin="22,20,22,20">
      <TextBlock Text="New task" FontSize="15" FontWeight="SemiBold"
                 Foreground="#F0F2F7" Margin="0,0,0,12"/>
      <Border Background="#1C2230" BorderBrush="#2A3245" BorderThickness="1"
              CornerRadius="8">
        <TextBox x:Name="TxtTitle" Background="Transparent" BorderThickness="0"
                 Foreground="#F0F2F7" Padding="12,10" FontSize="14"
                 CaretBrush="#7C5CFF" VerticalContentAlignment="Center"/>
      </Border>
      <TextBlock Text="PRIORITY" FontSize="10" FontWeight="Bold"
                 Foreground="#555E70" Margin="2,16,0,6"/>
      <StackPanel x:Name="PriorityRow" Orientation="Horizontal"/>
      <TextBlock Text="DUE" FontSize="10" FontWeight="Bold"
                 Foreground="#555E70" Margin="2,14,0,6"/>
      <StackPanel x:Name="DueRow" Orientation="Horizontal"/>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
        <Button x:Name="BtnQCancel" Content="Cancel"
                Background="Transparent" Foreground="#8A94A8" BorderThickness="0"
                Padding="14,8" FontSize="12" Cursor="Hand"/>
        <Button x:Name="BtnQSave" Content="Add task"
                Background="#7C5CFF" Foreground="#FFFFFF" BorderThickness="0"
                Padding="20,8" FontSize="12" FontWeight="SemiBold"
                Margin="8,0,0,0" Cursor="Hand"/>
      </StackPanel>
    </StackPanel>
  </Border>
</Window>
'@

$script:QuickAddPriority = 'normal'
$script:QuickAddDue      = ''
$script:QuickAddPriBtns  = $null
$script:QuickAddDueBtns  = $null

function Update-QuickAddVisuals {
    $accent  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(124,92,255))
    $elev    = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(28,34,48))
    $txt     = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(240,242,247))
    $txtDim  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(138,148,168))
    $bColor  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(42,50,69))

    if ($script:QuickAddPriBtns) {
        foreach ($p in $script:QuickAddPriBtns.Keys) {
            $b = $script:QuickAddPriBtns[$p]
            if ($p -eq $script:QuickAddPriority) {
                $b.Background = $accent; $b.Foreground = $txt; $b.BorderBrush = $accent
            } else {
                $b.Background = $elev; $b.Foreground = $txtDim; $b.BorderBrush = $bColor
            }
        }
    }
    if ($script:QuickAddDueBtns) {
        foreach ($k in $script:QuickAddDueBtns.Keys) {
            $b = $script:QuickAddDueBtns[$k]
            $matchTag = ($b.Tag -eq $script:QuickAddDue)
            if ($matchTag) {
                $b.Background = $accent; $b.Foreground = $txt; $b.BorderBrush = $accent
            } else {
                $b.Background = $elev; $b.Foreground = $txtDim; $b.BorderBrush = $bColor
            }
        }
    }
}

function Show-QuickAdd {
    try {
        $prevFg = [VigilWin32]::GetForegroundWindow()

        $clip = ''
        try { $clip = [System.Windows.Clipboard]::GetText() } catch {}
        if ($clip) {
            $clip = $clip.Trim() -replace '\s+', ' '
            if ($clip.Length -gt 200) { $clip = $clip.Substring(0, 200) }
        }

        [xml]$qx = $quickAddXaml
        $qreader = New-Object System.Xml.XmlNodeReader $qx
        $qwin = [Windows.Markup.XamlReader]::Load($qreader)

        $txtTitle  = $qwin.FindName('TxtTitle')
        $priRow    = $qwin.FindName('PriorityRow')
        $dueRow    = $qwin.FindName('DueRow')
        $btnSave   = $qwin.FindName('BtnQSave')
        $btnCancel = $qwin.FindName('BtnQCancel')

        $script:QuickAddPriority = 'normal'
        $script:QuickAddDue      = ''
        $script:QuickAddPriBtns  = @{}
        $script:QuickAddDueBtns  = @{}

        # Priority pills
        foreach ($p in @('low','normal','high','critical')) {
            $btn = New-Object System.Windows.Controls.Button
            $btn.Content = $p.Substring(0,1).ToUpper() + $p.Substring(1)
            $btn.Margin = New-Object System.Windows.Thickness(0,0,6,0)
            $btn.Padding = New-Object System.Windows.Thickness(14,7,14,7)
            $btn.FontSize = 12
            $btn.BorderThickness = New-Object System.Windows.Thickness(1)
            $btn.Cursor = [System.Windows.Input.Cursors]::Hand
            $btn.Tag = $p
            $btn.Add_Click({
                param($sender, $e)
                $script:QuickAddPriority = [string]$sender.Tag
                Update-QuickAddVisuals
            })
            $script:QuickAddPriBtns[$p] = $btn
            [void]$priRow.Children.Add($btn)
        }

        # Due pills
        $nowDate = (Get-Date).Date
        $daysToFri = (([int][System.DayOfWeek]::Friday) - [int]$nowDate.DayOfWeek + 7) % 7
        if ($daysToFri -eq 0) { $daysToFri = 7 }
        $dueChoices = @(
            @{ key = '';                                                       label = 'None' }
            @{ key = $nowDate.AddHours(17).ToString('o');                      label = 'Today 5pm' }
            @{ key = $nowDate.AddDays(1).AddHours(9).ToString('o');            label = 'Tomorrow 9am' }
            @{ key = $nowDate.AddDays($daysToFri).AddHours(17).ToString('o');  label = 'This week' }
        )
        foreach ($d in $dueChoices) {
            $btn = New-Object System.Windows.Controls.Button
            $btn.Content = $d.label
            $btn.Margin = New-Object System.Windows.Thickness(0,0,6,0)
            $btn.Padding = New-Object System.Windows.Thickness(14,7,14,7)
            $btn.FontSize = 12
            $btn.BorderThickness = New-Object System.Windows.Thickness(1)
            $btn.Cursor = [System.Windows.Input.Cursors]::Hand
            $btn.Tag = $d.key
            $btn.Add_Click({
                param($sender, $e)
                $script:QuickAddDue = [string]$sender.Tag
                Update-QuickAddVisuals
            })
            $script:QuickAddDueBtns[$d.label] = $btn
            [void]$dueRow.Children.Add($btn)
        }

        Update-QuickAddVisuals

        # Position on active monitor
        $pt = [System.Windows.Forms.Cursor]::Position
        $scr = [System.Windows.Forms.Screen]::FromPoint($pt).WorkingArea
        $qwin.Left = $scr.X + (($scr.Width  - 460) / 2)
        $qwin.Top  = $scr.Y + (($scr.Height - 320) / 2)

        if ($clip) {
            $txtTitle.Text = $clip
            $txtTitle.SelectAll()
        }

        $saveAction = {
            $t = $txtTitle.Text
            if ($t) { $t = $t.Trim() }
            if (-not $t) { return }
            $task = New-VigilTask -Title $t -Priority $script:QuickAddPriority
            if ($script:QuickAddDue) { $task.dueDate = $script:QuickAddDue }
            $script:Tasks = @($script:Tasks) + @($task)
            Save-VigilTasks $script:Tasks
            Refresh-Render
            $qwin.Close()
        }.GetNewClosure()

        $btnSave.Add_Click($saveAction)
        $btnCancel.Add_Click({ $qwin.Close() }.GetNewClosure())

        $txtTitle.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq 'Return') { & $saveAction; $e.Handled = $true }
            elseif ($e.Key -eq 'Escape') { $qwin.Close(); $e.Handled = $true }
        }.GetNewClosure())

        $qwin.Add_KeyDown({
            param($s, $e)
            if ($e.Key -eq 'Escape') { $qwin.Close(); $e.Handled = $true }
        }.GetNewClosure())

        $qwin.Add_Closed({
            if ($prevFg -ne [IntPtr]::Zero) {
                [VigilWin32]::SetForegroundWindow($prevFg) | Out-Null
            }
        }.GetNewClosure())

        $qwin.Show()
        $qwin.Activate() | Out-Null
        $txtTitle.Focus() | Out-Null
    } catch {
        $emsg = 'Show-QuickAdd FAILED: {0}' -f $_.Exception.Message
        Write-VigilLog $emsg
    }
}

# --- Global hotkey registration (Ctrl+Win+A) ------------------------------

$script:HotkeyRegistered = $false

# Subscribe to the C# event — plain Action delegate, no ref-param cast needed
[VigilHotkey]::add_HotkeyPressed({
    try { Show-QuickAdd } catch {
        $em = 'Hotkey handler error: ' + $_.Exception.Message
        Write-VigilLog $em
    }
})

$window.Add_Loaded({
    try {
        $helper = New-Object System.Windows.Interop.WindowInteropHelper($window)
        $hwnd = $helper.Handle
        $MOD_CONTROL = [uint32]0x0002
        $MOD_WIN     = [uint32]0x0008
        $VK_A        = [uint32]0x41
        $mods = [uint32]($MOD_CONTROL -bor $MOD_WIN)
        $ok = [VigilHotkey]::Register($hwnd, $mods, $VK_A)
        if ($ok) {
            $script:HotkeyRegistered = $true
            Write-VigilLog 'Hotkey registered: Ctrl+Win+A'
        } else {
            Write-VigilLog 'Hotkey registration FAILED - Ctrl+Win+A may already be in use'
        }
    } catch {
        $em = 'Hotkey setup error: ' + $_.Exception.Message
        Write-VigilLog $em
    }
})

$window.Add_Closing({
    if ($script:HotkeyRegistered) {
        try {
            $helper = New-Object System.Windows.Interop.WindowInteropHelper($window)
            [VigilHotkey]::Unregister($helper.Handle)
        } catch {}
    }
})

Refresh-Render

# Wrap ShowDialog so any handler failure lands in vigil.log with context
try {
    [void]$window.ShowDialog()
} catch {
    $ex = $_.Exception
    $msg = 'ShowDialog FAILED: {0}' -f $ex.Message
    Write-VigilLog $msg
    if ($ex.InnerException) {
        $im = 'Inner: {0}' -f $ex.InnerException.Message
        Write-VigilLog $im
    }
    if ($ex.StackTrace) {
        Write-VigilLog $ex.StackTrace
    }
    throw
}
