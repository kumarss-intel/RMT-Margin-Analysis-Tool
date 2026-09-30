<#
.SYNOPSIS
    Creates (or removes) the MarginIQ launcher shortcuts with the MarginIQ icon.

.DESCRIPTION
    Called by setup.bat, but can also be run on its own. Every shortcut starts
    the GUI exactly like Launch_RMT_GUI.bat does (the tool's .venv pythonw.exe
    running rmt_gui.py, no console window) and carries:
      * the MarginIQ icon      (assets\marginiq.ico)
      * the AppUserModelID     "Intel.CCG.CVE.MarginIQ" - the same ID rmt_gui.py
        sets at start-up, so the running window groups under the pinned /
        Start-menu shortcut and shows the MarginIQ icon on the taskbar.

    Locations:
      -Desktop      <Desktop>\MarginIQ.lnk
      -Local        <tool folder>\MarginIQ.lnk  (icon'd twin of Launch_RMT_GUI.bat)
      -StartMenu    Start > All apps > MarginIQ
      -QuickLaunch  %APPDATA%\Microsoft\Internet Explorer\Quick Launch\MarginIQ.lnk
      -Taskbar      Start-menu shortcut + pin attempt. Windows 10 1809+ / 11 block
                    programmatic pinning; the script then prints the one-click
                    manual steps (no window is opened).

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\create_shortcuts.ps1 -Desktop -Local -StartMenu

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\create_shortcuts.ps1 -Remove
#>
[CmdletBinding()]
param(
    [string]$ToolDir = '',
    [switch]$Desktop,
    [switch]$Local,
    [switch]$StartMenu,
    [switch]$QuickLaunch,
    [switch]$Taskbar,
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'
$AppId   = 'Intel.CCG.CVE.MarginIQ'
$Name    = 'MarginIQ'
$Comment = 'MarginIQ - Intel CCG CVE DDR5 RMT Margin Analysis Tool'

if (-not $ToolDir) { $ToolDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
$ToolDir  = (Resolve-Path -LiteralPath $ToolDir).Path
$gui      = Join-Path $ToolDir 'rmt_gui.py'
$icon     = Join-Path $ToolDir 'assets\marginiq.ico'
$pythonw  = Join-Path $ToolDir '.venv\Scripts\pythonw.exe'
$launcher = Join-Path $ToolDir 'Launch_RMT_GUI.bat'

$paths = [ordered]@{
    Desktop     = Join-Path ([Environment]::GetFolderPath('Desktop')) "$Name.lnk"
    Local       = Join-Path $ToolDir "$Name.lnk"
    StartMenu   = Join-Path ([Environment]::GetFolderPath('Programs')) "$Name.lnk"
    QuickLaunch = Join-Path $env:APPDATA "Microsoft\Internet Explorer\Quick Launch\$Name.lnk"
}
$pinnedDir = Join-Path $env:APPDATA 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'

if ($Remove) {
    foreach ($p in @($paths.Values) + (Join-Path $pinnedDir "$Name.lnk")) {
        if (Test-Path -LiteralPath $p) {
            Remove-Item -LiteralPath $p -Force
            Write-Host "      Removed $p"
        }
    }
    Write-Host '      (Unpin a taskbar icon with right-click > Unpin from taskbar.)'
    exit 0
}

if (-not (Test-Path -LiteralPath $gui)) { Write-Host "[ERROR] rmt_gui.py not found in $ToolDir"; exit 2 }
if (-not (Test-Path -LiteralPath $icon)) { Write-Host "[WARN] Icon not found ($icon) - shortcuts will use the default icon." }

# Stamp the AppUserModelID on a .lnk via IPropertyStore (not exposed by WScript.Shell).
$src = @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class MarginIQLnk {
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    struct PROPERTYKEY { public Guid fmtid; public uint pid; }

    [StructLayout(LayoutKind.Explicit)]
    struct PROPVARIANT {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr p;
        [FieldOffset(16)] public IntPtr p2;
    }

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStore {
        [PreserveSig] int GetCount(out uint c);
        [PreserveSig] int GetAt(uint i, out PROPERTYKEY k);
        [PreserveSig] int GetValue(ref PROPERTYKEY k, out PROPVARIANT v);
        [PreserveSig] int SetValue(ref PROPERTYKEY k, ref PROPVARIANT v);
        [PreserveSig] int Commit();
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    class CShellLink { }

    public static void SetAppId(string lnkPath, string appId) {
        object link = new CShellLink();
        try {
            IPersistFile pf = (IPersistFile)link;
            pf.Load(lnkPath, 2 /* STGM_READWRITE */);
            IPropertyStore store = (IPropertyStore)link;
            PROPERTYKEY key = new PROPERTYKEY();
            key.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");
            key.pid = 5;
            PROPVARIANT pv = new PROPVARIANT();
            pv.vt = 31; /* VT_LPWSTR */
            pv.p = Marshal.StringToCoTaskMemUni(appId);
            try {
                Marshal.ThrowExceptionForHR(store.SetValue(ref key, ref pv));
                Marshal.ThrowExceptionForHR(store.Commit());
            } finally {
                Marshal.FreeCoTaskMem(pv.p);
            }
            pf.Save(lnkPath, true);
        } finally {
            Marshal.ReleaseComObject(link);
        }
    }
}
'@
$appIdOk = $true
try { Add-Type -TypeDefinition $src -Language CSharp -ErrorAction Stop }
catch {
    if (-not ('MarginIQLnk' -as [type])) {
        $appIdOk = $false
        Write-Host "[WARN] Could not load the AppUserModelID helper: $($_.Exception.Message)"
    }
}

$shell = New-Object -ComObject WScript.Shell

function New-MarginIQShortcut([string]$Path) {
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $lnk = $shell.CreateShortcut($Path)
    if (Test-Path -LiteralPath $pythonw) {
        $lnk.TargetPath = $pythonw
        $lnk.Arguments  = "`"$gui`""
    } else {
        # No .venv yet: fall back to the batch launcher (minimised console).
        $lnk.TargetPath  = $launcher
        $lnk.WindowStyle = 7
    }
    $lnk.WorkingDirectory = $ToolDir
    $lnk.Description      = $Comment
    if (Test-Path -LiteralPath $icon) { $lnk.IconLocation = "$icon,0" }
    $lnk.Save()
    if ($appIdOk) {
        try { [MarginIQLnk]::SetAppId($Path, $AppId) }
        catch { Write-Host "[WARN] AppUserModelID not set on ${Path}: $($_.Exception.Message)" }
    }
    Write-Host "      Created  $Path"
}

$want = [ordered]@{
    Desktop     = [bool]$Desktop
    Local       = [bool]$Local
    StartMenu   = [bool]($StartMenu -or $Taskbar)
    QuickLaunch = [bool]$QuickLaunch
}
$failed = 0
foreach ($k in $want.Keys) {
    if (-not $want[$k]) { continue }
    try { New-MarginIQShortcut $paths[$k] }
    catch { $failed++; Write-Host "[WARN] Could not create the $k shortcut: $($_.Exception.Message)" }
}
if ($QuickLaunch) {
    Write-Host '      Note: the Quick Launch toolbar is only shown on Windows 10 when enabled'
    Write-Host '            (taskbar > Toolbars > New toolbar > the Quick Launch folder).'
    Write-Host '            On Windows 11 use the Start menu entry / taskbar pin instead.'
}

if ($Taskbar) {
    $pinned = Join-Path $pinnedDir "$Name.lnk"
    $startLnk = $paths.StartMenu
    if (-not (Test-Path -LiteralPath $pinned) -and (Test-Path -LiteralPath $startLnk)) {
        try {
            $folder = (New-Object -ComObject Shell.Application).Namespace((Split-Path -Parent $startLnk))
            $item = $folder.ParseName((Split-Path -Leaf $startLnk))
            $verb = $item.Verbs() | Where-Object { ($_.Name -replace '&', '') -match 'Pin to taskbar|taskbarpin' } | Select-Object -First 1
            if ($verb) { $verb.DoIt(); Start-Sleep -Milliseconds 800 }
        } catch { }
    }
    if (Test-Path -LiteralPath $pinned) {
        Write-Host '      Pinned   MarginIQ to the taskbar.'
    } else {
        Write-Host ''
        Write-Host '      Windows does not allow scripts to pin to the taskbar, so finish with one click:'
        Write-Host '        Start > All apps > MarginIQ  -> right-click -> Pin to taskbar'
        Write-Host '        (or launch MarginIQ, right-click its taskbar button -> Pin to taskbar)'
    }
}

if ($failed) { exit 1 }
exit 0
