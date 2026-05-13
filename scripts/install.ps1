#Requires -Version 5.1
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string]$DistRoot
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    throw "Run this installer from 64-bit PowerShell so System32 is not redirected to SysWOW64."
}

function Assert-SecureBootAllowsAppInit {
    if (-not (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue)) {
        Write-Host "Secure Boot state could not be queried because Confirm-SecureBootUEFI is unavailable; continuing."
        return
    }

    try {
        $secureBootEnabled = Confirm-SecureBootUEFI
    } catch [System.PlatformNotSupportedException] {
        Write-Host "Secure Boot state could not be queried on this platform; continuing."
        return
    } catch [System.UnauthorizedAccessException] {
        throw "Could not query Secure Boot state. Run this installer from an elevated PowerShell session."
    }

    if ($secureBootEnabled) {
        throw "Secure Boot is enabled. Windows disables AppInit_DLLs when Secure Boot is enabled, so comctl32v6hook cannot be installed usefully."
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $DistRoot) {
    $DistRoot = Join-Path $repoRoot "dist"
}

function ConvertTo-FullPath {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [string]$BasePath = $repoRoot
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Copy-DeploymentPair {
    param(
        [Parameter(Mandatory)]
        [string]$SourceDirectory,

        [Parameter(Mandatory)]
        [string]$TargetDirectory
    )

    $dllSource = Join-Path $SourceDirectory "comctl32v6hook.dll"
    $manifestSource = Join-Path $SourceDirectory "comctl32v6hook.manifest"

    if (-not (Test-Path -LiteralPath $dllSource)) {
        throw "Missing deployment DLL: $dllSource"
    }
    if (-not (Test-Path -LiteralPath $manifestSource)) {
        throw "Missing deployment manifest next to DLL: $manifestSource"
    }

    Copy-Item -LiteralPath $dllSource -Destination (Join-Path $TargetDirectory "comctl32v6hook.dll") -Force
    Copy-Item -LiteralPath $manifestSource -Destination (Join-Path $TargetDirectory "comctl32v6hook.manifest") -Force
}

function Assert-DeploymentPair {
    param(
        [Parameter(Mandatory)]
        [string]$SourceDirectory
    )

    $dllSource = Join-Path $SourceDirectory "comctl32v6hook.dll"
    $manifestSource = Join-Path $SourceDirectory "comctl32v6hook.manifest"

    if (-not (Test-Path -LiteralPath $dllSource)) {
        throw "Missing deployment DLL: $dllSource. Prepare the dist output first or pass -DistRoot with prepared output."
    }
    if (-not (Test-Path -LiteralPath $manifestSource)) {
        throw "Missing deployment manifest: $manifestSource. Prepare the dist output first or pass -DistRoot with prepared output."
    }
}

function Set-RegistryValue {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [object]$Value,

        [Parameter(Mandatory)]
        [Microsoft.Win32.RegistryValueKind]$Type
    )

    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
}

function ConvertFrom-ImageFileMachine {
    param(
        [Parameter(Mandatory)]
        [uint16]$Machine
    )

    switch ($Machine) {
        0x014c { return "x86" }
        0x8664 { return "x64" }
        0xaa64 { return "arm64" }
        default {
            throw ("Unsupported native machine type: 0x{0:X4}" -f $Machine)
        }
    }
}

function Get-NativeArchitecture {
    if (-not ("Comctl32v6hook.NativeMethods" -as [type])) {
        Add-Type -Namespace Comctl32v6hook -Name NativeMethods -MemberDefinition @"
            [System.Runtime.InteropServices.DllImport("kernel32.dll")]
            public static extern System.IntPtr GetCurrentProcess();

            [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
            [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
            public static extern bool IsWow64Process2(System.IntPtr hProcess, out ushort processMachine, out ushort nativeMachine);
"@
    }

    [uint16]$processMachine = 0
    [uint16]$nativeMachine = 0
    if ([Comctl32v6hook.NativeMethods]::IsWow64Process2(
            [Comctl32v6hook.NativeMethods]::GetCurrentProcess(),
            [ref]$processMachine,
            [ref]$nativeMachine)) {
        if ($nativeMachine -ne 0) {
            return ConvertFrom-ImageFileMachine $nativeMachine
        }

        if ($processMachine -ne 0) {
            return ConvertFrom-ImageFileMachine $processMachine
        }
    }

    $architecture = $env:PROCESSOR_ARCHITEW6432
    if (-not $architecture) {
        $architecture = $env:PROCESSOR_ARCHITECTURE
    }
    if ([string]::IsNullOrWhiteSpace($architecture)) {
        throw "Could not determine native architecture because PROCESSOR_ARCHITECTURE is not set."
    }

    switch ($architecture.ToUpperInvariant()) {
        "AMD64" { return "x64" }
        "ARM64" { return "arm64" }
        "X86" { return "x86" }
        default {
            throw "Could not determine native architecture from PROCESSOR_ARCHITECTURE=$architecture."
        }
    }
}

$DistRoot = ConvertTo-FullPath $DistRoot

$system32 = Join-Path $env:WINDIR "System32"
$syswow64 = Join-Path $env:WINDIR "SysWOW64"
$nativeArch = Get-NativeArchitecture
$nativeDist = Join-Path $DistRoot $nativeArch
$nativeAppInitPath = Join-Path $system32 "comctl32v6hook.dll"

Assert-SecureBootAllowsAppInit

Assert-DeploymentPair -SourceDirectory $nativeDist
Copy-DeploymentPair -SourceDirectory $nativeDist -TargetDirectory $system32

if ([Environment]::Is64BitOperatingSystem) {
    $wow64Dist = Join-Path $DistRoot "x86"
    Assert-DeploymentPair -SourceDirectory $wow64Dist
    Copy-DeploymentPair -SourceDirectory $wow64Dist -TargetDirectory $syswow64
    $wow64AppInitPath = Join-Path $syswow64 "comctl32v6hook.dll"
}

$nativeKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows"
Set-RegistryValue -Path $nativeKey -Name "AppInit_DLLs" -Value $nativeAppInitPath -Type String
Set-RegistryValue -Path $nativeKey -Name "LoadAppInit_DLLs" -Value 1 -Type DWord
Set-RegistryValue -Path $nativeKey -Name "RequireSignedAppInit_DLLs" -Value 0 -Type DWord

if ([Environment]::Is64BitOperatingSystem) {
    $wow64Key = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows"
    Set-RegistryValue -Path $wow64Key -Name "AppInit_DLLs" -Value $wow64AppInitPath -Type String
    Set-RegistryValue -Path $wow64Key -Name "LoadAppInit_DLLs" -Value 1 -Type DWord
    Set-RegistryValue -Path $wow64Key -Name "RequireSignedAppInit_DLLs" -Value 0 -Type DWord
}

Write-Host "Installed comctl32v6hook:"
Write-Host "  Native ($nativeArch): $nativeAppInitPath"
if ([Environment]::Is64BitOperatingSystem) {
    Write-Host "  Wow64 : $wow64AppInitPath"
}
Write-Host ""
Write-Host "Restart target GUI processes. A reboot is recommended if an older AppInit path was already loaded."
