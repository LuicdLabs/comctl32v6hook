#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet("x64", "x86", "arm64", "both", "all")]
    [string]$Architecture = "both",

    [ValidateSet("Debug", "Release", "RelWithDebInfo", "MinSizeRel")]
    [string]$Configuration = "Release",

    [string]$BuildRoot,

    [string]$DistRoot,

    [string]$VcpkgRoot,

    [switch]$NoBootstrapVcpkg,

    [switch]$Clean
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $BuildRoot) {
    $BuildRoot = Join-Path $repoRoot "build"
}
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

function Invoke-Native {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [string[]]$Arguments,

        [string]$WorkingDirectory
    )

    $oldLocation = Get-Location
    try {
        if ($WorkingDirectory) {
            Set-Location -LiteralPath $WorkingDirectory
        }

        & $FilePath @Arguments | ForEach-Object {
            Write-Host $_
        }
        if ($LASTEXITCODE -ne 0) {
            throw "$FilePath failed with exit code $LASTEXITCODE."
        }
    } finally {
        Set-Location $oldLocation
    }
}

function Resolve-CMake {
    $cmd = Get-Command cmake.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $programRoots = @(
        ${env:ProgramFiles},
        ${env:ProgramFiles(x86)}
    ) | Where-Object { $_ }

    foreach ($root in $programRoots) {
        $pattern = Join-Path $root "Microsoft Visual Studio\*\*\Common7\IDE\CommonExtensions\Microsoft\CMake\CMake\bin\cmake.exe"
        $match = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object -Property FullName -Descending |
            Select-Object -First 1

        if ($match) {
            return $match.FullName
        }
    }

    throw "cmake.exe was not found. Install CMake or Visual Studio with CMake tools."
}

function Get-VisualStudioGeneratorRank {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($Name -match "^Visual Studio\s+(\d+)\s+") {
        return [int]$Matches[1]
    }

    return 0
}

function Resolve-CMakeGenerator {
    param(
        [Parameter(Mandatory)]
        [string]$CMake
    )

    $json = & $CMake -E capabilities
    if ($LASTEXITCODE -ne 0) {
        throw "Could not query CMake generator capabilities."
    }

    $capabilities = $json | ConvertFrom-Json
    $generators = @($capabilities.generators) | Where-Object {
        $_.name -like "Visual Studio*" -and $_.platformSupport
    }

    if (-not $generators) {
        throw "No Visual Studio CMake generator was found. Install Visual Studio with the Desktop development with C++ workload."
    }

    $generator = $generators |
        Sort-Object `
            @{ Expression = { Get-VisualStudioGeneratorRank $_.name }; Descending = $true },
            @{ Expression = { $_.name }; Descending = $true } |
        Select-Object -First 1

    return $generator.name
}

function Resolve-Git {
    $cmd = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $programRoots = @(
        ${env:ProgramFiles},
        ${env:ProgramFiles(x86)}
    ) | Where-Object { $_ }

    foreach ($root in $programRoots) {
        $patterns = @(
            "Microsoft Visual Studio\*\*\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\cmd\git.exe",
            "Microsoft Visual Studio\*\*\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer\Git\mingw64\bin\git.exe"
        )

        foreach ($pattern in $patterns) {
            $match = Get-ChildItem -Path (Join-Path $root $pattern) -ErrorAction SilentlyContinue |
                Sort-Object -Property FullName -Descending |
                Select-Object -First 1

            if ($match) {
                return $match.FullName
            }
        }
    }

    return $null
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

function Get-VcpkgToolchainPath {
    param(
        [Parameter(Mandatory)]
        [string]$Root
    )

    return Join-Path $Root "scripts\buildsystems\vcpkg.cmake"
}

function Initialize-VcpkgRoot {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [bool]$AllowBootstrap
    )

    $toolchain = Get-VcpkgToolchainPath $Root
    $vcpkgExe = Join-Path $Root "vcpkg.exe"

    if (-not (Test-Path -LiteralPath $Root)) {
        if (-not $AllowBootstrap) {
            throw "vcpkg was not found. Set VCPKG_ROOT or remove -NoBootstrapVcpkg to let this script clone a local copy."
        }

        $git = Resolve-Git
        if (-not $git) {
            throw "vcpkg was not found and git.exe is required to clone it. Install Git or set VCPKG_ROOT."
        }

        $parent = Split-Path -Parent $Root
        if ($parent) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }

        Write-Host "Cloning vcpkg into $Root"
        Invoke-Native $git @("clone", "--depth", "1", "https://github.com/microsoft/vcpkg.git", $Root)
    }

    if (-not (Test-Path -LiteralPath $toolchain)) {
        throw "The vcpkg toolchain file was not found at $toolchain."
    }

    if (-not (Test-Path -LiteralPath $vcpkgExe)) {
        if (-not $AllowBootstrap) {
            throw "vcpkg.exe was not found at $vcpkgExe. Remove -NoBootstrapVcpkg to bootstrap it."
        }

        $bootstrap = Join-Path $Root "bootstrap-vcpkg.bat"
        if (-not (Test-Path -LiteralPath $bootstrap)) {
            throw "Could not find $bootstrap."
        }

        Write-Host "Bootstrapping vcpkg at $Root"
        Invoke-Native $bootstrap @("-disableMetrics") -WorkingDirectory $Root
    }

    return $Root
}

function Resolve-VcpkgRoot {
    $candidates = New-Object "System.Collections.Generic.List[string]"

    function Add-Candidate {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) {
            return
        }

        $fullPath = ConvertTo-FullPath $Path
        if (-not $candidates.Contains($fullPath)) {
            $candidates.Add($fullPath) | Out-Null
        }
    }

    Add-Candidate $VcpkgRoot
    Add-Candidate $env:VCPKG_ROOT
    Add-Candidate $env:VCPKG_INSTALLATION_ROOT
    Add-Candidate (Join-Path $repoRoot ".deps\vcpkg")
    Add-Candidate "C:\vcpkg"
    Add-Candidate (Join-Path $env:USERPROFILE "vcpkg")
    Add-Candidate (Join-Path $env:USERPROFILE "source\vcpkg")

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath (Get-VcpkgToolchainPath $candidate)) {
            return Initialize-VcpkgRoot -Root $candidate -AllowBootstrap (-not $NoBootstrapVcpkg)
        }
    }

    $bootstrapRoot = if ($VcpkgRoot) {
        ConvertTo-FullPath $VcpkgRoot
    } else {
        Join-Path $repoRoot ".deps\vcpkg"
    }

    return Initialize-VcpkgRoot -Root $bootstrapRoot -AllowBootstrap (-not $NoBootstrapVcpkg)
}

function Get-BuildSpecs {
    $names = switch ($Architecture) {
        "x64" { @("x64") }
        "x86" { @("x86") }
        "arm64" { @("arm64") }
        "all" {
            if ([Environment]::Is64BitOperatingSystem) {
                @("arm64", "x64", "x86")
            } else {
                @("x86")
            }
        }
        "both" {
            if ([Environment]::Is64BitOperatingSystem) {
                @((Get-NativeArchitecture), "x86") | Select-Object -Unique
            } else {
                @("x86")
            }
        }
    }

    foreach ($name in $names) {
        $platform = switch ($name) {
            "arm64" { "ARM64" }
            "x64" { "x64" }
            "x86" { "Win32" }
        }

        [pscustomobject]@{
            Name = $name
            CMakePlatform = $platform
            Triplet = "$name-windows"
            BuildDirectory = Join-Path $BuildRoot $name
            DistDirectory = Join-Path $DistRoot $name
        }
    }
}

function Copy-BuildOutput {
    param(
        [Parameter(Mandatory)]
        [object]$Spec
    )

    $sourceDirectory = Join-Path $Spec.BuildDirectory $Configuration
    $dllSource = Join-Path $sourceDirectory "comctl32v6hook.dll"
    $manifestSource = Join-Path $sourceDirectory "comctl32v6hook.manifest"

    if (-not (Test-Path -LiteralPath $dllSource)) {
        throw "Missing build output: $dllSource"
    }
    if (-not (Test-Path -LiteralPath $manifestSource)) {
        throw "Missing manifest next to DLL: $manifestSource"
    }

    New-Item -ItemType Directory -Path $Spec.DistDirectory -Force | Out-Null
    Copy-Item -LiteralPath $dllSource -Destination (Join-Path $Spec.DistDirectory "comctl32v6hook.dll") -Force
    Copy-Item -LiteralPath $manifestSource -Destination (Join-Path $Spec.DistDirectory "comctl32v6hook.manifest") -Force
}

function Write-Step {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host ""
    Write-Host "== $Message"
}

$BuildRoot = ConvertTo-FullPath $BuildRoot
$DistRoot = ConvertTo-FullPath $DistRoot

$cmake = Resolve-CMake
$generator = Resolve-CMakeGenerator $cmake
$vcpkgRootResolved = Resolve-VcpkgRoot
$toolchain = Get-VcpkgToolchainPath $vcpkgRootResolved
$specs = @(Get-BuildSpecs)
$nativeArchitecture = Get-NativeArchitecture

Write-Host "Repository : $repoRoot"
Write-Host "CMake      : $cmake"
Write-Host "Generator  : $generator"
Write-Host "vcpkg      : $vcpkgRootResolved"
Write-Host "Build root : $BuildRoot"
Write-Host "Dist root  : $DistRoot"
Write-Host "Config     : $Configuration"
Write-Host "Native arch: $nativeArchitecture"

foreach ($spec in $specs) {
    if ($Clean -and (Test-Path -LiteralPath $spec.BuildDirectory)) {
        Write-Step "Cleaning $($spec.Name)"
        Remove-Item -LiteralPath $spec.BuildDirectory -Recurse -Force
    }

    Write-Step "Configuring $($spec.Name)"
    Invoke-Native $cmake @(
        "-S", $repoRoot,
        "-B", $spec.BuildDirectory,
        "-G", $generator,
        "-A", $spec.CMakePlatform,
        "-DCMAKE_TOOLCHAIN_FILE=$toolchain",
        "-DVCPKG_TARGET_TRIPLET=$($spec.Triplet)"
    )

    Write-Step "Building $($spec.Name)"
    Invoke-Native $cmake @(
        "--build", $spec.BuildDirectory,
        "--config", $Configuration,
        "--parallel"
    )

    Copy-BuildOutput $spec
    Write-Host "Output     : $($spec.DistDirectory)"
}

Write-Host ""
Write-Host "Build complete."
