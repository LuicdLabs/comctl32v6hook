#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [switch]$KeepFiles
)

$ErrorActionPreference = "Stop"

if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    throw "Run this uninstaller from 64-bit PowerShell so System32 is not redirected to SysWOW64."
}

if (-not ("NativeMethods.Kernel32" -as [type])) {
    Add-Type -Namespace NativeMethods -Name Kernel32 -MemberDefinition @"
        [System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint = "MoveFileExW", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        public static extern bool MoveFileEx(string lpExistingFileName, System.IntPtr lpNewFileName, int dwFlags);
"@
}

$MoveFileDelayUntilReboot = 0x00000004

function Clear-AppInitKey {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (Test-Path $Path) {
        New-ItemProperty -Path $Path -Name "AppInit_DLLs" -Value "" -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $Path -Name "LoadAppInit_DLLs" -Value 0 -PropertyType DWord -Force | Out-Null
    }
}

function Register-DeployedFileForRemovalAfterReboot {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ([NativeMethods.Kernel32]::MoveFileEx($fullPath, [System.IntPtr]::Zero, $MoveFileDelayUntilReboot)) {
        Write-Host "Scheduled removal after reboot: $fullPath"
    } else {
        $lastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "Could not schedule removal after reboot for $fullPath. MoveFileEx failed with Win32 error $lastError."
    }
}

$system32 = Join-Path $env:WINDIR "System32"
$syswow64 = Join-Path $env:WINDIR "SysWOW64"

Clear-AppInitKey "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows"
if ([Environment]::Is64BitOperatingSystem) {
    Clear-AppInitKey "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Windows"
}

if (-not $KeepFiles) {
    Register-DeployedFileForRemovalAfterReboot (Join-Path $system32 "comctl32v6hook.dll")
    Register-DeployedFileForRemovalAfterReboot (Join-Path $system32 "comctl32v6hook.manifest")

    if ([Environment]::Is64BitOperatingSystem) {
        Register-DeployedFileForRemovalAfterReboot (Join-Path $syswow64 "comctl32v6hook.dll")
        Register-DeployedFileForRemovalAfterReboot (Join-Path $syswow64 "comctl32v6hook.manifest")
    }
}

Write-Host "AppInit_DLLs entries were disabled."
if ($KeepFiles) {
    Write-Host "Deployment files were kept."
} else {
    Write-Host "Deployment files were scheduled for deletion on the next reboot."
}
Write-Host "Reboot Windows to complete file removal if the DLL was already loaded."
