<#
.SYNOPSIS
Builds the WamInterop C++ project with Visual Studio Build Tools.

.DESCRIPTION
Locates MSBuild through vswhere.exe, then builds WamInterop.vcxproj using the
requested configuration, platform, and target.

.EXAMPLE
.\Build.ps1

Builds Release|x64.

.EXAMPLE
.\Build.ps1 -Configuration Debug -Platform Win32 -Target Rebuild
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [ValidateSet('Win32', 'x64')]
    [string]$Platform = 'x64',

    [ValidateSet('Build', 'Rebuild', 'Clean')]
    [string]$Target = 'Build'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectPath = Join-Path $PSScriptRoot 'WamInterop\WamInterop.vcxproj'
if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) {
    throw "Project file was not found: $projectPath"
}

$vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswherePath -PathType Leaf)) {
    throw 'vswhere.exe was not found. Install Visual Studio Build Tools with the Desktop development with C++ workload.'
}

$msbuildPath = & $vswherePath `
    -latest `
    -products '*' `
    -requires Microsoft.Component.MSBuild Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
    -find 'MSBuild\**\Bin\MSBuild.exe' |
    Select-Object -First 1

if (-not $msbuildPath -or -not (Test-Path -LiteralPath $msbuildPath -PathType Leaf)) {
    throw 'MSBuild with the Visual C++ tools was not found. Install the Desktop development with C++ workload.'
}

Write-Host "MSBuild:      $msbuildPath"
Write-Host "Project:      $projectPath"
Write-Host "Configuration: $Configuration|$Platform"
Write-Host "Target:        $Target"

& $msbuildPath $projectPath `
    "/t:$Target" `
    "/p:Configuration=$Configuration" `
    "/p:Platform=$Platform" `
    '/m' `
    '/nologo' `
    '/verbosity:minimal'

if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed with exit code $LASTEXITCODE."
}
