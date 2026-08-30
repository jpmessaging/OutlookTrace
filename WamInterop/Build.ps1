<#
.SYNOPSIS
Builds the WamInterop with Visual Studio Build Tools.

.DESCRIPTION
Locates MSBuild through vswhere.exe, then builds WamInterop.slnx using the requested configuration, platform, and action

.EXAMPLE
.\Build.ps1

Builds Debug|x64.

.EXAMPLE
.\Build.ps1 -Configuration Release -Platform x64 -Action Rebuild
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',

    [ValidateSet('Win32', 'x64')]
    [string]$Platform = 'x64',

    [ValidateSet('Build', 'Rebuild', 'Clean')]
    [string]$Action = 'Build'
)

$slnPath = Join-Path $PSScriptRoot 'WamInterop.slnx'

if (-not (Test-Path -LiteralPath $slnPath -PathType Leaf)) {
    Write-Error "Cannot find the solution file: $slnPath"
    return
}

$vswherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'

if (-not (Test-Path -LiteralPath $vswherePath -PathType Leaf)) {
    Write-Error 'vswhere.exe was not found. Install Visual Studio Build Tools with the Desktop development with C++ workload.'
    return
}

$msbuildPath = & $vswherePath `
    -latest `
    -products '*' `
    -requires Microsoft.Component.MSBuild Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
    -find 'MSBuild\**\Bin\MSBuild.exe' |
    Select-Object -First 1

if (-not $msbuildPath -or -not (Test-Path -LiteralPath $msbuildPath -PathType Leaf)) {
    Write-Error 'MSBuild with the Visual C++ tools was not found. Install the Desktop development with C++ workload.'
    return
}

Write-Host "MSBuild       : $msbuildPath"
Write-Host "Solution      : $slnPath"
Write-Host "Configuration : $Configuration|$Platform"
Write-Host "Action        : $Action"

& $msbuildPath $slnPath `
    "-t:$Action" `
    "-p:Configuration=$Configuration" `
    "-p:Platform=$Platform" `
    '-maxCpuCount' `
    '-nologo' `
    '-verbosity:minimal'

if ($LASTEXITCODE -ne 0) {
    Write-Error "MSBuild failed with exit code $LASTEXITCODE."
}
