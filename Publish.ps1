<#
.SYNOPSIS
Publish this Powershell module

.DESCRIPTION
This is a helper script to:
- Download dependent Nuget Packages
- Publish this PowerShell module to a given PS repo.

.EXAMPLE
.\Publish.ps1 -Repository MyRepo
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    # Name of a Repository
    [string]$Repository
)

<#
.SYNOPSIS
Publish this moudle
.EXAMPLE
Publish-MyModule -Repository MyRepo
#>
function Publish-MyModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        # Name of a Repository
        [string]$Repository
    )

    # Name of this PowerShell module is assumed to be the name of psd1 file.
    $moduleName = Get-ChildItem $PSScriptRoot -Filter '*.psd1' | Select-Object -ExpandProperty BaseName

    # Does the given Repo exist?
    if (-not ($repo = Get-PSRepository $Repository)) {
        # Write-Error "Cannot find a repository called $Repository. Run Get-PSRepository to see the registered repositories."
        return
    }

    # Make sure to import PowerShellGet BEFORE changing PSModulePath.
    Import-Module PowerShellGet

    # Set PSModulePath to here
    $savedPSModulePath = $env:PSModulePath
    $env:PSModulePath = $PSScriptRoot

    try {
        # Files to be excluded (Need relative path (relative to Module root, not current dir))
        $exclude = @(
            $PSCommandPath # This script
            Get-ChildItem $PSScriptRoot -Filter '*.md' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
            Get-ChildItem (Join-Path $PSScriptRoot '.vscode') -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
            Get-ChildItem (Join-Path $PSScriptRoot 'test') -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        ) | ForEach-Object { $_.SubString($PSScriptRoot.Length + 1) }

        Write-Verbose "Exclude: $($exclude -join ',')"

        Write-Progress -Activity "Publishing $moduleName" -Status "Please wait" -PercentComplete -1
        Publish-Module -name $moduleName -Repository $Repository -Exclude $exclude
        Write-Progress -Activity "Publishing $moduleName" -Completed

        [PSCustomObject]@{
            PublishLocation = $repo.PublishLocation
        }

        Invoke-Item $repo.PublishLocation
    }
    finally {
        # Restore PSModulePath
        $env:PSModulePath = $savedPSModulePath
    }
}

<#
.SYNOPSIS
Download the dependent Nuget Packages.
.EXAMPLE
Get-Dependency -NugetPackage 'MathNet.Numerics'
#>
function Get-Dependency {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        # Name of Nuget Packages
        [string[]]$NugetPackage,
        # Name of the destination folder to which dependencies will be downloaded.
        [string]$Destination = 'lib',
        # Force download even if the file already exists
        [switch]$Force
    )

    $libRoot = Join-Path $PSScriptRoot $Destination
    if (-not (Test-Path $libRoot)) {
        New-Item $libRoot -ItemType Directory -ErrorAction Stop
    }

    foreach ($library in $NugetPackage) {
        $libFilePath = Join-Path $libRoot "$library.dll"

        if (-not $Force) {
            # Check if the libary already exists.
            if (Test-Path $libFilePath) {
                Write-Verbose "$libFilePath already exists."
                continue
            }
        }

        # Download library
        $zipPath = Join-Path $libRoot 'temp.zip'
        Write-Progress -Activity "Downloading $library" -Status "Please wait" -PercentComplete -1
        Invoke-Command -ScriptBlock {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/$library" -OutFile $zipPath
        }
        Write-Progress -Activity "Downloading $library" -Completed

        # Expand it to a temp folder
        $tempPath = Join-Path $libRoot 'temp'
        Expand-Archive -Path $zipPath -DestinationPath $tempPath -Force

        # Extract DLL file (path is hardcoded for now)
        $modulePath = Join-Path $tempPath "lib\net461\$library.dll"
        Move-Item -Path $modulePath -Destination $libFilePath -Force

        # Clean up
        Remove-Item $zipPath
        Remove-Item $tempPath -Recurse -Force
    }
}

$nugetPackages = @('Microsoft.Identity.Client', 'Microsoft.Identity.Client.Extensions.Msal')
$destination = 'modules'
Get-Dependency -NugetPackage $nugetPackages -Destination $destination -ErrorAction Stop
Publish-MyModule @PSBoundParameters