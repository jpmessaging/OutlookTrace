<#
.NOTES
Copyright (c) 2020 Ryusuke Fujita

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

# Outlook's ETW pvoviders
$outlook2016Providers =
@"
"{9efff48f-728d-45a1-8001-536349a2db37}" 0xFFFFFFFFFFFFFFFF 64
"{f50d9315-e17e-43c1-8370-3edf6cc057be}" 0xFFFFFFFFFFFFFFFF 64
"{02fd33df-f746-4a10-93a0-2bc6273bc8e4}" 0x00000015         64
"{b866d7ae-7c99-4c20-aa98-278fc044fb98}" 0xFFFFFFFFFFFFFFFF 64
"{35150b03-b441-412b-ace4-291895289743}" 0xFFFFFFFFFFFFFFFF 64
"{d6dd4818-7123-4abb-ad96-b044c1387b49}" 0xFFFFFFFFFFFFFFFF 64
"{5457acd3-0d66-4e53-b9e0-3be59f8b9f9d}" 0xFFFFFFFFFFFFFFFF 64
"{02cac15f-d4be-400e-9127-d54982aa4ae9}" 0xFFFFFFFFFFFFFFFF 64
"{d8d0510d-3f14-4da9-a096-b9c7ad386da0}" 0xFFFFFFFFFFFFFFFF 64
"{aa8fa310-0939-4ce3-b9bb-ae05b2695110}" 0xFFFFFFFFFFFFFFFF 64
"{31c1f514-1937-40ce-b0bf-2db7cb1b6a17}" 0xFFFFFFFFFFFFFFFF 64
"{6b6b571b-f4e3-4fbb-a83f-0790d11d19ab}" 0xFFFFFFFFFFFFFFFF 64
"{c911b508-e06d-4f76-8835-ea1b78e2f66d}" 0xFFFFFFFFFFFFFFFF 64
"{f762ce39-ac6c-4e1c-b55f-0e11586e6d07}" 0xFFFFFFFFFFFFFFFF 64
"{691e1c12-2693-4d4a-852c-7478657bbe6e}" 0xFFFFFFFFFFFFFFFF 64
"{081f51e8-2528-44af-ad0b-6e2e5c7242ad}" 0xFFFFFFFFFFFFFFFF 64
"{284b8d30-4aa6-4a0f-9143-ce2e8e1f10f0}" 0xFFFFFFFFFFFFFFFF 64
"{265f23e0-615d-4082-8e17-ddcd7e6f7eb4}" 0xFFFFFFFFFFFFFFFF 64
"{11adbd74-7df2-4e8e-802b-b3bcbfd04a78}" 0xFFFFFFFFFFFFFFFF 64
"{287bf315-5a11-4b2f-b069-b761ade25a49}" 0xFFFFFFFFFFFFF7FF 64
"{0dae1c38-7bfb-4960-8ea5-54139b54b751}" 0xFFFFFFFFFFFFFFFF 64
"{13967ee5-6b23-4bcd-a496-1d788449a8cf}" 0xFFFFFFFFFFFFFFFF 64
"{31b56255-5883-4f3e-8350-d7d6d88a4908}" 0xFFFFFFFFFFFFFFFF 64
"{059b2f1f-fc6d-4236-8c06-4357a91b17a1}" 0xFFFFFFFFFFFFFFFF 64
"{ad58872e-4df6-4b26-9841-5d7887a1c7a5}" 0xFFFFFFFFFFFFFFFF 64
"{03b1de06-84f4-4fa7-ba4c-cc1b82b56004}" 0xFFFFFFFFFFFFFFFF 64
"{daf0b914-9c1c-450a-81b2-fea7244f6ffa}" 0xFFFFFFFFFFFFFFFF 64
"{bb00e856-a12f-4ab7-b2c8-4e80caea5b07}" 0xFFFFFFFFFFFFFFFF 64
"{a1b69d49-2195-4f59-9d33-bdf30c0fe473}" 0xFFFFFFFFFFFFFFFF 64
"{b4f150b4-67db-4742-8846-2cd7b16ee60e}" 0xFFFFFFFFFFFFFFFF 64
"{8736922d-e8b2-47eb-8564-23e77e728cf3}" 0x00000414         64
"@

$outlook2013Providers =
@"
"{284b8d30-4aa6-4a0f-9143-ce2e8e1f10f0}" 0xFFFFFFFFFFFFFFFF 64
"{02cac15f-d4be-400e-9127-d54982aa4ae9}" 0xFFFFFFFFFFFFFFFF 64
"{aa8fa310-0939-4ce3-b9bb-ae05b2695110}" 0xFFFFFFFFFFFFFFFF 64
"{6b6b571b-f4e3-4fbb-a83f-0790d11d19ab}" 0xFFFFFFFFFFFFFFFF 64
"{c911b508-e06d-4f76-8835-ea1b78e2f66d}" 0xFFFFFFFFFFFFFFFF 64
"{f762ce39-ac6c-4e1c-b55f-0e11586e6d07}" 0xFFFFFFFFFFFFFFFF 64
"{691e1c12-2693-4d4a-852c-7478657bbe6e}" 0xFFFFFFFFFFFFFFFF 64
"{11adbd74-7df2-4e8e-802b-b3bcbfd04a78}" 0xFFFFFFFFFFFFFFFF 64
"{287bf315-5a11-4b2f-b069-b761ade25a49}" 0xFFFFFFFFFFFFFFFF 64
"{265f23e0-615d-4082-8e17-ddcd7e6f7eb4}" 0xFFFFFFFFFFFFFFFF 64
"{31c1f514-1937-40ce-b0bf-2db7cb1b6a17}" 0xFFFFFFFFFFFFFFFF 64
"{d8d0510d-3f14-4da9-a096-b9c7ad386da0}" 0xFFFFFFFFFFFFFFFF 64
"{b9522d9f-e2cd-44d4-b567-0d5182060e55}" 0xFFFFFFFFFFFFFFFF 64
"{96991e14-71db-4799-a66c-270004757fd8}" 0xFFFFFFFFFFFFFFFF 64
"{8736922d-e8b2-47eb-8564-23e77e728cf3}" 0x0000014         64
"{464a42fb-36bd-4749-a67c-02138387138c}" 0xFFFFFFFFFFFFFFFF 64
"{02fd33df-f746-4a10-93a0-2bc6273bc8e4}" 0xFFFFFFFFFFFFFFFF 64
"@

$outlook2010Providers =
@"
"{f94cbe33-31c2-492d-9bf8-573beff84c94}" 0x0FB7FFEF 64
"{e3c8312d-b20c-4831-995e-5ec5f5522215}" 0x00124586 64
"@

function Start-OutlookTrace {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = 'outlook.etl',
        $SessionName = 'OutlookTrace'
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    Write-Verbose "Creating a provider listing according to the version"
    $providerFile = Join-Path $Path -ChildPath 'Office.prov'
    $officeInfo = Get-OfficeInfo
    $major = $officeInfo.Version.Split('.')[0] -as [int]

    switch ($major) {
        14 {Set-Content $outlook2010Providers -Path $providerFile -ErrorAction Stop; break}
        15 {Set-Content $outlook2013Providers -Path $providerFile -ErrorAction Stop; break}
        16 {Set-Content $outlook2016Providers -Path $providerFile -ErrorAction Stop; break}
        default {throw "Couldn't find the version from $_"}
    }

    Write-Verbose "Starting an ETW session"

    # In order to use EVENT_TRACE_FILE_MODE_NEWFILE, file name must contain "%d"
    if ($FileName -notlike "*%d*") {
        $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) + "_%d.etl"
    }
    $traceFile = Join-Path $Path -ChildPath $FileName

    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"
    $logmanCommand = "logman start trace $SessionName -pf `"$providerFile`" -o `"$traceFile`" -bs 128 -max 256 -mode `"$logFileMode`" -ets"

    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$logmanCommand)) {
        $logmanResult = Invoke-Expression $logmanCommand

        if ($LASTEXITCODE -ne 0) {
            throw "logman failed to start. exit code = $LASTEXITCODE.`n$logmanResult"
        }
    }
}

function Stop-OutlookTrace {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        $SessionName = 'OutlookTrace'
    )

    Write-Verbose "Stopping Outlook trace"
    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Stopping Outlook Trace")) {
       return
    }

    $logmanResult = & logman stop $SessionName -ets
    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman failed to stop. exit code = $LASTEXITCODE.`n$logmanResult"
    }
}

function Start-NetshTrace {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = 'nettrace-winhttp-webio.etl'
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    # Use "InternetClient_dbg" for Win10
    $win32os = Get-WmiObject win32_operatingsystem
    $osMajor = $win32os.Version.Split(".")[0] -as [int]
    if ($osMajor -ge 10) {
        $scenario = "InternetClient_dbg"
    }
    else {
        $scenario = "InternetClient"
    }

    # Win10's netsh supports sessionname parameter.
    # Without explicit session name, netsh creates "-NetTrace-***".  This prefix "-" prevents logman from stopping the session.
    if ($osMajor -ge 10) {
        $SessionName = "NetshTrace"
    }

    Write-Verbose "Clearing dns cache"
    & ipconfig /flushdns | Out-Null

    $traceFile = Join-Path $Path -ChildPath $FileName

    Write-Verbose "Starting netsh trace. $netshCommand"
    if ($SessionName) {
        $netshCommand = "netsh trace start sessionname=$SessionName scenario=$scenario capture=yes tracefile=`"$traceFile`" overwrite=yes maxSize=2000"
    }
    else {
        $netshCommand = "netsh trace start scenario=$scenario capture=yes tracefile=`"$traceFile`" overwrite=yes maxSize=2000"
    }

    if (-not ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, $netshCommand))) {
        return
    }

    $netshResult = Invoke-Expression $netshCommand
    if ($LASTEXITCODE -ne 0) {
        throw "netsh failed to start. exit code = $LASTEXITCODE.`n$netshResult"
    }
}

function Stop-NetshTrace {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [switch]$SkipCabFile,
        $SessionName
    )

    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Stopping netsh trace")) {
        return
    }

    if (-not $SessionName)  {
        # Find an etw session
        $sessions = & logman -ets
        foreach ($session in $sessions) {
            if ($session -like '*NetTrace*') {
                $SessionName = $session.Substring(0, $session.IndexOf(' '))
                break
            }
        }

        if (-not $SessionName){
            throw "Cannot find a netsh trace session"
        }
    }
    if ($SkipCabFile) {
        # Manually stop the session
        Write-Verbose "Stopping $SessionName"
        $result = & logman stop $SessionName -ets
    }
    else {
        Write-Progress -Activity "Stopping netsh trace" -Status "This might take a while" -PercentComplete -1

        # Win10 supports sessionname paramter.
        $win32os = Get-WmiObject win32_operatingsystem
        $osMajor = $win32os.Version.Split(".")[0] -as [int]
        if ($osMajor -ge 10) {
            # netsh's "sessionname" needs a prefix before "-NetTrace"
            $shortSessionName = $SessionName.Substring(0, $SessionName.IndexOf("-"))
            $result = & netsh trace stop sessionname=$shortSessionName
        }
        else {
            $result = & netsh trace stop
        }

        Write-Progress -Activity "Stopping netsh trace" -Status "Done" -Completed
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to stop netsh trace. exit code = $LASTEXITCODE.`n$result"
    }
}

function Start-PSR {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = "psr.zip",
        [switch]$ShowGUI
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    # File name must be ***.zip
    if ([IO.Path]::GetExtension($FileName) -ne ".zip"){
        $FileName = [IO.Path]::GetFileNameWithoutExtension($FileName) + '.zip'
    }

    if (-not ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Starting PSR'))) {
        return
    }

    # For Win7, maxsc is 100
    $maxScreenshotCount = 100

    $win32os = Get-WmiObject win32_operatingsystem
    $osMajor = $win32os.Version.Split(".")[0] -as [int]
    $osMinor = $win32os.Version.Split(".")[1] -as [int]

    if ($osMajor -gt 6 -or $osMajor -eq 6 -and $osMinor -ge 3) {
        $maxScreenshotCount = 300
    }

    $outputFile = Join-Path $Path -ChildPath $FileName
    if ($ShowGUI) {
        & psr /start /maxsc $maxScreenshotCount /maxlogsize 10 /output $outputFile /exitonsave 1
    }
    else {
        & psr /start /maxsc $maxScreenshotCount /maxlogsize 10 /output $outputFile /exitonsave 1 /gui 0
    }

    # PSR doesn't return anything even on failure. Check if process is spawned.
    $process = Get-Process -Name psr -ErrorAction SilentlyContinue
    if (-not $process) {
        throw "PSR failed to start"
    }
}

function Stop-PSR {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
    )

    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Stopping PSR")) {
        return
    }

    $process = Get-Process -Name psr -ErrorAction SilentlyContinue
    if (-not $process){
        Write-Error "There's no psr.exe process"
        return
    }

    & psr /stop

    Wait-Process -InputObject $process
}

function Compress-Folder {
    [CmdletBinding()]
    param(
        # Specifies a path to one or more locations.
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string]$Destination,
        [string]$ZipFileName,
        [switch]$IncludeDateTime,
        [switch]$RemoveFiles,
        [switch]$UseShellApplication
    )

    $Path = Resolve-Path $Path
    $zipFileNameWithouExt = [System.IO.Path]::GetFileNameWithoutExtension($ZipFileName)
    if ($IncludeDateTime) {
        $zipFileName = $zipFileNameWithouExt + "_" + "$(Get-Date -Format "yyyyMMdd_HHmmss").zip"
    }
    else {
        $zipFileName = "$zipFileNameWithouExt.zip"
    }

    # If Destination is not given, use %TEMP% folder.
    if (-not $Destination) {
        $Destination = $env:TEMP
    }

    if (-not (Test-Path $Destination)) {
        New-Item $Destination -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Destination = Resolve-Path $Destination
    $zipFilePath = Join-Path $Destination -ChildPath $zipFileName

    $NETFileSystemAvailable = $false

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $NETFileSystemAvailable = $true
    }
    catch {
        Write-Warning "System.IO.Compression.FileSystem wasn't found. Using alternate method"
    }

    if ($NETFileSystemAvailable -and $UseShellApplication -eq $false) {
        [System.IO.Compression.ZipFile]::CreateFromDirectory($Path, $zipFilePath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
    }
    else {
        # Use Shell.Application COM

        # Create a zip file manually
        $shellApp = New-Object -ComObject Shell.Application
        Set-Content $zipFilePath ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (Get-Item $zipFilePath).IsReadOnly = $false

        $zipFile = $shellApp.NameSpace($zipFilePath)

        # If target folder is empty, CopyHere() fails. So make sure it's not empty
        if (@(Get-ChildItem $Path).Count -gt 0) {
            # Start copying the whole and wait until it's done. CopyHere works asynchronously.
            $zipFile.CopyHere($Path)

            # Now wait and poll
            $inProgress = $true
            $delayMilliseconds = 200
            Start-Sleep -Milliseconds 3000
            [System.IO.FileStream]$file = $null
            while ($inProgress) {
                Start-Sleep -Milliseconds $delayMilliseconds

                try {
                    $file = [System.IO.File]::Open($zipFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
                    $inProgress = $false
                }
                catch [System.IO.IOException] {
                    Write-Debug $_.Exception.Message
                }
                finally {
                    if ($file) {
                        $file.Close()
                    }
                }
            }
        }
    }

    if (Test-Path $zipFilePath) {
        # If requested, remove zipped files
        if ($RemoveFiles) {
            Write-Verbose "Removing zipped files"
            Get-ChildItem $Path -Exclude $ZipFileName | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            $filesRemoved = $true
        }

        New-Object PSCustomObject -Property @{
            ZipFilePath = $zipFilePath #Join-Path $zipFilePath -ChildPath $zipFileName
            FilesRemoved = $filesRemoved -eq $true
        }
    }
    else {
        throw "Zip file wasn't successfully created at $zipFilePath"
    }
}

function Save-EventLog {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory=$true)]
        $Path
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType directory $Path | Out-Null
    }

    $logs = @("Application","System")

    foreach ($log in $logs) {
        $fileName = $log.Replace('/', '_') + '.evtx'
        $filePath = Join-Path $Path -ChildPath $fileName

        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,"evtutil epl $log $filePath /ow")) {
            wevtutil epl $log $filePath /ow
        }
    }
}

function Get-MicrosoftUpdate {
    [CmdletBinding()]
    param(
        [switch]$OfficeOnly,
        [switch]$AppliedOnly
    )

    # Constants
    # https://docs.microsoft.com/en-us/windows/desktop/api/msi/nf-msi-msienumpatchesexa
    $PatchState = @{
        1 = 'MSIPATCHSTATE_APPLIED'
        2 = 'MSIPATCHSTATE_SUPERSEDED'
        4 = 'MSIPATCHSTATE_OBSOLETED'
        8 = 'MSIPATCHSTATE_REGISTERED'
        15 = 'MSIPATCHSTATE_ALL'
    }

    $productsKey = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products

    if ($OfficeOnly) {
        $productsKey = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products | Where-Object {$_.Name -match "F01FEC"}
    }

    $result = @(
        foreach ($key in $productsKey)
        {
            $patches = Get-ChildItem -pa Registry::$($key.Name) | Where-Object {$_.PSChildName -eq 'Patches' -and $_.SubKeyCount -gt 0} | Get-ChildItem | Get-ItemProperty

            if (-not $patches) {
                continue
            }

            foreach ($patch in $patches) {
                # extract KB number
                $KB = $null
                if ($patch.MoreInfoURL -match 'https?://support.microsoft.com/kb/(?<KB>\d+)') {
                    $KB = $Matches['KB']
                }

                <#
                MsiGetPatchInfoExW
                https://docs.microsoft.com/en-us/windows/desktop/api/msi/nf-msi-msigetpatchinfoexw
                Returns "1" if this patch is currently applied to the product. Returns "2" if this patch is superseded by another patch. Returns "4" if this patch is obsolete. These values correspond to the constants the dwFilter parameter of MsiEnumPatchesEx uses.
                #>
                New-Object PSCustomObject -Property @{
                    DisplayName = $patch.DisplayName
                    KB = $KB
                    MoreInfoURL = $patch.MoreInfoURL
                    Installed = $patch.Installed
                    PatchState = $PatchState[$patch.State]
                }
            }
        } # end of foreach ($key in $productsKey)
    )

    if ($AppliedOnly) {
        $result = $result | Where-Object {$_.PatchState -eq 'MSIPATCHSTATE_APPLIED'}
    }

    $result
}

function Save-MicrosoftUpdate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $cmdletName = $PSCmdlet.MyInvocation.MyCommand.Name
    $name = $cmdletName.Substring($cmdletName.IndexOf('-') + 1)
    Get-MicrosoftUpdate | Export-Clixml -Depth 4 -Path $(Join-Path $Path -ChildPath "$name.xml")
}

function Save-OfficeRegistry {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $registryKeys = @(
        "HKCU\Software\Microsoft\Office",
        "HKCU\Software\Policies\Microsoft\Office",
        "HKLM\Software\Microsoft\Office",
        "HKLM\Software\PoliciesMicrosoft\Office")

    foreach ($key in $registryKeys) {
        $filePath = Join-Path $Path -ChildPath "$($key.Replace("\","_")).reg"
        if (Test-Path $filePath) {
            Remove-Item $filePath -Force
        }
        $err = $(reg export $key $filePath) 2>&1

        if ($LASTEXITCODE -ne 0) {
            # keys under Policies may not exist. So ignore.
            if ($key -notlike "*Policies*") {
                Write-Error "$key is not exported. exit code = $LASTEXITCODE. $err"
            }
        }
    }
}

function Save-OSConfiguration {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    Get-WmiObject -Class Win32_ComputerSystem | Export-Clixml -Path $(Join-Path $Path -ChildPath "Win32_ComputerSystem.xml")
    Get-WmiObject -Class Win32_OperatingSystem | Export-Clixml -Path $(Join-Path $Path -ChildPath "Win32_OperatingSystem.xml")
    Get-ProxySetting | Export-Clixml -Path $(Join-Path $Path -ChildPath "ProxySetting.xml")
}


function Get-ProxySetting {
    [CmdletBinding()]
    param(
    )

    # Get Users's Internet Settings
    $internetSettings = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"

    # Get WebProxy class to get IE config
    $webProxyDefault = [System.Net.WebProxy]::GetDefaultProxy()

    # Get Machine's Winhttp Settings
    $netshRaw = & netsh winhttp show proxy
    foreach ($line in $netshRaw){
        if ($line -match "Proxy Server\(s\)\s*:\s*(?<proxyServer>.*)") {
            $winHttpProxyServer = $Matches['proxyServer']
        }
        elseif ($line -match "Bypass List\s*:\s*(?<bypassList>.*)") {
            $winHttpBypassList = $Matches['bypassList']
        }
        elseif ($line -like "*Direct access*") {
            $winHttpDirectAccess = $true
        }
    }

    New-Object PSCustomObject -Property @{
        ProxyEnabled = $($internetSettings.ProxyEnable -eq 1)
        ProxyServer = $internetSettings.ProxyServer
        ProxyOverride = $internetSettings.ProxyOverride

        WebProxyDefault = $webProxyDefault

        WinHttpDirectAccess = $winHttpDirectAccess -eq $true
        WinHttpProxyServer = $winHttpProxyServer
        WinHttpBypassList = $winHttpBypassList
    }
}

function Start-LdapTrace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, HelpMessage = "Directory for output file")]
        $Path,
        [parameter(Mandatory=$true,
                   HelpMessage = "Process name to trace. e.g. Outlook.exe")]
        $TargetProcess,
        $SessionName = 'LdapTrace'
    )

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    if ([IO.Path]::GetExtension($TargetProcess)  -ne 'exe') {
        $TargetProcess = [IO.Path]::GetFileNameWithoutExtension($TargetProcess) + ".exe"
    }

    # Create a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing
    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing"
    New-Item $keypath -Name $TargetProcess -ErrorAction SilentlyContinue | Out-Null
    $key = Get-Item (Join-Path $keypath -ChildPath $TargetProcess)

    if (!$key) {
        throw "Failed to create the key under $keypath. Make sure to run as an administrator"
    }

    # Start ETW session
    $traceFile = Join-Path $Path -ChildPath "ldap_%d.etl"
    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"
    $logmanResult = Invoke-Expression "logman create trace $SessionName -ow -o `"$traceFile`" -p Microsoft-Windows-LDAP-Client 0x1a59afa3 0xff -bs 1024 -mode `"$logFileMode`" -max 256 -ets"

    if ($LASTEXITCODE -ne 0) {
        throw "Failed to start LDAP trace. exit code = $LASTEXITCODE. $logmanResult"
    }
}

function Stop-LdapTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'LdapTrace',
        $TargetProcess
    )

    $logmanResult = Invoke-Expression "logman stop $SessionName -ets"

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to stop LDAP trace. exit code = $LASTEXITCODE. $logmanResult"
    }

    # Remove a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing (ignore any errors)
    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing\$TargetProcess"
    Remove-Item $keypath -ErrorAction SilentlyContinue | Out-Null
}

function Save-OfficeModuleInfo {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $officeInfo = Get-OfficeInfo

    $officePaths = @(
        $officeInfo.InstallPath

        if ($env:CommonProgramFiles) {
            Join-Path $env:CommonProgramFiles 'microsoft shared'
        }

        if (${env:CommonProgramFiles(x86)}) {
            Join-Path ${env:CommonProgramFiles(x86)} 'microsoft shared'
        }
    )

    Write-Verbose "officePaths are $officePaths"

    # Get exe and dll
    if (-not $PSCmdlet.ShouldProcess($officePaths[0], "Exporting module info")) {
        return
    }

    $items = @(
        foreach ($officePath in $officePaths) {
            Get-ChildItem -Path $officePath\* -Include *.dll,*.exe -Recurse
        }
    )

    $result = @(
        foreach ($item in $items) {
            if ($item.VersionInfo.FileVersionRaw) {
                $fileVersion = $item.VersionInfo.FileVersionRaw
            }
            else {
                $fileVersion = $item.VersionInfo.FileVersion
            }

            New-Object PSCustomObject -Property @{
                Name = $item.Name
                FullName = $item.FullName
                #VersionInfo = $item.VersionInfo # too much info and FileVersionRaw is harder to find
                FileVersion = $fileVersion
            }
        }
    )

    $cmdletName = $PSCmdlet.MyInvocation.MyCommand.Name
    $name = $cmdletName.Substring($cmdletName.IndexOf('-') + 1)
    $result | Export-Clixml -Depth 4 -Path $(Join-Path $Path -ChildPath "$name.xml")
}

function Save-MSInfo32 {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path -ErrorAction Stop)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $filePath = Join-Path $Path -ChildPath "$($env:COMPUTERNAME).nfo"
    Start-Process "msinfo32.exe" -ArgumentList "/nfo $filePath" -Wait

    # It seems msinfo32.exe return 1 on success
    if ($LASTEXITCODE -ne 1) {
        Write-Error "msinfo32.exe failed. exit code = $LASTEXITCODE. $result"
    }
}

function Start-CAPITrace {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $SessionName = 'CapiTrace'
    )

    $traceFile = Join-Path $Path -ChildPath 'capi_%d.etl'
    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"
    $logmanResult = Invoke-Expression "logman create trace $SessionName -ow -o `"$traceFile`" -p `"Security: SChannel`" 0xffffffffffffffff 0xff -bs 1024 -mode `"$logFileMode`" -max 256 -ets"

    if ($LASTEXITCODE -ne 0) {
        throw "logman failed to create a session. exit code = $LASTEXITCODE. $logmanResult"
    }

    # Note: Depending on the OS version, not all providers are available.
    $logmanResult = Invoke-Expression "logman update trace $SessionName -p `"Schannel`" 0xffffffffffffffff 0xff -ets"
    $logmanResult = Invoke-Expression "logman update trace $SessionName -p `"{44492B72-A8E2-4F20-B0AE-F1D437657C92}`" 0xffffffffffffffff 0xff -ets"
    $logmanResult = Invoke-Expression "logman update trace $SessionName -p `"Microsoft-Windows-Schannel-Events`" 0xffffffffffffffff 0xff -ets"
}

function Stop-CapiTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'CapiTrace'
    )

    $logmanResult = Invoke-Expression "logman stop $SessionName -ets"
    if ($LASTEXITCODE -ne 0){
        Write-Error "failed to stop $SessionName. exit code = $LASTEXITCODE. $logmanResult"
    }
}

function Start-FiddlerCap {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path
    $fiddlerPath = Join-Path $Path -ChildPath "FiddlerCap"
    $fiddlerExe = Join-Path $fiddlerPath -ChildPath 'FiddlerCap.exe'

    # If FiddlerCap is not available, download Setup file and extract.
    if (-not (Test-Path $fiddlerExe)) {
        $fiddlerCapUrl = "https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe"
        $fiddlerSetupFile = Join-Path $Path -ChildPath 'FiddlerCapSetup.exe'

        try {
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($fiddlerCapUrl, $fiddlerSetupFile)
        }
        catch {
            throw "Failed to download FiddlerCapSetup from $fiddlerCapUrl. $_"
        }
        finally {
            if ($webClient) {
                $webClient.Dispose()
            }
        }

        # Silently extract. Path must be absolute.
        Invoke-Expression  "$fiddlerSetupFile /S /D=$fiddlerPath"
    }

    # Start FiddlerCap.exe
    $process = Start-Process $fiddlerExe -PassThru

    <#
    [PSCustomObject]@{
        Process = $process
        FiddlerPath = $fiddlerPath
    }
    #>
    New-Object PSCustomObject -Property @{
        Process = $process
        FiddlerPath = $fiddlerPath
    }
}

function Start-TcoTrace {
    [CmdletBinding()]
    param(
    )

    $officeInfo = Get-OfficeInfo
    $majorVersion = $officeInfo.Version.Split('.')[0]

    # Create registry key & values. Ignore errors (might fail due to existing values)
    $keypath = "HKCU:\Software\Microsoft\Office\$majorVersion.0\Common\Debug"
    if (-not (Test-Path $keypath)) {
        New-Item $keypath -ErrorAction Stop | Out-Null
    }

    New-ItemProperty $keypath -Name 'TCOTrace' -PropertyType DWORD -Value 7 -ErrorAction SilentlyContinue | Out-Null
    New-ItemProperty $keypath -Name 'MsoHttpVerbose' -PropertyType DWORD -Value 1 -ErrorAction SilentlyContinue | Out-Null

    # If failed, throw a terminating error
    Get-ItemProperty $keypath -Name 'TCOTrace' -ErrorAction Stop | Out-Null
    Get-ItemProperty $keypath -Name 'MsoHttpVerbose' -ErrorAction Stop | Out-Null
}

function Stop-TcoTrace {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path
    )

    $officeInfo = Get-OfficeInfo
    $majorVersion = $officeInfo.Version.Split('.')[0]

    # Remove registry values
    $keypath = "HKCU:\Software\Microsoft\Office\$majorVersion.0\Common\Debug"
    if (-not (Test-Path $keypath)) {
        Write-Warning "$keypath does not exist"
        return
    }

    Remove-ItemProperty $keypath -Name 'TCOTrace' -ErrorAction SilentlyContinue | Out-Null
    Remove-ItemProperty $keypath -Name 'MsoHttpVerbose' -ErrorAction SilentlyContinue | Out-Null

    # TCO Trace logs are in %TEMP%
    foreach ($item in @(Get-ChildItem -Path "$env:TEMP\*" -Include "office.log", "*.exe.log")) {
        Copy-Item $item -Destination $Path
    }
}

function Get-OfficeInfo {
    [CmdletBinding()]
    param()

    # Use the cache if it's available
    if ($Script:OfficeInfoCache) {
        return $Script:OfficeInfoCache
    }

    # There might be more than one version of Office installed.
    $officeInstallations = @(
    foreach ($install in @(Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall)){
        $prop = Get-ItemProperty $install.PsPath
        if ($prop.DisplayName -like "Microsoft Office*" -and $prop.DisplayIcon -and $prop.ModifyPath -notlike "*OMUI*") {
            New-Object PSObject -Property @{
                Version = $prop.DisplayVersion
                Location = $prop.InstallLocation
                DisplayName = $prop.DisplayName
                ModifyPath = $prop.ModifyPath
                DisplayIcon = $prop.DisplayIcon
            }
        }
    }
    )

    if (-not $officeInstallations) {
        throw "Microsoft Office is not installed"
    }

    # Use the latest
    $latestOffice = $officeInstallations | Sort-Object -Property Version -Descending | Select-Object -First 1

    $outlookReg = Get-ItemProperty HKLM:'\SOFTWARE\Clients\Mail\Microsoft Outlook' -ErrorAction Stop
    $mapiDll = Get-ItemProperty $outlookReg.DLLPathEx -ErrorAction Stop

    $Script:OfficeInfoCache =
    New-Object PSCustomObject -Property @{
        DisplayName = $latestOffice.DisplayName
        Version = $latestOffice.Version
        InstallPath = $latestOffice.Location
        MapiDllFileInfo = $mapiDll
    }

    $Script:OfficeInfoCache
}

function Collect-OutlookInfo {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [parameter(Mandatory=$true)]
        $Path,
        [parameter(Mandatory=$true)]
        [ValidateSet('Outlook', 'Netsh', 'PSR', 'LDAP', 'CAPI', 'Configuration','Fiddler', 'TCO', 'All')]
        [array]$Component,
        [switch]$SkipCabFile
    )

    if (-not (Test-Path $Path -ErrorAction Stop)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $tempPath = Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())
    New-Item $tempPath -ItemType directory -ErrorAction Stop | Out-Null

    Write-Verbose "Starting traces"
    try {
        if ($Component -contains 'Configuration' -or $Component -contains 'All') {
            Write-Progress -Activity "Saving configuration" -Status "Please wait" -PercentComplete -1
            Save-EventLog -Path $tempPath
            Save-MicrosoftUpdate -Path $tempPath
            Save-OfficeRegistry -Path $tempPath
            Save-OfficeModuleInfo -Path $tempPath
            Save-OSConfiguration -Path $tempPath
            Write-Progress -Activity "Saving configuration" -Status "Done" -Completed
            # Do we need MSInfo32?
            # Save-MSInfo32 -Path $tempPath
        }

        if ($Component -contains 'Fiddler' -or $Component -contains 'All') {
            Start-FiddlerCap -Path $Path | Out-Null
            $fiddlerCapStarted = $true

            Write-Warning "FiddlerCap has started. Please manually configure and start capture."
        }

        if ($Component -contains 'Netsh' -or $Component -contains 'All') {
            Start-NetshTrace -Path $tempPath
            $netshTraceStarted = $true
        }

        if ($Component -contains 'Outlook' -or $Component -contains 'All') {
            Start-OutlookTrace -Path $tempPath
            $outlookTraceStarted = $true
        }

        if ($Component -contains 'PSR' -or $Component -contains 'All') {
            Start-PSR -Path $tempPath -ShowGUI
            $psrStarted = $true
        }

        if ($Component -contains 'LDAP' -or $Component -contains 'All') {
            Start-LDAPTrace -Path $tempPath -TargetProcess 'Outlook.exe'
            $ldapTraceStarted = $true
        }

        if ($Component -contains 'CAPI' -or $Component -contains 'All') {
            Start-CAPITrace -Path $tempPath
            $capiTraceStarted = $true
        }

        if ($Component -contains 'TCO' -or $Component -contains 'All') {
            Start-TCOTrace
            $tcoTraceStarted = $true
        }

        if ($netshTraceStarted -or $outlookTraceStarted -or $psrStarted -or $ldapTraceStarted -or $capiTraceStarted -or $tcoTraceStarted -or $fiddlerCapStarted){
            Read-Host "Hit enter to stop tracing"
        }
    }
    finally {
        if ($psrStarted) {
            Stop-PSR
        }

        if ($netshTraceStarted) {
            Stop-NetshTrace -SkipCabFile:$SkipCabFile
        }

        if ($outlookTraceStarted) {
            Stop-OutlookTrace
        }

        if ($ldapTraceStarted) {
            Stop-LDAPTrace -TargetProcess 'Outlook.exe'
        }

        if ($capiTraceStarted) {
            Stop-CAPITrace
        }

        if ($tcoTraceStarted) {
            Stop-TcoTrace -Path $tempPath
        }

        if ($fiddlerCapStarted) {
            Write-Warning "Please stop FiddlerCap and save the capture manually."
        }
    }

    Write-Verbose "Compressing $tempPath"
    $zipFileName = "Outlook_$($env:COMPUTERNAME)_$(Get-Date -Format "yyyyMMdd_HHmmss")"

    Compress-Folder -Path $tempPath -ZipFileName $zipFileName -Destination $Path -RemoveFiles | Out-Null

    if (Test-Path $tempPath) {
        Remove-Item $tempPath -Force
    }

    Write-Host "The collected data is in `"$(Join-Path $Path $zipFileName).zip`"" -ForegroundColor Green
    Invoke-Item $Path
}

