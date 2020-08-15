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

$wamProviders =
@"
{077b8c4a-e425-578d-f1ac-6fdf1220ff68}
{5836994d-a677-53e7-1389-588ad1420cc5}
{05f02597-fe85-4e67-8542-69567ab8fd4f}
{d0034f5e-3686-5a74-dc48-5a22dd4f3d5b}
{4DE9BC9C-B27A-43C9-8994-0915F1A5E24F}
{556045FD-58C5-4A97-9881-B121F68B79C5}
{EC3CA551-21E9-47D0-9742-1195429831BB}
{63b6c2d2-0440-44de-a674-aa51a251b123}
{4180c4f7-e238-5519-338f-ec214f0b49aa}
{EB65A492-86C0-406A-BACE-9912D595BD69}
{d49918cf-9489-4bf1-9d7b-014d864cf71f}
{7acf487e-104b-533e-f68a-a7e9b0431edb}
{4E749B6A-667D-4C72-80EF-373EE3246B08}
{bfed9100-35d7-45d4-bfea-6c1d341d4c6b}
{ac01ece8-0b79-5cdb-9615-1b6a4c5fc871}
{1941f2b9-0939-5d15-d529-cd333c8fed83}
{0001376b-930d-50cd-2b29-491ca938cd54}
{072665fb-8953-5a85-931d-d06aeab3d109}
{f6a774e5-2fc7-5151-6220-e514f1f387b6}
{a48e7274-bb8f-520d-7e6f-1737e9d68491}
{88cd9180-4491-4640-b571-e3bee2527943}
{833e7812-d1e2-5172-66fd-4dd4b255a3bb}
{30ad9f59-ec19-54b2-4bdf-76dbfc7404a6}
{d229987f-edc3-5274-26bf-82be01d6d97e}
{8cde46fc-ca33-50ff-42b3-c64c1c731037}
{25756703-e23b-4647-a3cb-cb24d473c193}
{569cf830-214c-5629-79a8-4e9b58ea24bc}
{8BFE6B98-510E-478D-B868-142CD4DEDC1A}
"@

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]$Text,
        [string]$Path = $Script:logPath
    )

    $currentTime = Get-Date
    $currentTimeFormatted = $currentTime.ToString("yyyy/MM/dd HH:mm:ss.fffffff(K)")

    if (-not $Script:logWriter) {
        # For the first time, open file & add header
        [IO.StreamWriter]$Script:logWriter = [IO.File]::AppendText($Path)
        $Script:logWriter.WriteLine("date-time,delta(ms),info")
    }

    [TimeSpan]$delta = 0;
    if ($Script:lastLogTime) {
        $delta = $currentTime.Subtract($Script:lastLogTime)
    }

    $Script:logWriter.WriteLine("$currentTimeFormatted,$($delta.TotalMilliseconds),$text")
    $Script:lastLogTime = $currentTime
}

function Close-Log {
    if ($Script:logWriter) {
        $Script:logWriter.Close()
        $Script:logWriter = $null
    }
}

function Start-WamTrace {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = 'wam.etl',
        $SessionName = 'WamTrace'
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    # Create a provider listing
    $providerFile = Join-Path $Path -ChildPath 'wam.prov'
    Set-Content $wamProviders -Path $providerFile -ErrorAction Stop

    if ($FileName -notlike "*%d*") {
        $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) + "_%d.etl"
    }
    $traceFile = Join-Path $Path -ChildPath $FileName

    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"
    $logmanCommand = "logman start trace $SessionName -pf `"$providerFile`" -o `"$traceFile`" -bs 128 -max 256 -mode `"$logFileMode`" -ets"
    $logmanResult = Invoke-Expression $logmanCommand

    if ($LASTEXITCODE -ne 0) {
        throw "logman failed to start. exit code = $LASTEXITCODE.`n$logmanResult"
    }
}

function Stop-WamTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'WamTrace'
    )

    Write-Verbose "Stopping WAM trace"
    $logmanResult = & logman stop $SessionName -ets

    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman failed to stop. exit code = $LASTEXITCODE.`n$logmanResult"
    }
}


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
    $officeInfo = Get-OfficeInfo -ErrorAction Stop
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
    $SessionName = $null
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
        $SessionName = "NetTrace"
    )

    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Stopping netsh trace")) {
        return
    }

    # Netsh session might not be found right after it started. So repeat with some delay (currently 1 + 2 + 3 = 6 seconds max).
    $maxRetry = 3
    $retry = 0
    $sessionFound = $false

    while ($retry -le $maxRetry -and -not $sessionFound) {
        if ($retry) {
            Write-Verbose "$SessionName was not found. Retrying after $retry seconds."
            Start-Sleep -Seconds $retry
        }

        $sessions = & logman -ets
        foreach ($session in $sessions) {
            if ($session -like "*$SessionName*") {
                if ($session -match "[a-z,A-Z,0-9,-]+") {
                    $SessionName = $Matches[0]
                }
                else {
                    throw "Found a Netsh session, but cannot extract the name"
                }
                $sessionFound = $true
                break
            }
        }

        ++$retry
    }

    if (-not $sessionFound){
        throw "Cannot find a netsh trace session"
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
        Write-Error "Failed to stop netsh trace. exit code = $LASTEXITCODE.`n$local:result"
        return
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
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        # Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $NETFileSystemAvailable = $true
    }
    catch {
        Write-Warning "System.IO.Compression.FileSystem wasn't found. Using alternate method"
    }

    if ($NETFileSystemAvailable -and $UseShellApplication -eq $false) {
        # Note: [System.IO.Compression.ZipFile]::CreateFromDirectory() fails when one or more files in the directory is locked.
        #[System.IO.Compression.ZipFile]::CreateFromDirectory($Path, $zipFilePath, [System.IO.Compression.CompressionLevel]::Optimal, $false)

        try {
            New-Item $zipFilePath -ItemType file | Out-Null

            $zipStream = New-Object System.IO.FileStream -ArgumentList $zipFilePath, ([IO.FileMode]::Open)
            $zipArchive = New-Object System.IO.Compression.ZipArchive -ArgumentList $zipStream, ([IO.Compression.ZipArchiveMode]::Create)

            $files = @(Get-ChildItem $Path -Recurse | Where-Object {-not $_.PSIsContainer})
            $count = 0

            foreach ($file in $files) {
                Write-Progress -Activity "Creating a zip file $zipFilePath" -Status "Adding $($file.FullName)" -PercentComplete (100 * $count / $files.Count)

                try {
                    $fileStream = New-Object System.IO.FileStream -ArgumentList $file.FullName, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::ReadWrite)
                    $zipEntry = $zipArchive.CreateEntry($file.FullName.Substring($Path.Length + 1))
                    $zipEntryStream = $zipEntry.Open()
                    $fileStream.CopyTo($zipEntryStream)

                    ++$count
                }
                catch {
                    Write-Error "Failed to add $($file.FullName). $_"
                }
                finally {
                    if ($local:fileStream) {
                        $fileStream.Dispose()
                    }

                    if ($local:zipEntryStream) {
                        $zipEntryStream.Dispose()
                    }
                }
            }
        }
        finally {
            if ($local:zipArchive) {
                $zipArchive.Dispose()
            }

            if ($local:zipStream) {
                $zipStream.Dispose()
            }

            Write-Progress -Activity "Creating a zip file $zipFilePath" -Completed
        }
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
            ZipFilePath = $zipFilePath.ToString()
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
    $logs += (wevtutil el) -like '*AAD*'

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
            $patches = Get-ChildItem -Path Registry::$($key.Name) | Where-Object {$_.PSChildName -eq 'Patches' -and $_.SubKeyCount -gt 0} | Get-ChildItem | Get-ItemProperty

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
        }
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

function Save-CachedAutodiscover {
    [CmdletBinding()]
    param(
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    # Check %LOCALAPPDATA%\Microsoft\Outlook
    $localAppdata = [System.Environment]::ExpandEnvironmentVariables('%LOCALAPPDATA%')
    $cachePath = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook'
    if (-not (Test-Path $cachePath)) {
        return
    }

    # Get Autodiscover XML files and copy them to Path
    $files = @(Get-ChildItem $cachePath -Filter *Autod*.xml -Force -Recurse)
    $files | Copy-Item -Destination $Path

    # Remove Hidden attribute
    foreach ($file in @(Get-ChildItem $Path -Force)) {
        if ((Get-ItemProperty $file.FullName).Attributes -band [IO.FileAttributes]::Hidden) {
            Set-ItemProperty $file.Fullname -Name Attributes -Value ((Get-ItemProperty $file.FullName).Attributes -bxor [IO.FileAttributes]::Hidden)
        }

        # Unfortunately, this does not work in PowerShellV2.
        # (Get-ItemProperty $file.FullName).Attributes -= 'Hidden'
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

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    if ([IO.Path]::GetExtension($TargetProcess)  -ne 'exe') {
        $TargetProcess = [IO.Path]::GetFileNameWithoutExtension($TargetProcess) + ".exe"
    }

    # Create a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing
    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing"
    if (-not (Test-Path $keypath)) {
        New-Item (Split-Path $keypath) -Name 'tracing' -ErrorAction SilentlyContinue | Out-Null
    }

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
        # no return here intended.
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

    # If MS Office is not installed, bail.
    $officeInfo = Get-OfficeInfo -ErrorAction SilentlyContinue
    if (-not $officeInfo) {
        Write-Error "It seems that Microsoft Office (Microsoft 365 Apps) is not installed."
        return
    }

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
            # ignore errs here.
            $($o = Get-ChildItem -Path $officePath\* -Include *.dll,*.exe -Recurse) 2>&1 | Out-Null
            $o
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

    $processName = "msinfo32.exe"
    $process = $null

    try {
        $process = Start-Process $processName -ArgumentList "/nfo $filePath" -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            Write-Error "$processName failed. exit code = $($process.ExitCode)"
        }
    }
    catch {
        Write-Error "Failed to start $processName.`n$_"
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }
}

function Start-CAPITrace {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $SessionName = 'CapiTrace'
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $traceFile = Join-Path $Path -ChildPath 'capi_%d.etl'
    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"
    $logmanResult = Invoke-Expression "logman create trace $SessionName -ow -o `"$traceFile`" -p `"Security: SChannel`" 0xffffffffffffffff 0xff -bs 1024 -mode `"$logFileMode`" -max 256 -ets"

    if ($LASTEXITCODE -ne 0) {
        throw "logman failed to create a session. exit code = $LASTEXITCODE. $logmanResult"
        return
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
            Write-Progress -Activity "Downloading FiddlerCap" -Status "Please wait" -PercentComplete -1
            $webClient.DownloadFile($fiddlerCapUrl, $fiddlerSetupFile)
            $fiddlerCapDownloaded = $true
        }
        catch {
            throw "Failed to download FiddlerCapSetup from $fiddlerCapUrl. $_"
        }
        finally {
            Write-Progress -Activity "Downloading FiddlerCap" -Status "Done" -Completed

            if ($local:webClient) {
                $webClient.Dispose()
            }
        }

        # This is to exit the execution even when the caller uses "-ErrorAction SilentlyContinue". In this case, "throw" does not terminiate.
        if (-not $fiddlerCapDownloaded) {
            return
        }

        # Silently extract. Path must be absolute.
        $process = $null
        $err = $($process = Start-Process $fiddlerSetupFile -ArgumentList "/S /D=$fiddlerPath" -Wait -PassThru) 2>&1

        try {
            if ($process.ExitCode -ne 0) {
                throw "Failed to extract $fiddlerExe. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
            }
        }
        finally {
            if ($process) {
                $process.Dispose()
            }
        }
    }

    # Start FiddlerCap.exe
    $process = $null
    $err = $($process = Start-Process $fiddlerExe -PassThru) 2>&1
    try {
        if (-not $process -or $process.HasExited) {
            throw "FiddlerCap failed to start or prematurely exited. $(if ($null -ne $process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    New-Object PSCustomObject -Property @{
        FiddlerPath = $fiddlerPath
    }
}

function Start-Procmon {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $PmlFileName = "Procmon.pml",
        $ProcmonSearchPath # Look for existing procmon.exe before downloading
    )

    # Explicitly check admin rights
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Please run as administrator."
        return
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path

    # Search procmon.exe or procmon64.exe under $Path (including subfolders).
    if ($ProcmonSearchPath -and (Test-Path $ProcmonSearchPath)) {
        $findResult = @(Get-ChildItem -Path $ProcmonSearchPath -Filter 'procmon*.exe' -Recurse)
        if ($findResult.Count -ge 1) {
            $procmonFile = $findResult[0].FullName
            if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
                $procmon64 = $findResult | Where-Object {$_.Name -eq 'procmon64.exe'} | Select-Object -First 1
                if ($procmon64) {
                    $procmonFile = $procmon64.FullName
                }
            }
        }
    }

    $procmonZipDownloaded = $false

    if (-not ($local:procmonFile -and (Test-Path $local:procmonFile))) {

        # If 'ProcessMonitor.zip' isn't there, download.
        $procmonDownloadUrl = 'https://download.sysinternals.com/files/ProcessMonitor.zip'
        $procmonFolderPath = Join-Path $Path -ChildPath 'procmon_temp'
        $procmonZipFile = Join-Path $procmonFolderPath -ChildPath 'ProcessMonitor.zip'

        if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
            $procmonFile = Join-Path $procmonFolderPath -ChildPath 'Procmon64.exe'
        }
        else {
            $procmonFile = Join-Path $procmonFolderPath -ChildPath 'Procmon.exe'
        }

        if (-not (Test-Path $procmonFolderPath)) {
            New-Item $procmonFolderPath -ItemType Directory -ErrorAction Stop | Out-Null
        }

        if (-not (Test-Path $procmonZipFile)) {
            Write-Progress -Activity "Downloading procmon from $procmonDownloadUrl" -Status "Please wait" -PercentComplete -1

            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($procmonDownloadUrl, $procmonZipFile)
                $procmonZipDownloaded = $true
            }
            catch {
                throw "Failed to download procmon from $procmonDownloadUrl. $_"
            }
            finally {
                if ($local:webClient) {
                    $webClient.Dispose()
                }

                Write-Progress -Activity "Downloading procmon from $procmonDownloadUrl" -Status "Done" -Completed
            }
        }

        # This is just in case the caller used "-ErrorAction SilientlyContinue".
        if (-not $procmonZipDownloaded) {
            return
        }

        # Unzip ProcessMonitor.zip
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            $NETFileSystemAvailable = $true
        }
        catch {
            Write-Warning "System.IO.Compression.FileSystem isn't found. Using alternate method"
        }

        if ($NETFileSystemAvailable) {
            # .NET 4 or later
            [System.IO.Compression.ZipFile]::ExtractToDirectory($procmonZipFile, $procmonFolderPath)
        }
        else {
            # Use Shell.Application COM
            # see https://docs.microsoft.com/en-us/previous-versions/windows/desktop/sidebar/system-shell-folder-copyhere
            $shell = New-Object -ComObject Shell.Application
            $shell.NameSpace($procmonFolderPath).CopyHere($shell.NameSpace($procmonZipFile).Items(), 4)
        }
    }

    if (-not $PmlFileName.EndsWith('.pml')) {
        $PmlFileName = "$PmlFileName.pml"
    }

    $pmlFile = Join-Path $Path -ChildPath $PmlFileName

    # Start procmon.exe or procmon64.exe depending on the native arch.
    $process = $null
    $err = $($process = Start-Process $procmonFile -ArgumentList "/AcceptEula /Minimized /Quiet /NoFilter /BackingFile `"$pmlFile`"" -PassThru) 2>&1

    try {
        if (-not $process -or $process.HasExited) {
            throw "procmon failed to start or prematurely exited. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    Write-Verbose "Procmon successfully started"
    New-Object PSObject -Property @{
        ProcmonPath = $procmonFile
        ProcmonProcessId = $process.Id
        PMLFile = $pmlFile
        ProcmonZipDownloaded = $procmonZipDownloaded
        ProcmonFolderPath = $procmonFolderPath
    }
}

function Stop-Procmon {
    [CmdletBinding()]
    param(
    )

    $process = @(Get-Process -Name Procmon*)
    if ($process.Count -eq 0) {
        throw "Cannot find procmon or procmon64"
    }

    # Stop procmon
    Write-Progress -Activity "Stopping procmon" -Status "Please wait" -PercentComplete -1
    $procmonFile = $process[0].Path
    $process = $null
    $err = $($process = Start-Process $procmonFile -ArgumentList "/Terminate" -Wait -PassThru) 2>&1

    try {
        if ($process.ExitCode -ne 0) {
            Write-Error "procmon failed to stop. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    Write-Progress -Activity "Stopping procmon" -Status "Done" -Completed
}

function Start-TcoTrace {
    [CmdletBinding()]
    param(
    )

    $officeInfo = Get-OfficeInfo -ErrorAction Stop
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

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    $officeInfo = Get-OfficeInfo -ErrorAction Stop
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
        foreach ($install in @(Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall) + @(Get-ChildItem HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall)){
            $prop = Get-ItemProperty $install.PsPath
            if (($prop.DisplayName -like "Microsoft Office*" -or $prop.DisplayName -like "Microsoft 365 Apps*") -and $prop.DisplayIcon -and $prop.ModifyPath -notlike "*MUI*") {
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
        Write-Error "Microsoft Office is not installed"
        return
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

function Add-WerDumpKey {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [string]$TargetProcess, # Target Process (e.g. Outlook.exe)
        [parameter(Mandatory=$true)]
        $Path # Folder to save dump files
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = (Resolve-Path $Path -ErrorAction Stop).Path

    # Check if $TargetProcess ends with ".exe".
    if (-not $TargetProcess.EndsWith(".exe")) {
        Write-Verbose "$TargetProcess does not end with '.exe'.  Adding '.exe'"
        $TargetProcess += '.exe'
    }

    # Create a key 'LocalDumps' under HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps, if it doesn't exist
    $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'
    if (-not(Test-Path (Join-Path $werKey 'LocalDumps'))) {
        New-Item $werKey -Name 'LocalDumps' -ErrorAction Stop | Out-Null
    }

    # Create a ProcessName key under LocalDumps, if it doesn't exist
    $localDumpsKey = Join-Path $werKey 'LocalDumps'
    if (-not (Test-Path (Join-Path $localDumpsKey $TargetProcess))) {
        New-Item $localDumpsKey -Name $TargetProcess -ErrorAction Stop | Out-Null
    }

    # Create "CustomDumpFlags", "DumpType", and "DumpFolder" values in ProcessName key
    $ProcessKey = Join-Path $localDumpsKey $TargetProcess
    # -Force will overwrite existing value
    # 0x61826 = MiniDumpWithTokenInformation | MiniDumpIgnoreInaccessibleMemory | MiniDumpWithThreadInfo (0x1000) | MiniDumpWithFullMemoryInfo (0x800) |MiniDumpWithUnloadedModules (0x20) | MiniDumpWithHandleData (4)| MiniDumpWithFullMemory (2)
    New-ItemProperty $ProcessKey -Name 'CustomDumpFlags' -Value 0x00061826 -Force -ErrorAction Stop | Out-Null
    New-ItemProperty $ProcessKey -Name 'DumpType' -Value 0 -PropertyType DWORD -Force -ErrorAction Stop | Out-Null
    New-ItemProperty $ProcessKey -Name 'DumpFolder' -Value $Path -PropertyType String -Force -ErrorAction Stop | Out-Null

    # Rename DW Installed keys to "_Installed" to temporarily disable
    $pcHealth = @(
        # For C2R
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'

        # For MSI
        'HKLM:\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
        'HKLM:\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'
    )

    foreach ($installedKey in $pcHealth) {
        if (Test-Path $installedKey) {
            Rename-Item $installedKey -NewName '_Installed'
        }
    }
}

function Remove-WerDumpKey {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        [string]$TargetProcess # Target Process (e.g. Outlook.exe)
    )

    # Check if $TargetProcess ends with ".exe".
    if (-not $TargetProcess.EndsWith(".exe")) {
        Write-Verbose "$TargetProcess does not end with '.exe'.  Adding '.exe'"
        $TargetProcess += '.exe'
    }

    $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'
    $localDumpsKey = Join-Path $werKey 'LocalDumps'
    $ProcessKey = Join-Path $localDumpsKey $TargetProcess

    if (Test-Path $ProcessKey) {
        Remove-Item $ProcessKey
    }
    else {
        Write-Error "$ProcessKey does not exist"
    }

    # Rename DW "_Installed" keys back to "Installed"
    $pcHealth = @(
        # For C2R
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'

        # For MSI
        'HKLM:\Software\Microsoft\PCHealth\ErrorReporting\DW\Installed'
        'HKLM:\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\Installed'
    )

    foreach ($installedKey in $pcHealth) {
        if (Test-Path $installedKey) {
            Rename-Item $installedKey -NewName 'Installed'
        }
    }
}

function Start-WfpTrace {
    [CmdletBinding()]
    param(
    [Parameter(Mandatory = $true)]
    $Path,
    [Parameter(Mandatory = $true)]
    [int]$InternalSeconds,
    [TimeSpan]$MaxDuration = [TimeSpan]::FromHours(1)  # Just for safety, make sure to stop after a period
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType directory $Path -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    $job = Start-Job -ScriptBlock {
        param($Path, $InternalSeconds, $MaxDuration)

        $expiration = [DateTime]::Now.Add($MaxDuration)

        while ($true) {
            if ([DateTime]::Now -gt $expiration) {
                Write-Output "WfpTrace expired after $MaxDuration"
                break
            }

            # dump filters
            $filterFilePath = Join-Path $Path "filters_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
            netsh wfp show filters file=$filterFilePath verbose=on | Out-Null

            # dump netevents
            $eventFilePath = Join-Path $Path "netevents_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
            netsh wfp show netevents file="$eventFilePath" <#timewindow=$InternalSeconds#> | Out-Null
            Start-Sleep -Seconds $InternalSeconds

        }
    } -ArgumentList $Path, $InternalSeconds, $MaxDuration

    $job
}

function Stop-WfpTrace {
    [CmdletBinding()]
    [Parameter(Mandatory = $true)]
    param (
    $WfpJob
    )

    Stop-Job -Job $WfpJob
    Remove-Job -Job $WfpJob
}


function Save-Dump {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path, # Folder to save a dump file
        [parameter(Mandatory = $true)]
        $ProcessId
    )

    # Define a class to import MiniDumpWriteDump.
   <#
    The Native signature:
    see https://docs.microsoft.com/en-us/windows/win32/api/minidumpapiset/nf-minidumpapiset-minidumpwritedump

    BOOL MiniDumpWriteDump(
    HANDLE                            hProcess,
    DWORD                             ProcessId,
    HANDLE                            hFile,
    MINIDUMP_TYPE                     DumpType,
    PMINIDUMP_EXCEPTION_INFORMATION   ExceptionParam,
    PMINIDUMP_USER_STREAM_INFORMATION UserStreamParam,
    PMINIDUMP_CALLBACK_INFORMATION    CallbackParam
    );
    #>
    $Dbghelp = @"
    using System;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using Microsoft.Win32.SafeHandles;

    public class DbgHelp
    {
        [DllImport("Dbghelp.dll", SetLastError=true)]
        public static extern bool MiniDumpWriteDump(
            IntPtr hProcess,
            uint ProcessId,
            IntPtr hFile,
            uint DumpType,
            IntPtr ExceptionParam,
            IntPtr UserStreamParam,
            IntPtr CallbackParam);

        // Wrapper to return GetLastError() in case of failure (HRESULT)
        public static int MiniDumpWriteDumpWrapper(
            IntPtr hProcess,
            uint ProcessId,
            IntPtr hFile,
            uint DumpType,
            IntPtr ExceptionParam,
            IntPtr UserStreamParam,
            IntPtr CallbackParam)
        {
            if (MiniDumpWriteDump(hProcess, ProcessId, hFile, DumpType, ExceptionParam, UserStreamParam, CallbackParam))
            {
                return 0;
            }
            else
            {
                return Marshal.GetLastWin32Error();
            }
        }
    }
"@

    if (-not ('DbgHelp' -as [type])) {
        Add-Type -TypeDefinition $Dbghelp
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path

    try{
        $process = Get-Process -Id $ProcessId -ErrorAction Stop

        $dumpFile = Join-Path $Path "$($process.Name)_$(Get-Date -Format 'yyyy-MM-dd-HHmmss').dmp"
        $dumpFileStream = [System.IO.File]::Create($dumpFile)

        Write-Progress -Activity "Saving a memory dump file $dumpFile" -Status "Please wait" -PercentComplete -1
        # Note: 0x61826 = MiniDumpWithTokenInformation | MiniDumpIgnoreInaccessibleMemory | MiniDumpWithThreadInfo (0x1000) | MiniDumpWithFullMemoryInfo (0x800) |MiniDumpWithUnloadedModules (0x20) | MiniDumpWithHandleData (4) | MiniDumpWithFullMemory (2)
        $hr = [Dbghelp]::MiniDumpWriteDumpWrapper($process.Handle,$ProcessId,$dumpFileStream.Handle,0x61826,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero)

        if ($hr -eq 0) {
            New-Object PSObject -Property @{DumpFile = $dumpFile;}
        }
        else {
            Write-Error $("Failed to save a memory dump of $Process. Error = 0x{0:x}" -f $hr)
        }
    }
    finally {
        Write-Progress -Activity "Saving a memory dump file $dumpFile" -Status "Done" -Completed

        if ($dumpFileStream) {
            $dumpFileStream.Close()
        }
    }
}

function Save-MSIPC {
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true)]
        $Path
    )

    if (-not (Test-Path $Path -ErrorAction Stop)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    # MSIPC info is in %LOCALAPPDATA%\Microsoft\MSIPC
    $msipcPath = [Environment]::ExpandEnvironmentVariables('%LOCALAPPDATA%\Microsoft\MSIPC')

    if (-not (Test-Path $msipcPath)) {
        Write-Error "$msipcPath does not exist"
        return
    }

    # Copy only folders (i.e. skip drm files)
    # Copy-Item (Join-Path $msipcPath '*') -Destination $Path -Recurse

    # gci -Directory is only available for PowerShell V3 and above. To support PowerShell V2 clients, Where-Object is used here.
    foreach ($folder in @(Get-ChildItem $msipcPath | Where-Object {$_.PSIsContainer})) {
        $dest = Join-Path $Path $folder.Name

        if (-not (Test-Path $dest -ErrorAction Stop)){
            New-Item -ItemType Directory $dest -ErrorAction Stop | Out-Null
        }

        Copy-Item (Join-Path $folder.FullName '*') -Destination $dest -Recurse
    }
}

function Collect-OutlookInfo {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [parameter(Mandatory=$true)]
        $Path,
        [parameter(Mandatory=$true)]
        [ValidateSet('Outlook', 'Netsh', 'PSR', 'LDAP', 'CAPI', 'Configuration', 'Fiddler', 'TCO', 'Dump', 'CrashDump', 'Procmon', 'WAM', 'WFP', 'All')]
        [array]$Component,
        [switch]$SkipCabFile,
        [int]$DumpCount = 3,
        [int]$DumpIntervalSeconds = 60,
        [switch]$StartOutlook
    )

    # Explicitly check admin rights
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Please run as administrator."
        return
    }

    # MS Office must be installed to collect Outlook & TCO.
    # This is just a fail fast. Start-OutlookTrace/TCOTrace fail anyany.
    if ($Component -contains 'Outlook' -or $Component -contains 'TCO' -or $Component -contains 'All') {
        if (-not (Get-OfficeInfo -ErrorAction SilentlyContinue)) {
            throw "Component `"Outlook`" and/or `"TCO`" is specified, but Microsoft Office is not installed."
        }
    }

    if (-not (Test-Path $Path -ErrorAction Stop)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path
    $tempPath = Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())
    New-Item $tempPath -ItemType directory -ErrorAction Stop | Out-Null

    # $logFile = Join-Path $tempPath -ChildPath "log.txt"

    Write-Verbose "Starting traces"
    try {
        if ($Component -contains 'Configuration' -or $Component -contains 'All') {
            Write-Progress -Activity "Saving configuration" -Status "Please wait" -PercentComplete -1
            Save-EventLog -Path (Join-Path $tempPath 'EventLog')
            Save-MicrosoftUpdate -Path (Join-Path $tempPath 'Configuration')
            Save-OfficeRegistry -Path (Join-Path $tempPath 'Configuration') -ErrorAction SilentlyContinue
            Save-OfficeModuleInfo -Path (Join-Path $tempPath 'Configuration') -ErrorAction SilentlyContinue
            Save-OSConfiguration -Path (Join-Path $tempPath 'Configuration')
            Save-CachedAutodiscover -Path (Join-Path $tempPath 'Cached AutoDiscover')
            Save-MSIPC -Path (Join-Path $tempPath 'MSIPC') -ErrorAction SilentlyContinue

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
            # When netsh trace is run for the first time, it does not capture packets (even with "capture=yes").
            # To workaround, netsh is started and stopped immediately.
            $tempNetshName = 'netsh_test'
            Start-NetshTrace -Path $tempPath -FileName "$tempNetshName.etl"
            Stop-NetshTrace -SkipCabFile
            Remove-Item (Join-Path $tempPath "$tempNetshName.etl") -Force -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $tempPath $tempNetshName) -Recurse -Force -ErrorAction SilentlyContinue

            Start-NetshTrace -Path (Join-Path $tempPath 'Netsh')
            $netshTraceStarted = $true
        }

        if ($Component -contains 'Outlook' -or $Component -contains 'All') {
            Start-OutlookTrace -Path (Join-Path $tempPath 'Outlook')
            $outlookTraceStarted = $true
        }

        if ($Component -contains 'PSR' -or $Component -contains 'All') {
            Start-PSR -Path $tempPath -ShowGUI
            $psrStarted = $true
        }

        if ($Component -contains 'LDAP' -or $Component -contains 'All') {
            Start-LDAPTrace -Path (Join-Path $tempPath 'LDAP') -TargetProcess 'Outlook.exe'
            $ldapTraceStarted = $true
        }

        if ($Component -contains 'CAPI' -or $Component -contains 'All') {
            Start-CAPITrace -Path (Join-Path $tempPath 'CAPI')
            $capiTraceStarted = $true
        }

        if ($Component -contains 'TCO' -or $Component -contains 'All') {
            Start-TCOTrace
            $tcoTraceStarted = $true
        }

        if ($Component -contains 'WAM' -or $Component -contains 'All') {
            Start-WamTrace -Path (Join-Path $tempPath 'WAM')
            $wamTraceStarted = $true
        }

        if ($Component -contains 'Procmon' -or $Component -contains 'All') {
            $procmonResult = Start-Procmon -Path (Join-Path $tempPath 'Procmon') -ProcmonSearchPath $Path
            $procmonStared = $true
        }

        if ($Component -contains 'WFP' -or $Component -contains 'All') {
            $wfpPath = Join-Path $tempPath -ChildPath 'WFP'
            New-Item $wfpPath -ItemType Directory | Out-Null
            $wfpJob = Start-WfpTrace -Path $wfpPath -InternalSeconds 15
            $wfpStarted = $true
        }

        if ($Component -contains 'CrashDump' -or $Component -contains 'All') {
            Add-WerDumpKey -Path $tempPath -TargetProcess 'Outlook.exe'
            $crashDumpStarted = $true
        }

        if ($Component -contains 'Dump') {
            $process = Get-Process -Name 'Outlook' -ErrorAction Stop

            for ($i = 0; $i -lt $DumpCount; $i++) {
                Write-Progress -Activity "Saving a memory dump of Outlook" -Status "($i/$DumpCount)" -PercentComplete ($i/$DumpCount*100)
                $dumpResult = Save-Dump -Path $tempPath -ProcessId $process.Id
                Write-Verbose "Saved dump file: $($dumpResult.DumpFile)"

                # If there are more dumps to save, wait.
                if ($i -lt ($DumpCount - 1)) {
                    $secondsRemaining = $DumpIntervalSeconds
                    while ($secondsRemaining -gt 0) {
                        Write-Progress -Activity "Waiting $DumpIntervalSeconds seconds till next dump" -Status "Waiting" -SecondsRemaining $secondsRemaining
                        Start-Sleep -Seconds 1
                        $secondsRemaining-=1
                    }
                }
            }
        }

        if ($StartOutlook) {
            # Does Outlook.exe already exist?
            $existingProcesss = Get-Process 'Outlook' -ErrorAction SilentlyContinue
            if ($existingProcesss) {
                Write-Warning "Outlook is already running. PID = $($existingProcesss.Id)"
                $response = Read-Host "Is it ok to close Outlook.exe and start again? [Yes|No]"
                if ($response -like "Y*") {
                    Stop-Process -InputObject $existingProcesss -Force
                }
                else {
                    throw "StartOutlook parameter is specified, but Outlook is already running.Stop Outlook and run again, or run without StartOutlook parameter."
                }
            }

            $process = $null
            $err = $($process = Start-Process 'Outlook.exe' -PassThru) 2>&1

            try {
                if (-not $process -or $process.HasExited) {
                    throw "StartOutlook parameter is specified, but Outlook failed to start or prematurely exited. $(if ($null -ne $process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
                }
                Write-Host "Outlook has started. PID = $($process.Id)" -ForegroundColor Green
            }
            finally {
                if ($process) {
                    $process.Dispose()
                }

            }
        }

        if ($netshTraceStarted -or $outlookTraceStarted -or $psrStarted -or $ldapTraceStarted -or $capiTraceStarted -or $tcoTraceStarted -or $fiddlerCapStarted -or $crashDumpStarted -or $procmonStared -or $wamTraceStarted -or $wfpStarted){
            Read-Host "Hit enter to stop"
        }
    }
    finally {
        Write-Progress -Activity 'Stopping' -Status "Please wait." -PercentComplete -1

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
            Stop-TcoTrace -Path (Join-Path $tempPath 'TCO')
        }

        if ($wamTraceStarted) {
            Stop-WamTrace
        }

        if ($procmonStared) {
            Stop-Procmon
            # Remove procmon
            if ($procmonResult -and $procmonResult.ProcmonZipDownloaded) {
                Remove-Item $procmonResult.ProcmonFolderPath -Force -Recurse
            }
        }

        if ($wfpStarted) {
            Stop-WfpTrace $wfpJob
        }

        if ($crashDumpStarted) {
            Remove-WerDumpKey -TargetProcess 'Outlook.exe'
        }

        if ($fiddlerCapStarted) {
            Write-Warning "Please stop FiddlerCap and save the capture manually."
        }

        Write-Progress -Activity 'Stopping' -Status 'Please wait.' -Completed
    }

    $zipFileName = "Outlook_$($env:COMPUTERNAME)_$(Get-Date -Format "yyyyMMdd_HHmmss")"
    Compress-Folder -Path $tempPath -ZipFileName $zipFileName -Destination $Path -RemoveFiles | Out-Null

    if (Test-Path $tempPath) {
        Remove-Item $tempPath -Force
    }

    Write-Host "The collected data is `"$(Join-Path $Path $zipFileName).zip`"" -ForegroundColor Green
    Invoke-Item $Path
}

Export-ModuleMember -Function Start-WamTrace, Stop-WamTrace, Start-OutlookTrace, Stop-OutlookTrace, Start-NetshTrace, Stop-NetshTrace, Start-PSR, Stop-PSR, Save-EventLog, Save-MicrosoftUpdate, Save-OfficeRegistry, Save-OSConfiguration, Get-ProxySetting, Save-CachedAutodiscover, Start-LdapTrace, Stop-LdapTrace, Save-OfficeModuleInfo, Save-MSInfo32, Start-CAPITrace, Stop-CapiTrace, Start-FiddlerCap, Start-Procmon, Stop-Procmon, Start-TcoTrace, Stop-TcoTrace, Get-OfficeInfo, Add-WerDumpKey, Remove-WerDumpKey, Start-WfpTrace, Stop-WfpTrace, Save-Dump, Save-MSIPC, Collect-OutlookInfo