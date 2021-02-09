<#
.NOTES
Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

$Version = 'v2021-02-05'
#Requires -Version 3.0

# Outlook's ETW pvoviders
$outlook2016Providers =
@'
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
'@

$outlook2013Providers =
@'
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
'@

$outlook2010Providers =
@'
"{f94cbe33-31c2-492d-9bf8-573beff84c94}" 0x0FB7FFEF 64
"{e3c8312d-b20c-4831-995e-5ec5f5522215}" 0x00124586 64
'@

$wamProviders =
@'
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
'@


function Open-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if ($Script:logWriter) {
       Close-Log
    }

    # Open a file & add header
    try {
        [IO.StreamWriter]$Script:logWriter = [IO.File]::AppendText($Path)
        $Script:logWriter.WriteLine("date-time,thread_relative_delta(ms),thread,function,info")
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string]$Message,
        [Parameter(ValueFromPipeline=$true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    process {
        # Ignore null or an empty string.
        if (-not $Message) {
            return
        }

        # If ErrorRecord is provided, use it.
        if ($ErrorRecord) {
            $Message += " ;ScriptCallStack: $($ErrorRecord.ScriptStackTrace.Replace([Environment]::NewLine, ' '))"
        }

        # If Open-Log is not called beforehand, just output to verbose.
        if (-not $Script:logWriter) {
            Write-Verbose $Message
            return
        }

        $currentTime = Get-Date
        $currentTimeFormatted = $currentTime.ToString('o')

        # Delta time is relative to thread.
        # Each thread has it's own copy of lastLogTime now.
        [TimeSpan]$delta = 0;
        if ($Script:lastLogTime) {
            $delta = $currentTime.Subtract($Script:lastLogTime)
        }

        # Format as CSV:
        $sb = New-Object System.Text.StringBuilder
        $sb.Append($currentTimeFormatted).Append(',') | Out-Null
        $sb.Append($delta.TotalMilliseconds).Append(',') | Out-Null
        $sb.Append([System.Threading.Thread]::CurrentThread.ManagedThreadId).Append(',') | Out-Null
        $sb.Append((Get-PSCallStack)[1].Command).Append(',') | Out-Null
        $sb.Append('"').Append($Message.Replace('"', "'")).Append('"') | Out-Null

        # Protect from concurrent write
        [System.Threading.Monitor]::Enter($Script:logWriter)
        try {
            $Script:logWriter.WriteLine($sb.ToString())
        }
        finally {
            [System.Threading.Monitor]::Exit($Script:logWriter)
        }

        $sb = $null
        $Script:lastLogTime = $currentTime
    }
}

function Close-Log {
    if ($Script:logWriter) {
        Write-Log "Closing logWriter."
        $Script:logWriter.Close()
        $Script:logWriter = $null
        $Script:lastLogTime = $null
    }
}

<#
.SYNOPSIS
Create a runspace pool so that Start-Task commands can use it.
Make sure to call Close-TaskRunspace to dispose the runspace pool.
#>
function Open-TaskRunspace {
    [CmdletBinding()]
    param(
        # Maximum number of runspaces that pool creates
        $MaxRunspaces = $env:NUMBER_OF_PROCESSORS,
        # PowerShell modules to import to InitialSessionState.
        [string[]]$Modules,
        # Variable to import to InitialSessionState.
        [System.Management.Automation.PSVariable[]]$Variables,
        # Import all non-const script-scoped variables to InitialSessionState.
        [switch]$IncludeScriptVariables
    )

    if (-not $Script:runspacePool) {
        Write-Log "Setting up a Runspace with an initialSessionState. MaxRunspaces: $MaxRunspaces."
        $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        # Add functions from this script module. This will find all the functions including non-exported ones.
        # Note: I just want to call "ImportPSModule". It works, but emits "WARNING: The names of some imported commands ...".
        # Just to avoid this, I'm manually adding each command.
        #   $initialSessionState.ImportPSModule($MyInvocation.MyCommand.Module.Path)
        Get-Command -Module $MyInvocation.MyCommand.Module | ForEach-Object {
            $initialSessionState.Commands.Add($(
                New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry($_.Name, $_.ScriptBlock)
            ))
        }

        # Import extra modules.
        if ($Modules) {
            $initialSessionState.ImportPSModule($Modules)
        }

        # Import Script-scoped variable.
        if ($IncludeScriptVariables) {
            foreach ($_ in @(Get-Variable -Scope Script | Where-Object {$_.Options -notmatch 'Constant' -and $_.Value})) {
                $initialSessionState.Variables.Add($(
                    New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $_.Name, $_.Value, <# description #>$null
                ))
            }
        }

        # Import given variables
        foreach ($_ in $Variables) {
            $initialSessionState.Variables.Add($(
                New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $_.Name, $_.Value, <# description #>$null
            ))
        }

        $Script:runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxRunspaces, $initialSessionState, $Host)
        $Script:runspacePool.Open()

        Write-Log "RunspacePool ($($Script:runspacePool.InstanceId.ToString())) is Opened."
    }
}

function Close-TaskRunspace {
    [CmdletBinding()]
    param()

    $id = $Script:runspacePool.InstanceId.ToString()
    $Script:runspacePool.Close()
    $Script:runspacePool = $null
    Write-Log "RunspacePool ($id) is Closed."
}

<#
.SYNOPSIS
Start a task to run a command or scriptblock to run asynchronously.

.EXAMPLE
$t = Start-Task { Invoke-LongRunning }
if (Wait-Task $t -Timeout 00:01:00) {
    $t | Receive-Task
}
else {
    Write-Error "Timeout."
}

.EXAMPLE
$t = Start-Task {param ($data) Invoke-LongRunning -Data $data} -ArgumentList $data
Note: Start-Task takes ScriptBlock and ArgumentList, just like Invoke-Command.

.EXAMPLE
Start-Task { Get-ChildItem C:\ } | Receive-Task -AutoRemoveTask
Note: Receive-Task waits for the task to complete and returns the result (and errors too).
#>
function Start-Task {
    [CmdletBinding()]
    param (
        # Command to execute.
        [Parameter(ParameterSetName='Command', Mandatory=$true, Position=0)]
        [string]$Command,
        # Parameters (name and value) to the command.
        [Parameter(ParameterSetName='Command')]
        $Parameters,
        # ScriptBlock to execute.
        [Parameter(ParameterSetName='Script', Mandatory=$true, Position=0)]
        [ScriptBlock]$ScriptBlock,
        # ArgumentList to ScriptBlock
        [Parameter(ParameterSetName='Script')]
        [object[]]$ArgumentList
    )

    if (-not $Script:runspacePool) {
        Write-Error -Message "Open-TaskRunspace must be called in advance."
        return
    }

    # Create a PowerShell instance and set paramters if any.
    [PowerShell]$ps = [PowerShell]::Create()
    $ps.RunspacePool = $Script:runspacePool

    switch -Wildcard ($PSCmdlet.ParameterSetName) {
        'Command' {
            $ps.AddCommand($Command) | Out-Null
            foreach ($key in $Parameters.Keys) {
                $ps.AddParameter($key, $Parameters[$key]) | Out-Null
            }
            break
        }

        'Script' {
            $ps.AddScript($ScriptBlock) | Out-Null
            foreach ($p in $ArgumentList) {
                $ps.AddArgument($p) | Out-Null
            }
            break
        }
    }

    # Start the command
    $ar = $ps.BeginInvoke()

    [PSCustomObject]@{
        AsyncResult = $ar
        PowerShell = $ps
        # These are diagnostic purpose
        ScriptBlock = $ScriptBlock
        ArgumentList = $ArgumentList
    }
}

<#
.SYNOPSIS
Wait for a task with optional timeout. By default, it waits indefinitely.
It returns the task object if the task completes before the timeout.
When timeout occurs, there is no output.
#>
function Wait-Task {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Task,
        # By default, it waits indefinitely
        # TimeSpan that represents -1 milliseconds is to wait indefinitely.
        [TimeSpan]$Timeout = [TimeSpan]::FromMilliseconds(-1)
    )

    process {
        foreach ($t in $Task) {
            [IAsyncResult]$ar = $t.AsyncResult
            try {
                if ($ar.AsyncWaitHandle.WaitOne($Timeout)) {
                    $t
                }
            }
            catch {
                Write-Error -ErrorRecord $_
            }
        }
    }
}

function Receive-Task {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Task,
        [switch]$AutoRemoveTask,
        [string]$TaskErrorVariable
    )

    process {
        foreach ($t in $Task) {
            [powershell]$ps = $t.PowerShell
            [IAsyncResult]$ar = $t.AsyncResult

            try {
                $ar.AsyncWaitHandle.WaitOne() | Out-Null
                $ps.EndInvoke($ar)
            }
            catch {
                Write-Error -Message "Task threw a terminating error.`nScriptBlock: $($t.ScriptBlock); ArgumentList: $($t.ArgumentList)`n$_" -Exception $_.Exception
            }

            if ($ps.HadErrors) {
                $ps.Streams.Error | ForEach-Object {
                    # Include the ErrorRecord's InvocationInfo so that it's easier to understand the origin of error.
                    Write-Error -Message "Task has a non-terminating error.`nScriptBlock: $($t.ScriptBlock); ArgumentList: $($t.ArgumentList);`n$($_.InvocationInfo.MyCommand): $($_.Exception.Message);`n$($_.InvocationInfo.PositionMessage)" -Exception $_.Exception

                    if ($TaskErrorVariable) {
                        # Scope 1 is the parent scope, but it's not necessarily the caller scope.
                        # If the caller is a function in this module, then scope 1 is the caller function.
                        # However, if it's called from outside of module, scope 1 is the module's script scope. Thus the caller does not get the error.
                        # Because this function is meant to be moudule-internal and should be called only within the moudle, Scope 1 is ok for now.
                        New-Variable -Name $TaskErrorVariable -Value $($ps.Streams.Error.ReadAll()) -Scope 1 -Force

                        # To see if it's called from within this moudle, maybe I can check the SessionState.
                        # if ($SessionState -eq $ExecutionContext.SessionState) {
                        #     New-Variable -Name $TaskErrorVariable -Value $($ps.Streams.Error.ReadAll()) -Scope 1 -Force
                        # }
                        # else {
                        #     $SessionState.PSVariable.Set($TaskErrorVariable,$ps.Streams.Error.ReadAll())
                        # }
                    }
                }
            }

            if ($AutoRemoveTask) {
                Remove-Task $t
            }
        }
    }
}

function Remove-Task {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Task
    )

    process {
        foreach ($t in $Task) {
            [powershell]$ps = $t.PowerShell
            [IAsyncResult]$ar = $t.AsyncResult

            # Note: Disposing PowerShell instance will stop the currently running command & its thread.
            # So there's no need to call EndInvoke() if you don't need the result.
            $ps.Dispose()
            $ar.AsyncWaitHandle.Close()
        }
    }
}

function Stop-Task {
    param (
        [Parameter(Mandatory = $true)]
        $Task
    )

    process {
        foreach ($t in $Task) {
            [powershell]$ps = $t.PowerShell
            $ps.Stop()
        }
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

    Write-Log "Starting a WAM trace."
    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"
    $err = $($stdout = Invoke-Command {
        $ErrorActionPreference = 'Continue'
        & logman.exe start trace $SessionName -pf $providerFile -o $traceFile -bs 128 -max 256 -mode $logFileMode -ets
    }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "logman failed. exit code:$LASTEXITCODE; stdout:`"$stdout`"; error:`"$err`""
        return
    }
}

function Stop-WamTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'WamTrace'
    )

    Write-Log "Stopping $SessionName"
    Stop-EtwSession $SessionName | Out-Null
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

    $Path = Resolve-Path $Path
    $providerFile = Join-Path $Path -ChildPath 'Office.prov'
    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $major = $officeInfo.Version.Split('.')[0] -as [int]
    Write-Log "Creating a provider listing according to the version $major"

    switch ($major) {
        14 {Set-Content $outlook2010Providers -Path $providerFile -ErrorAction Stop; break}
        15 {Set-Content $outlook2013Providers -Path $providerFile -ErrorAction Stop; break}
        16 {Set-Content $outlook2016Providers -Path $providerFile -ErrorAction Stop; break}
        default {throw "Couldn't find the version from $_"}
    }

    # In order to use EVENT_TRACE_FILE_MODE_NEWFILE, file name must contain "%d"
    if ($FileName -notlike "*%d*") {
        $FileName = [System.IO.Path]::GetFileNameWithoutExtension($FileName) + "_%d.etl"
    }

    $traceFile = Join-Path $Path -ChildPath $FileName
    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"

    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$logmanCommand)) {
        Write-Log "Starting an Outlook trace. SessionName:`"$SessionName`"; traceFile:`"$traceFile`"; logFileMode:`"$logFileMode`""

        $err = $($stdout = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & logman.exe start trace $SessionName -pf $providerFile -o $traceFile -bs 128 -max 256 -mode $logFileMode -ets
        }) 2>&1

        if ($err -or $LASTEXITCODE -ne 0) {
            Write-Error "logman failed. exit code:$LASTEXITCODE; stdout:`"$stdout`"; error:`"$err`""
            return
        }
    }
}

function Stop-OutlookTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'OutlookTrace'
    )

    Write-Log "Stopping $SessionName"
    Stop-EtwSession $SessionName | Out-Null
}


function Start-NetshTrace {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = 'nettrace-winhttp-webio.etl',
        [ValidateSet('None', 'Mini', 'Full')]
        $RerpotMode = 'None'
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    # Use "InternetClient_dbg" for Win10
    $win32os = Get-WmiObject win32_operatingsystem
    $osMajor = $win32os.Version.Split(".")[0] -as [int]
    if ($osMajor -ge 10) {
        $scenario = "InternetClient_dbg"
    }
    else {
        $scenario = "InternetClient"
    }

    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    if (-not (Get-Command $netshexe -ErrorAction SilentlyContinue)) {
        Write-Error "Cannot find $netshexe."
        return
    }

    Write-Log "Clearing dns cache"
    & ipconfig /flushdns | Out-Null

    Write-Log "Starting a netsh trace."
    $traceFile = Join-Path $Path -ChildPath $FileName
    $err = $($stdout = Invoke-Command  {
        $ErrorActionPreference = 'Continue'
        & $netshexe trace start scenario=$scenario capture=yes tracefile="`"$traceFile`"" overwrite=yes maxSize=2000  # correlation=yes
    }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "netsh failed.`nexit code: $LASTEXITCODE; stdout: $stdout; error: $err"
        return
    }

    # Even with "report=no" (by default), "HKEY_CURRENT_USER\System\CurrentControlSet\Control\NetTrace\Session\MiniReportEnabled" might be set to 1.
    # (This depends on Win10 version with a scenario. For InternetClient_dbg scenario, Win10 2004 and above does not generate mini report).
    # In order to suppress generating a minireport (i.e. C:\Windows\System32\gatherNetworkInfo.vbs), set MiniReportEnabled to 0 before netsh trace stop.
    # * You could set "report=disabled", but if you want the mini report specifically (not Full report), you need to manually configure the registry value.
    $netshRegPath = 'HKCU:\System\CurrentControlSet\Control\NetTrace\Session\'
    switch ($RerpotMode) {
        'None' { Set-ItemProperty -Path $netshRegPath -Name 'MiniReportEnabled' -Type DWord -Value 0; break }
        'Mini' { Set-ItemProperty -Path $netshRegPath -Name 'MiniReportEnabled' -Type DWord -Value 1; break}
        'Full' { Set-ItemProperty -Path $netshRegPath -Name 'ReportEnabled' -Type DWord -Value 1; break }
    }

    Write-Log "RerpotMode $RerpotMode is configured."
}

function Stop-NetshTrace {
    [CmdletBinding()]
    param (
        $SessionName = "NetTrace"
    )

    # Netsh session might not be found right after it started. So repeat with some delay (currently 1 + 2 + 3 = 6 seconds max).
    $maxRetry = 3
    $retry = 0
    $sessionFound = $false

    while ($retry -le $maxRetry -and -not $sessionFound) {
        if ($retry) {
            Write-Log "$SessionName was not found. Retrying after $retry seconds."
            Start-Sleep -Seconds $retry
        }

        $sessions = @(Get-EtwSession | Where-Object {$_.SessionName -like "*$SessionName*"})
        if ($sessions.Count -eq 1) {
            $SessionName = $sessions[0].SessionName
            $sessionFound = $true
            break
        }
        elseif ($sesionNames.Count -gt 1) {
            Write-Error "Found multiple sessions matching $SessionName"
            return
        }

        ++$retry
    }

    if (-not $sessionFound){
        Write-Error "Cannot find a netsh trace session"
        return
    }

    # Get a netsh trace report mode
    $sessionProps = Get-ItemProperty 'HKCU:\System\CurrentControlSet\Control\NetTrace\Session\'
    $reportMode = 'None'

    if ($sessionProps.ReportEnabled) {
        $reportMode = 'Full'
    }
    elseif ($sessionProps.MiniReportEnabled) {
        $reportMode = 'Mini'
    }

    Write-Log "ReportMode $reportMode is found."

    if ($reportMode -ne 'None') {
        Write-Progress -Activity "Stopping netsh trace" -Status "This might take a while. Generating a $reportMode Report" -PercentComplete -1
    }

    if ($env:PROCESSOR_ARCHITEW6432) {
        $netshexe = Join-Path $env:SystemRoot 'SysNative\netsh.exe'
    }
    else {
        $netshexe = Join-Path $env:SystemRoot 'System32\netsh.exe'
    }

    if (-not (Get-Command $netshexe -ErrorAction SilentlyContinue)) {
        Write-Error "Cannot find $netshexe."
        return
    }

    Write-Log "Stopping $SessionName with netsh trace stop"

    $err = $($stdout = Invoke-Command {
        $ErrorActionPreference = 'Continue'
        & $netshexe trace stop
    }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "Failed to stop netsh trace ($SessionName). exit code: $LASTEXITCODE; stdout: $stdout; error: $err"
    }

    Write-Progress -Activity "Stopping netsh trace" -Status "Done" -Completed
}

# Instead of logman, use Win32 QueryAllTracesW, StopTraceW.
# https://docs.microsoft.com/en-us/windows/win32/api/evntrace/nf-evntrace-queryalltracesw
$ETWType = @'
// https://docs.microsoft.com/en-us/windows/win32/etw/wnode-header
[StructLayout(LayoutKind.Sequential)]
public struct WNODE_HEADER
{
    public uint BufferSize;
    public uint ProviderId;
    public ulong HistoricalContext;
    public ulong KernelHandle;
    public Guid Guid;
    public uint ClientContext;
    public uint Flags;
}

// https://docs.microsoft.com/en-us/windows/win32/api/evntrace/ns-evntrace-event_trace_properties
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct EVENT_TRACE_PROPERTIES
{
    public WNODE_HEADER Wnode;
    public uint BufferSize;
    public uint MinimumBuffers;
    public uint MaximumBuffers;
    public uint MaximumFileSize;
    public uint LogFileMode;
    public uint FlushTimer;
    public uint EnableFlags;
    public int AgeLimit;
    public uint NumberOfBuffers;
    public uint FreeBuffers;
    public uint EventsLost;
    public uint BuffersWritten;
    public uint LogBuffersLost;
    public uint RealTimeBuffersLost;
    public IntPtr LoggerThreadId;
    public int LogFileNameOffset;
    public int LoggerNameOffset;
}

public struct EventTraceProperties
{
    public EVENT_TRACE_PROPERTIES Properties;
    public string SessionName;
    public string LogFileName;

    public EventTraceProperties(EVENT_TRACE_PROPERTIES properties, string sessionName, string logFileName)
    {
        Properties = properties;
        SessionName = sessionName;
        LogFileName = logFileName;
    }
}

[DllImport("kernel32.dll", ExactSpelling = true)]
public static extern void RtlZeroMemory(IntPtr dst, int length);

[DllImport("Advapi32.dll", ExactSpelling = true)]
public static extern int QueryAllTracesW(IntPtr[] PropertyArray, uint PropertyArrayCount, ref int LoggerCount);

[DllImport("Advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
public static extern int StopTraceW(ulong TraceHandle, string InstanceName, IntPtr Properties); // TRACEHANDLE is defined as ULONG64

const int MAX_SESSIONS = 64;
const int MAX_NAME_COUNT = 1024; // max char count for LogFileName & SessionName
const uint ERROR_SUCCESS = 0;

// https://docs.microsoft.com/en-us/windows/win32/etw/wnode-header
// > The size of memory must include the room for the EVENT_TRACE_PROPERTIES structure plus the session name string and log file name string that follow the structure in memory.
static readonly int PropertiesSize = Marshal.SizeOf(typeof(EVENT_TRACE_PROPERTIES)) + 2 * sizeof(char) * MAX_NAME_COUNT; // EVENT_TRACE_PROPERTIES + LogFileName & LoggerName
static readonly int LoggerNameOffset = Marshal.SizeOf(typeof(EVENT_TRACE_PROPERTIES));
static readonly int LogFileNameOffset = LoggerNameOffset + sizeof(char) * MAX_NAME_COUNT;

public static List<EventTraceProperties> QueryAllTraces()
{
    IntPtr pBuffer = IntPtr.Zero;
    List<EventTraceProperties> eventProperties = null;
    try
    {
        // Allocate native memorty to hold the entire data.
        int BufferSize = PropertiesSize * MAX_SESSIONS;
        pBuffer = Marshal.AllocCoTaskMem(BufferSize);
        RtlZeroMemory(pBuffer, BufferSize);

        IntPtr[] sessions = new IntPtr[64];

        for (int i = 0; i < 64; ++i)
        {
            //sessions[i] = pBuffer + (i * PropertiesSize); // This does not compile in .NET 2.0
            sessions[i] = new IntPtr(pBuffer.ToInt64() + (i * PropertiesSize));

            // Marshal from managed to native
            EVENT_TRACE_PROPERTIES props = new EVENT_TRACE_PROPERTIES();
            props.Wnode.BufferSize = (uint)PropertiesSize;
            props.LoggerNameOffset = LoggerNameOffset;
            props.LogFileNameOffset = LogFileNameOffset;
            Marshal.StructureToPtr(props, sessions[i], false);
        }

        int loggerCount = 0;
        int status = QueryAllTracesW(sessions, MAX_SESSIONS, ref loggerCount);

        if (status != ERROR_SUCCESS)
        {
            throw new Win32Exception(status);
        }

        eventProperties = new List<EventTraceProperties>();
        for (int i = 0; i < loggerCount; ++i)
        {
            // Marshal back from native to managed.
            EVENT_TRACE_PROPERTIES props = (EVENT_TRACE_PROPERTIES)Marshal.PtrToStructure(sessions[i], typeof(EVENT_TRACE_PROPERTIES));
            string sessionName = Marshal.PtrToStringUni(new IntPtr(sessions[i].ToInt64() + LoggerNameOffset));
            string logFileName = Marshal.PtrToStringUni(new IntPtr(sessions[i].ToInt64() + LogFileNameOffset));

            //eventProperties.Add(new EventTraceProperties { Properties = props, SessionName = sessionName, LogFileName = logFileName });
            eventProperties.Add(new EventTraceProperties(props,sessionName, logFileName));
        }
    }
    finally
    {
        if (pBuffer != IntPtr.Zero)
        {
            Marshal.FreeCoTaskMem(pBuffer);
            pBuffer = IntPtr.Zero;
        }
    }

    return eventProperties;
}

public static EventTraceProperties StopTrace(string SessionName)
{
    IntPtr pProps = IntPtr.Zero;
    try
    {
        pProps = Marshal.AllocCoTaskMem(PropertiesSize);
        RtlZeroMemory(pProps, PropertiesSize);

        EVENT_TRACE_PROPERTIES props = new EVENT_TRACE_PROPERTIES();
        props.Wnode.BufferSize = (uint)PropertiesSize;
        props.LoggerNameOffset = LoggerNameOffset;
        props.LogFileNameOffset = LogFileNameOffset;
        Marshal.StructureToPtr(props, pProps, false);

        int status = StopTraceW(0, SessionName, pProps);
        if (status != ERROR_SUCCESS)
        {
            throw new Win32Exception(status);
        }

        props = (EVENT_TRACE_PROPERTIES)Marshal.PtrToStructure(pProps, typeof(EVENT_TRACE_PROPERTIES));
        string sessionName = Marshal.PtrToStringUni(new IntPtr(pProps.ToInt64() + LoggerNameOffset));
        string logFileName = Marshal.PtrToStringUni(new IntPtr(pProps.ToInt64() + LogFileNameOffset));

        //return new EventTraceProperties { Properties = props, SessionName = sessionName, LogFileName = logFileName };
        return new EventTraceProperties(props, sessionName, logFileName);
    }
    finally
    {
        if (pProps != IntPtr.Zero)
        {
            Marshal.FreeCoTaskMem(pProps);
        }
    }
}
'@

function Get-EtwSession {
    [CmdletBinding()]
    param()

    if (-not ('Win32.ETW' -as [type])) {
        Add-type -MemberDefinition $ETWType -Namespace Win32 -Name ETW -UsingNamespace System.Collections.Generic, System.ComponentModel
    }

    try {
        [Win32.ETW]::QueryAllTraces()
    }
    catch {
        Write-Error -Message "QueryAllTraces failed. $_" -Exception $_.Exception
    }
}

function Stop-EtwSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SessionName
    )

    if (-not ('Win32.ETW' -as [type])) {
        Add-type -MemberDefinition $ETWType -Namespace Win32 -Name ETW -UsingNamespace System.Collections.Generic, System.ComponentModel
    }

    try {
        return [Win32.ETW]::StopTrace($SessionName)
    }
    catch {
        Write-Error -Message "StopTrace for $SessionName failed. $_" -Exception $_.Exception
    }
}

function Start-PSR {
    param(
        [parameter(Mandatory = $true)]
        $Path,
        $FileName = "PSR.zip",
        [switch]$ShowGUI
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    # File name must be ***.zip
    if ([IO.Path]::GetExtension($FileName) -ne ".zip"){
        $FileName = [IO.Path]::GetFileNameWithoutExtension($FileName) + '.zip'
    }

    # For Win7, maxsc is 100
    $maxScreenshotCount = 100

    $win32os = Get-WmiObject win32_operatingsystem
    $osMajor = $win32os.Version.Split(".")[0] -as [int]
    $osMinor = $win32os.Version.Split(".")[1] -as [int]

    if ($osMajor -gt 6 -or ($osMajor -eq 6 -and $osMinor -ge 3)) {
        $maxScreenshotCount = 300
    }

    if (-not (Get-Command 'psr.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "psr.exe is not available."
        return
    }

    Write-Log "Starting PSR $(if ($ShowGUI) {'with UI'} else {'without UI'}). maxScreenshotCount: $maxScreenshotCount"
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
        Write-Error "PSR failed to start"
        return
    }
}

function Stop-PSR {
    [CmdletBinding()]
    param ()

    $process = Get-Process -Name psr -ErrorAction SilentlyContinue
    if (-not $process){
        Write-Error 'There is no psr.exe process'
        return
    }

    Write-Log 'Stopping PSR'
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
                    Write-Error -Message "Failed to add $($file.FullName). $_" -Exception $_.Exception
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
            Write-Progress -Activity "Cleaning up" -Status "Please wait" -PercentComplete -1
            Get-ChildItem $Path -Exclude $ZipFileName | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Progress -Activity "Cleaning up" -Status "Please wait" -Completed
            $filesRemoved = $true
        }

        [PSCustomObject]@{
            ZipFilePath = $zipFilePath.ToString()
            FilesRemoved = $filesRemoved -eq $true
        }
    }
    else {
        throw "Zip file wasn't successfully created at $zipFilePath"
    }
}

function Save-EventLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Path
    )

    if (-not (Test-Path $Path -ErrorAction Stop)) {
        New-Item -ItemType directory $Path | Out-Null
    }
    $Path = Resolve-Path $Path

    $logs = @(
        'Application'
        'System'
        (wevtutil el) -match "Microsoft-Windows-Windows Firewall With Advanced Security|AAD"
    )

    $tasks = @(
        foreach ($log in $logs) {
            $fileName = $log.Replace('/', '_') + '.evtx'
            $filePath = Join-Path $Path -ChildPath $fileName
            Write-Log "Saving $log to $filePath"
            Start-Task -ScriptBlock {
                param ($log, $filePath)
                wevtutil epl $log $filePath /ow
                wevtutil al $filePath
            } -ArgumentList $log, $filePath
        }
    )

    $tasks | Receive-Task -AutoRemoveTask
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

    $hklm = $productsKey = $null
    try {
        # Use .NET registry API (with [RegistryView]::Registry64) instead of PowerShell here to avoid registry redirection occurs on 32bit PowerShell on 64bit OS for HKLM\SOFTWARE.
        if ('Microsoft.Win32.RegistryView' -as [type]) {
            $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64);
        }
        elseif (-not $env:PROCESSOR_ARCHITEW6432) {
            # RegistryView is not available, but it's OK because no WOW64.
            $hklm = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [string]::Empty);
        }
        else {
            # This is the case where registry rediction takes place (32bit PowerShell on 64bit OS). Bail.
            Write-Error "32bit PowerShell is running on 64bit OS and .NET 4.0 is not used. Please run 64bit PowerShell."
            return
        }

        $productsKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products')

        foreach ($productName in $productsKey.GetSubKeyNames()) {
            if ($null -eq $productName -or ($OfficeOnly -and $productName -notmatch 'F01FEC')) {
                continue
            }

            $productKey = $productsKey.OpenSubKey($productName)

            foreach ($subkeyName in $productKey.GetSubKeyNames()) {
                if ($subkeyName -ne 'Patches') {
                    continue
                }

                $patchesKey = $productKey.OpenSubKey($subkeyName)
                foreach ($patchName in $patchesKey.GetSubKeyNames()) {
                    $patchKey = $patchesKey.OpenSubKey($patchName)

                    $state = $patchKey.GetValue('State')

                    if ($AppliedOnly -and $PatchState[$state] -ne 'MSIPATCHSTATE_APPLIED') {
                        continue
                    }

                    $displayName = $patchKey.GetValue('DisplayName')
                    $moreInfoURL = $patchKey.GetValue('MoreInfoURL')
                    $installed = $patchKey.GetValue('Installed')

                    if (-not $displayName -and -not $moreInfoURL) {
                        continue
                    }

                    # extract KB number
                    $KB = $null
                    if ($moreInfoURL -match 'https?://support.microsoft.com/kb/(?<KB>\d+)') {
                        $KB = $Matches['KB']
                    }

                    [PSCustomObject]@{
                        DisplayName = $displayName
                        KB = $KB
                        MoreInfoURL = $moreInfoURL
                        Installed = $installed
                        PatchState = $PatchState[$state]
                    }

                    $patchKey.Close()
                }

                $patchesKey.Close()
            }

            $productKey.Close()
        }
    }
    finally {
        if ($productsKey) {
            $productsKey.Close()
        }

        if ($hklm) {
            $hklm.Close()
        }
    }
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

function Get-InstalledUpdate
{
    [CmdletBinding()]
    param()

    # Ask items in AppUpdatesFolder from Shell
    # FOLDERID_AppUpdates == a305ce99-f527-492b-8b1a-7e76fa98d6e4
    $shell = $appUpdates = $null

    try {
        $shell = New-Object -ComObject Shell.Application
        $appUpdates = $shell.NameSpace('Shell:AppUpdatesFolder')
        if ($null -eq $appUpdates) {
            Write-Log "Cannot obtain Shell:AppUpdatesFolder. Probabliy 32bit PowerShell is used on 64bit OS"
            Write-Error "Cannot obtain Shell:AppUpdatesFolder"
            return
        }

        $items = $appUpdates.Items()

        foreach ($item in $items) {
            # https://docs.microsoft.com/en-us/windows/win32/shell/folder-getdetailsof
            [PSCustomObject]@{
                Name        = $item.Name
                Program     = $appUpdates.GetDetailsOf($item, 2)
                Version     = $appUpdates.GetDetailsOf($item, 3)
                Publisher   = $appUpdates.GetDetailsOf($item, 4)
                URL         = $appUpdates.GetDetailsOf($item, 7)
                InstalledOn = $appUpdates.GetDetailsOf($item, 12)
            }
            [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($item) | Out-Null
        }
    }
    finally {
        if ($appUpdates) {
            [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($appUpdates) | Out-Null
        }
        if ($shell) {
            [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($shell) | Out-Null
        }
    }
}


function Get-LogonUser {
    [CmdletBinding()]
    param(
        [switch]$IgnoreCache
    )

    # If there's a cache, use it unless IgnoreCache is specified
    if (-not $IgnoreCache -and $Script:LogonUser) {
        Write-Log "Returning a cache."
        return $Script:LogonUser
    }

    # quser.exe might not be available.
    if (-not (Get-Command quser.exe -ErrorAction SilentlyContinue)) {
        Write-Error "quser.exe is not available."
        return
    }

    $err = $($quserResult = Invoke-Command {
        $ErrorActionPreference = 'Continue'
        & quser.exe
    }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "quser failed. exit code:$LASTEXITCODE; stdout:`"$quserResult`"; error:`"$err`""
        return
    }

    $currentSession = $quserResult | Where-Object {$_.StartsWith('>')} | Select-Object -First 1
    if (-not $currentSession) {
        Write-Error "Cannot find current session with quser."
        return
    }

    Write-Log "Current session: $currentSession"
    $match = [Regex]::Match($currentSession, '^>(?<name>.+?)\s{2,}')
    $userName = $match.Groups['name'].Value

    # WMI Win32_UserAccount can be very slow. I'm avoiding here.
    # Get-WmiObject -Class Win32_UserAccount -Filter "Name = '$userName'"

    $sid = ConvertTo-UserSID $userName

    $Script:LogonUser = [PSCustomObject]@{
        Name = $userName
        SID = $sid
    }

    $Script:LogonUser
}

function ConvertTo-UserSID {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string]$UserName
    )

    try {
        $account = New-Object System.Security.Principal.NTAccount($UserName)
        $sid = $account.Translate([System.Security.Principal.SecurityIdentifier]).Value
        return $sid
    }
    catch {
        Write-Error -Message "Cannot obtain user SID for $UserName." -Exception $_.Exception
    }
}

<#
.SYNOPSIS
Test if a given string is a valid SID
#>
function Test-SID {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        [string]$SID
    )

    try {
        New-Object System.Security.Principal.SecurityIdentifier($SID) | Out-Null
        return $true
    }
    catch {
        # ignore error here
    }

    $false
}

<#
.SYNOPSIS
Get a given local user's registry root. If User is empty, it just returns HKCU.
#>
function Get-UserRegistryRoot {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User,
        # Skip "Registry::" prefix
        [switch]$SkipRegistryPrefix
    )

    if ($User) {
        # If user SID is given use it as it is; otherwise convert SID. when failed to convert, just return. An ErrorRecord will be written to error stream by ConvertTo-UserSID
        if (Test-SID $User) {
            $userSID = $User
        }
        else {
            $userSID = ConvertTo-UserSID $User
            if (-not $userSID) {
                return
            }
        }

        $userRegRoot = "HKEY_USERS\$userSID"

        if (-not ($userRegRoot -and (Test-Path "Registry::$userRegRoot"))) {
            Write-Error "Cannot find $userRegRoot."
            return
        }
    }
    else {
        Write-Log "User is empty. Use HKCU."
        $userRegRoot = 'HKCU'
    }

    if (-not $SkipRegistryPrefix) {
        $userRegRoot = "Registry::$userRegRoot"
    }

    $userRegRoot
}

<#
.SYNOPSIS
Get a given user's profile path (i.e. same as USERPROFILE environment variable)
#>
function Get-UserProfilePath {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User
    )

    if (-not $User) {
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    # If user SID is given use it as it is; otherwise convert SID. when failed to convert, just return. An ErrorRecord will be written to error stream by ConvertTo-UserSID
    if (Test-SID $User) {
        $userSID = $User
    }
    else {
        $userSID = ConvertTo-UserSID $User
        if (-not $userSID) {
            return
        }
    }

    # Get the value of ProfileImagePath
    $userProfile = Get-ItemProperty "Registry::HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$userSID\"
    $userProfile.ProfileImagePath
}

<#
.SYNOPSIS
Get a given user's shell folder path (e.g. "LocalAppData", "Desktop" etc.)
#>
function Get-UserShellFolder {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User,
        [parameter(Mandatory = $true)]
        # Shell folder name (e.g. "AppData", "Desktop", "Local AppData" etc.)
        [string]$ShellFolderName
    )

    $userRegRoot = Get-UserRegistryRoot -User $User
    if (-not $userRegRoot) {
        return
    }

    $shellFolders = Get-ItemProperty $(Join-Path $userRegRoot "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
    $folderPath = $shellFolders.$ShellFolderName

    if (-not $folderPath) {
        return
    }

    # Folder path is like "%USERPROFILE%\AppData\Local". Replace USERPROFILE.
    $userProfile = Get-UserProfilePath $User
    $folderPath.Replace('%USERPROFILE%', $userProfile)
}

function Save-OfficeRegistry {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        [string]$User
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $registryKeys = @(
        "HKCU\Software\Microsoft\Exchange"
        "HKCU\Software\Policies\Microsoft\Exchange"
        "HKCU\Software\Microsoft\Office"
        "HKCU\Software\Policies\Microsoft\Office"
        "HKCU\Software\Wow6432Node\Microsoft\Office"
        "HKCU\Software\Wow6432Node\Policies\Microsoft\Office"
        "HKLM\Software\Microsoft\Office"
        "HKLM\Software\PoliciesMicrosoft\Office"
        "HKLM\Software\WOW6432Node\Microsoft\Office"
        "HKLM\Software\WOW6432Node\Policies\Microsoft\Office")

    $userRegRoot = Get-UserRegistryRoot $User -SkipRegistryPrefix
    if ($userRegRoot) {
        $registryKeys = $registryKeys | ForEach-Object {$_.Replace("HKCU", $userRegRoot)}
    }

    # Make sure NOT to use WOW64 version of reg.exe when running on 32bit PowerShell on 64bit OS.
    # I could use "/reg:64" option of reg.exe, but it's not available for Win7.
    if ($env:PROCESSOR_ARCHITEW6432) {
        $regexe = Join-Path $env:SystemRoot 'SysNative\reg.exe'
    }
    else {
        $regexe = Join-Path $env:SystemRoot 'System32\reg.exe'
    }

    # If, for some reason, reg.exe is not available, bail.
    if (-not (Get-Command $regexe -ErrorAction SilentlyContinue)) {
        Write-Error "$regexe is not avaialble."
        return
    }

    foreach ($key in $registryKeys) {
        $err = $($queryResult = Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & $regexe Query $key
        }) 2>&1

        if ($null -eq $queryResult) {
            Write-Log "$key does not exist"
            continue;
        }

        # Cannot use Test-Path because when running 32bit PS on 64bit OS, HKLM\Software is redirected to WOW6432Node
        # if (-not (Test-Path $key)) {
        #     Write-Log "$key does not exist"
        #     continue
        # }

        $filePath = Join-Path $Path -ChildPath "$($key.Replace('\','_')).reg"

        if (Test-Path $filePath) {
            Remove-Item $filePath -Force
        }

        Write-Log "Saving $key to $filePath"
        $err = $(Invoke-Command {
            $ErrorActionPreference = 'Continue'
            & $regexe export $key $filePath | Out-Null
        }) 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Error "$key is not exported. exit code = $LASTEXITCODE. $err"
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

    # Key = command to run, Value = file name used for saving
    $commands = [ordered]@{
        {Get-WmiObject -Class Win32_ComputerSystem} = 'Win32_ComputerSystem.xml'
        {Get-WmiObject -Class Win32_OperatingSystem} = 'Win32_OperatingSystem.xml'
        {Get-ProxySetting} = $null
        {Get-NLMConnectivity} = $null
        {Get-WSCAntivirus} = $null
        {Get-InstalledUpdate} = $null
        {Get-JoinInformation} = $null
        {Get-DeviceJoinStatus} = 'DeviceJoinStatus.txt'
    }

    $commands.GetEnumerator() | ForEach-Object {
        $command = $_.Key
        $fileName = $_.Value
        Run-Command $command -Path $Path -FileName $fileName
    }
}

function Save-OSConfigurationMT {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $tasks = @(
    Start-Task {param($path) Get-WmiObject -Class Win32_ComputerSystem | Export-Clixml -Path $path} -ArgumentList (Join-Path $Path "Win32_ComputerSystem.xml")
    Start-Task {param($path) Get-WmiObject -Class Win32_OperatingSystem | Export-Clixml -Path $path} -ArgumentList (Join-Path $Path "Win32_OperatingSystem.xml")
    Start-Task {param($path) Get-ProxySetting | Export-Clixml -Path $path} -ArgumentList (Join-Path $Path "ProxySetting.xml")
    Start-Task {param($path) Get-NLMConnectivity | Export-Clixml -Path $path} -ArgumentList (Join-Path $Path "NLMConnectivity.xml")
    Start-Task {param($path) Get-WSCAntivirus -ErrorAction SilentlyContinue | Export-Clixml -Path $path} -ArgumentList (Join-Path $Path "WSCAntivirus.xml")
    Start-Task {param($path) Get-InstalledUpdate -ErrorAction SilentlyContinue | Export-Clixml -Path $path} -ArgumentList (Join-Path $Path "InstalledUpdate.xml")
    Start-Task {param($path) Get-JoinInformation -ErrorAction SilentlyContinue | Export-Clixml -Path $path} -ArgumentList (Join-Path $Path "JoinInformation.xml")
    Start-Task {param($path) Get-DeviceJoinStatus -ErrorAction SilentlyContinue | Out-File -FilePath $path} -ArgumentList (Join-Path $Path "DeviceJoinStatus.txt")
    )

    Write-Verbose "waiting for tasks..."
    $tasks | Receive-Task -AutoRemoveTask
}

function Save-NetworkInfo {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    # These are from C:\Windows\System32\gatherNetworkInfo.vbs with some extra.
    # Key = command to run, Value = file name used for saving. When file name is $null, Run-Command decides the file name.
    $commands = [ordered]@{
        {Get-NetAdapter -IncludeHidden} = $null
        {Get-NetAdapterAdvancedProperty} = $null
        {Get-NetAdapterBinding -IncludeHidden} = $null
        {Get-NetIpConfiguration -Detailed} = $null
        {Get-DnsClientNrptPolicy} = $null
        # 'Resolve-DnsName bing.com'
        # 'ping bing.com -4'
        # 'ping bing.com -6'
        # 'Test-NetConnection bing.com -InformationLevel Detailed'
        # 'Test-NetConnection bing.com -InformationLevel Detailed -CommonTCPPort HTTP'
        {Get-NetRoute} = $null
        {Get-NetIPaddress} = $null
        {Get-NetLbfoTeam} = $null

        # {Get-Service -Name:VMMS} = $null
        # {Get-VMSwitch} = $null
        # {Get-VMNetworkAdapter -all} = $null
        # {Get-WindowsOptionalFeature -Online} = $null
        # {Get-Service} = $null
        # {Get-PnpDevice | Get-PnpDeviceProperty -KeyName DEVPKEY_Device_InstanceId,DEVPKEY_Device_DevNodeStatus,DEVPKEY_Device_ProblemCode} = $null

        {Get-NetIPInterface} = $null
        {Get-NetConnectionProfile} = $null
        {ipconfig /all} = $null

        # Dump Windows Firewall config
        {netsh advfirewall monitor show currentprofile} = $null # current profiles
        {netsh advfirewall monitor show firewall} = $null # firewall configuration
        {netsh advfirewall monitor show consec} = $null # connection security configuration
        {netsh advfirewall firewall show rule name=all verbose} = $null # firewall rules
        {netsh advfirewall consec show rule name=all verbose} = $null # connection security rules
        {netsh advfirewall monitor show firewall rule name=all} = $null # firewall rules from Dynamic Store
        {netsh advfirewall monitor show consec rule name=all} = $null # connection security rules from Dynamic Store
    }

    $commands.GetEnumerator() | ForEach-Object {
        $command = $_.Key
        $fileName = $_.Value
        Run-Command $command -Path $Path -FileName $fileName
    }
}

<#
.DESCRIPTION
Run a given script block. If Path is given, save the result there.
If FileName is given, it's used for the file name for saving the result. If its extension is not ".xml", Out-File will be used. Otherwise Export-CliXml will be used.
If FileName is not give, the file name will be auto-decided. If the command is an application, then Out-File will be used. Otherwise Export-CliXml will be used.
#>
function Run-Command {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position=0)]
        [ScriptBlock]$ScriptBlock,
        [object[]]$ArgumentList,
        # Folder to save to
        $Path,
        # File name used for saving
        [string]$FileName
    )

    $result = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # To redirect error, call operator (&) is used, instead of $ScriptBlock.InvokeReturnAsIs().
        $($result = & $ScriptBlock @ArgumentList) 2>&1 | Write-Log
    }
    catch {
        Write-Log "'$ScriptBlock' threw a terminating error. $_"
    }

    $sw.Stop()
    Write-Log "'$ScriptBlock' took $($sw.ElapsedMilliseconds) ms."

    if ($null -eq $result) {
        Write-Log "It returned nothing."
        return
    }

    # If Path is given, save the result.
    if ($Path) {
        $sb = $ScriptBlock.ToString()
        $exportAsXml = $true

        if ($FileName) {
            if ([IO.Path]::GetExtension($FileName) -ne '.xml') {
                $exportAsXml = $false
            }
        }
        else {
            # Decide the filename & export method based on the command type
            $Command = ([RegEx]::Match($sb, '\w+-\w+')).Value
            if (-not $Command) {
                $Command = $sb
            }

            if ($Command.IndexOf(' ') -ge 0) {
                $commandName = $Command.SubString(0, $Command.IndexOf(' '))
            }
            else {
                $commandName = $Command
            }

            $cmd = Get-Command $commandName -ErrorAction SilentlyContinue
            if ($cmd.CommandType -eq 'Application') {
                # To be more strict, I could use [System.IO.Path]::GetInvalidFileNameChars(). But it's ok for now.
                $FileName = $Command.Replace('/', '-') + ".txt"
                $exportAsXml = $false
            }
            else {
                $FileName = $commandName.SubString($commandName.IndexOf('-') + 1) + ".xml"
            }
        }

        if ($exportAsXml) {
            $result | Export-Clixml -Path (Join-Path $Path $FileName)
        }
        else {
            $result | Out-File -FilePath (Join-Path $Path $FileName)
        }
    }
    else {
        $result
    }
}

function Save-NetworkInfoMT {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    # These are from C:\Windows\System32\gatherNetworkInfo.vbs with some extra.
    $PSDefaultParameterValues.Add('Start-Task:ArgumentList', $Path)

    $tasks = @(
    Start-Task {param ($Path) Get-NetAdapter -IncludeHidden | Export-Clixml (Join-Path $Path 'NetAdapter.xml')}
    Start-Task {param ($Path) Get-NetAdapterAdvancedProperty | Export-Clixml (Join-Path $Path 'NetAdapterAdvancedProperty.xml')}
    Start-Task {param ($Path) Get-NetAdapterBinding -IncludeHidden | Export-Clixml (Join-Path $Path 'NetAdapterBinding.xml')}
    Start-Task {param ($Path) Get-NetIpConfiguration -Detailed | Export-Clixml (Join-Path $Path 'NetIpConfiguration.xml')}
    Start-Task {param ($Path) Get-DnsClientNrptPolicy | Export-Clixml (Join-Path $Path 'DnsClientNrptPolicy.xml')}
    Start-Task {param ($Path) Get-NetRoute | Export-Clixml (Join-Path $Path 'NetRoute.xml')}
    Start-Task {param ($Path) Get-NetIPaddress | Export-Clixml (Join-Path $Path 'NetIPaddress.xml')}
    Start-Task {param ($Path) Get-NetLbfoTeam | Export-Clixml (Join-Path $Path 'NetLbfoTeam.xml')}
    Start-Task {param ($Path) Get-NetIPInterface | Export-Clixml (Join-Path $Path 'NetIPInterface.xml')}
    Start-Task {param ($Path) Get-NetConnectionProfile | Export-Clixml (Join-Path $Path 'NetConnectionProfile.xml')}
    Start-Task {param ($Path) ipconfig /all | Out-File (Join-Path $Path 'ipconfig_all.txt')}
    Start-Task {param ($Path) netsh advfirewall monitor show currentprofile | Out-File (Join-Path $Path 'netsh advfirewall monitor show currentprofile.txt')}
    Start-Task {param ($Path) netsh advfirewall monitor show firewall | Out-File (Join-Path $Path 'netsh advfirewall monitor show firewall.txt')}
    Start-Task {param ($Path) netsh advfirewall firewall show rule name=all verbose | Out-File (Join-Path $Path 'netsh advfirewall firewall show rule name=all verbose.txt')}
    Start-Task {param ($Path) netsh advfirewall consec show rule name=all verbose | Out-File (Join-Path $Path 'netsh advfirewall consec show rule name=all verbose.txt')}
    Start-Task {param ($Path) netsh advfirewall monitor show firewall rule name=all | Out-File (Join-Path $Path 'netsh advfirewall monitor show firewall rule name=all.txt')}
    Start-Task {param ($Path) netsh advfirewall monitor show consec rule name=all | Out-File (Join-Path $Path 'netsh advfirewall monitor show consec rule name=all.txt')}
    )

    $PSDefaultParameterValues.Remove('Start-Task:ArgumentList')

    Write-Log "Waiting for tasks to complete."
    $tasks | Receive-Task -AutoRemoveTask
    Write-Log "All tasks are complete."
}

function Get-ProxySetting {
    [CmdletBinding()]
    param(
    )

    Write-Log "Running as $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

    # props hold the return object properties.
    $props= @{}

    # Get WebProxy class to get IE config
    # N.B. GetDefaultProxy won't be really needed, but I'm keeping it for now.
    # It's possible that [System.Net.WebProxy]::GetDefaultProxy() throws
    try {
        $props['WebProxyDefault'] = [System.Net.WebProxy]::GetDefaultProxy()
    }
    catch {
        Write-Log "$_"
    }

    # Get WinHttp & current user's IE proxy settings.
    # Use Win32 WinHttpGetDefaultProxyConfiguration & WinHttpGetIEProxyConfigForCurrentUser.
    # I'm not using "netsh winhttp show proxy", because the output is system language dependent.  Netsh just calls this function anyway.
    $WinHttpDef = @'
// https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_proxy_info
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct WINHTTP_PROXY_INFO
{
    public uint dwAccessType;
    public string lpszProxy;
    public string lpszProxyBypass;
}

// https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_current_user_ie_proxy_config
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
{
    public bool fAutoDetect;
    public string lpszAutoConfigUrl;
    public string lpszProxy;
    public string lpszProxyBypass;
}

// From winhttp.h
// WinHttpOpen dwAccessType values (also for WINHTTP_PROXY_INFO::dwAccessType)
public enum ProxyAccessType
{
    WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0,
    WINHTTP_ACCESS_TYPE_NO_PROXY = 1,
    WINHTTP_ACCESS_TYPE_NAMED_PROXY = 3,
    WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY = 4
}

// https://docs.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpgetdefaultproxyconfiguration
[DllImport("winhttp.dll", SetLastError = true)]
public static extern bool WinHttpGetDefaultProxyConfiguration(out WINHTTP_PROXY_INFO proxyInfo);

// https://docs.microsoft.com/en-us/windows/win32/api/winhttp/nf-winhttp-winhttpgetieproxyconfigforcurrentuser
[DllImport("winhttp.dll", SetLastError = true)]
public static extern bool WinHttpGetIEProxyConfigForCurrentUser(out WINHTTP_CURRENT_USER_IE_PROXY_CONFIG proxyConfig);
'@

    if (-not ('Win32.WinHttp' -as [type])) {
        Add-Type -MemberDefinition $WinHttpDef -Name WinHttp -Namespace Win32
    }

    $proxyInfo = New-Object Win32.WinHttp+WINHTTP_PROXY_INFO
    if ([Win32.WinHttp]::WinHttpGetDefaultProxyConfiguration([ref] $proxyInfo)) {
        $props['WinHttpDirectAccess'] = $proxyInfo.dwAccessType -eq [Win32.WinHttp+ProxyAccessType]::WINHTTP_ACCESS_TYPE_NO_PROXY
        $props['WinHttpProxyServer'] = $proxyInfo.lpszProxy
        $props['WinHttpBypassList'] = $proxyInfo.lpszProxyBypass
        $props['WINHTTP_PROXY_INFO'] = $proxyInfo # for debugging purpuse
    }
    else {
        Write-Error ("Win32 WinHttpGetDefaultProxyConfiguration failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }

    $userIEProxyConfig = New-Object Win32.WinHttp+WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
    if ([Win32.WinHttp]::WinHttpGetIEProxyConfigForCurrentUser([ref] $userIEProxyConfig)) {
        $props['UserIEAutoDetect'] = $userIEProxyConfig.fAutoDetect
        $props['UserIEAutoConfigUrl'] = $userIEProxyConfig.lpszAutoConfigUrl
        $props['UserIEProxy'] = $userIEProxyConfig.lpszProxy
        $props['UserIEProxyBypass'] = $userIEProxyConfig.lpszProxyBypass
        $props['WINHTTP_CURRENT_USER_IE_PROXY_CONFIG'] = $userIEProxyConfig # for debugging purpuse
        $props['User'] = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    else {
        Write-Error ("Win32 WinHttpGetIEProxyConfigForCurrentUser failed with 0x{0:x8}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }

    Write-Log "UserIE*** properties correspond to WINHTTP_CURRENT_USER_IE_PROXY_CONFIG obtained by WinHttpGetIEProxyConfigForCurrentUser. See https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_proxy_info"
    Write-Log "WinHttp*** properties correspond to WINHTTP_PROXY_INFO obtained by WinHttpGetDefaultProxyConfiguration. See https://docs.microsoft.com/en-us/windows/win32/api/winhttp/ns-winhttp-winhttp_current_user_ie_proxy_config"

    [PSCustomObject]$props
}

function Get-NLMConnectivity {
    [CmdletBinding()]
    param()

    $CLSID_NetworkListManager = [Guid]'DCB00C01-570F-4A9B-8D69-199FDBA5723B'
    $type = [Type]::GetTypeFromCLSID($CLSID_NetworkListManager)
    $nlm = [Activator]::CreateInstance($type)

    $isConnectedToInternet = $nlm.IsConnectedToInternet
    $conn = $nlm.GetConnectivity()
    Write-Log ("INetworkListManager::GetConnectivity 0x{0:x8}" -f $conn)

    $refCount = [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($nlm);
    Write-Log "NetworkListManager COM object's remaining ref count: $refCount"
    $nlm = $null

    # NLM_CONNECTIVITY enumeration
    # https://docs.microsoft.com/en-us/windows/win32/api/netlistmgr/ne-netlistmgr-nlm_connectivity

    # From netlistmgr.h
    $NLM_CONNECTIVITY = @{
        NLM_CONNECTIVITY_DISCONNECTED      = 0
        NLM_CONNECTIVITY_IPV4_NOTRAFFIC    = 1
        NLM_CONNECTIVITY_IPV6_NOTRAFFIC    = 2
        NLM_CONNECTIVITY_IPV4_SUBNET	   = 0x10
        NLM_CONNECTIVITY_IPV4_LOCALNETWORK = 0x20
        NLM_CONNECTIVITY_IPV4_INTERNET	   = 0x40
        NLM_CONNECTIVITY_IPV6_SUBNET	   = 0x100
        NLM_CONNECTIVITY_IPV6_LOCALNETWORK = 0x200
        NLM_CONNECTIVITY_IPV6_INTERNET	   = 0x400
    }

    $connectivity = New-Object System.Collections.Generic.List[string]

    foreach ($entry in $NLM_CONNECTIVITY.GetEnumerator()) {
        if ($conn -band $entry.Value) {
            $connectivity.Add($entry.Key)
        }
    }

    [PSCustomObject]@{
        IsConnectedToInternet = $isConnectedToInternet
        Connectivity = $connectivity
    }
}

function Get-WSCAntivirus {
    [CmdletBinding()]
    param()

    $WscDef = @'
    public enum WSC_SECURITY_PROVIDER_HEALTH
    {
        WSC_SECURITY_PROVIDER_HEALTH_GOOD,
        WSC_SECURITY_PROVIDER_HEALTH_NOTMONITORED,
        WSC_SECURITY_PROVIDER_HEALTH_POOR,
        WSC_SECURITY_PROVIDER_HEALTH_SNOOZE
    }

    // https://docs.microsoft.com/en-us/windows/win32/api/wscapi/nf-wscapi-wscgetsecurityproviderhealth
    [DllImport("Wscapi.dll", SetLastError = true)]
    public static extern int WscGetSecurityProviderHealth(uint Providers, out int pHealth);
'@

    if (-not ('Win32.WSC' -as [type])) {
        Add-Type -MemberDefinition $WscDef -Name WSC -Namespace Win32
    }

    # from Wscapi.h
    $WSC_SECURITY_PROVIDER_ANTIVIRUS = 4 -as [Uint32]
    [Win32.WSC+WSC_SECURITY_PROVIDER_HEALTH]$health = [Win32.WSC+WSC_SECURITY_PROVIDER_HEALTH]::WSC_SECURITY_PROVIDER_HEALTH_POOR

    # This call could fail with a terminating error on the server OS since Wscapi.dll is not available.
    # Catch it and convert it a non-terminating error so that the caller can ignore with ErrorAction.
    try {
        $hr = [Win32.WSC]::WscGetSecurityProviderHealth($WSC_SECURITY_PROVIDER_ANTIVIRUS, [ref]$health)
        [PSCustomObject]@{
            HRESULT = $hr
            Health  = $health
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
}

function Get-JoinInformation {
    [CmdletBinding()]
    param()

    $def = @'
[DllImport("Netapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
public static extern uint NetGetJoinInformation(string server, out IntPtr name, out uint status);

[DllImport("Netapi32.dll", ExactSpelling = true)]
public static extern uint NetApiBufferFree(IntPtr Buffer);

public enum NETSETUP_JOIN_STATUS
{
    NetSetupUnknownStatus = 0,
    NetSetupUnjoined,
    NetSetupWorkgroupName,
    NetSetupDomainName
}
'@

    if (-not ('Win32.NetAPI' -as [type])) {
        Add-Type -MemberDefinition $def -Namespace 'Win32' -Name 'NetAPI'
    }

    [IntPtr]$pName = [IntPtr]::Zero
    [uint32]$status = 0

    $sc = [Win32.NetAPI]::NetGetJoinInformation($null, [ref]$pName, [ref]$status)

    if ($sc -ne 0) {
        Write-Error "NetGetJoinInformation failed with $sc." -Exception (New-Object ComponentModel.Win32Exception($sc))
        return
    }

    $name = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pName)

    $sc = [Win32.NetAPI]::NetApiBufferFree($pName)
    if ($sc -ne 0) {
        Write-Error "NetApiBufferFree failed with $sc." -Exception (New-Object ComponentModel.Win32Exception($sc))
        return
    }

    [PSCustomObject]@{
        Name = $name
        JoinStatus = [Enum]::GetName([Win32.NetAPI+NETSETUP_JOIN_STATUS], $status)
    }
}

function Get-OutlookProfile {
    [CmdletBinding()]
    param(
        [string]$User
    )

    if (-not $User) {
        $User = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }

    $userRegRoot = Get-UserRegistryRoot $User
    if (-not $userRegRoot) {
        return
    }

    # List Outlook "profiles" keys
    $profiles = @(
        $versionKeys = @(Get-ChildItem (Join-Path $userRegRoot 'Software\Microsoft\Office\') -ErrorAction SilentlyContinue | Where-Object {$_.Name -match '\d\d\.0'})
        if ($versionKeys.Count) {
            foreach ($versionKey in $versionKeys) {
                Get-ChildItem (Join-Path $versionKey.PsPath '\Outlook\Profiles') -ErrorAction SilentlyContinue
                $versionKey.Close()
            }
        }
    )

    if (-not $profiles.Count) {
        Write-Log "There are no profiles for $User."
        return
    }

    $PR_PROFILE_CONFIG_FLAGS = '00036601'

    # Flag constants:
    # This corresponds [Use Cached Exchange Mode]
    $CONFIG_OST_CACHE_PRIVATE = 0x180
    # This corresponds to [Download Public Folder Favorites]
    $CONFIG_OST_CACHE_PUBLIC = 0x400
    # This corresponds to [Download shared folders]
    $CONFIG_OST_CACHE_DELEGATE_PIM = 0x800

    foreach ($profile in $profiles) {
        $flags = 0
        $CACHE_PRIVATE = $false
        $CACHE_PUBLIC = $false
        $CACHE_DELEGATE_PIM = $false

        $subkeys = Get-ChildItem $profile.PSPath -Recurse

        foreach ($subkey in $subkeys) {
            if ($subkey.Property | Where-Object {$_ -eq $PR_PROFILE_CONFIG_FLAGS}) {
                $bytes = $subkey.GetValue($PR_PROFILE_CONFIG_FLAGS)
                $flags = [BitConverter]::ToUInt32($bytes, 0)
                break
            }
        }

        # Close all the sub keys
        $subkeys | ForEach-Object {$_.Close()}

        if (($flags -band $CONFIG_OST_CACHE_PRIVATE) -ne 0) {
            $CACHE_PRIVATE = $true
        }

        if (($flags -band $CONFIG_OST_CACHE_PUBLIC) -ne 0) {
            $CACHE_PUBLIC = $true
        }

        if (($flags -band $CONFIG_OST_CACHE_DELEGATE_PIM) -ne 0) {
            $CACHE_DELEGATE_PIM = $true
        }

        [PSCustomObject]@{
            User = $User
            Profile = $profile.Name
            CachedMode = $CACHE_PRIVATE -or $CACHE_PUBLIC -or $CACHE_DELEGATE_PIM
            DownloadPublicFolderFavorites = $CACHE_PUBLIC
            DownloadSharedFolders = $CACHE_DELEGATE_PIM
            PR_PROFILE_CONFIG_FLAGS = $flags
        }

        $profile.Close()
    }
}

function Save-CachedAutodiscover {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        # Where to save
        $Path,
        # Target user name
        [string]$User
    )

    # Check %LOCALAPPDATA%\Microsoft\Outlook
    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $cachePath = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook'
    }
    else {
        return
    }

    if (-not (Test-Path $cachePath)) {
        Write-Log "$cachePath is not found."
        return
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    Write-Log "Searching $cachePath."

    # Get Autodiscover XML files and copy them to Path
    try {
        Get-ChildItem $cachePath -Filter '*Autod*.xml' -Force -Recurse | Copy-Item -Destination $Path
    }
    catch {
        # Just in case Copy-Item throws a terminating error.
        Write-Error -ErrorRecord $_
    }

    # Remove Hidden attribute
    foreach ($file in @(Get-ChildItem $Path -Force)) {
        if ((Get-ItemProperty $file.FullName).Attributes -band [IO.FileAttributes]::Hidden) {
            try {
                # Unfortunately, this does not work before PowerShellV5.
                (Get-ItemProperty $file.FullName).Attributes -= 'Hidden'
                continue
            }
            catch {
                # ignore error
            }

            # This could fail if attributes other than Archive, Hidden, Normal, ReadOnly, or System are set (such as NotContentIndexed)
            Set-ItemProperty $file.Fullname -Name Attributes -Value ((Get-ItemProperty $file.FullName).Attributes -bxor [IO.FileAttributes]::Hidden)
        }
    }
}

function Remove-CachedAutodiscover {
    [CmdletBinding()]
    param(
        # Target user name
        [string]$User
    )

     # Check %LOCALAPPDATA%\Microsoft\Outlook
    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $cachePath = Join-Path $localAppdata -ChildPath 'Microsoft\Outlook'
    }
    else {
        return
    }

    if (-not (Test-Path $cachePath)) {
        Write-Log "$cachePath is not found."
        return
    }

    Write-Log "Searching $cachePath."

    # Remove Autodiscover XML files
    Get-ChildItem $cachePath -Filter '*Autod*.xml' -Force -Recurse | ForEach-Object {
        Remove-Item $_.FullName -Force
    }
}

function Start-LdapTrace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, HelpMessage = "Directory for output file")]
        $Path,
        [parameter(Mandatory=$true, HelpMessage = "Process name to trace. e.g. Outlook.exe")]
        $TargetProcess,
        $SessionName = 'LdapTrace'
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path

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
        Write-Error "Failed to create the key under $keypath. Make sure to run as an administrator"
        return
    }

    # Start ETW session
    $traceFile = Join-Path $Path -ChildPath "ldap_%d.etl"
    $logFileMode = "globalsequence | EVENT_TRACE_FILE_MODE_NEWFILE"
    Write-Log "Starting a LDAP trace"

    $err = $($stdout = Invoke-Command {
        $ErrorActionPreference = 'Continue'
        & logman.exe create trace $SessionName -ow -o $traceFile -p Microsoft-Windows-LDAP-Client 0x1a59afa3 0xff -bs 1024 -mode $logFileMode -max 256 -ets
    }) 2>&1

    if ($err -or $LASTEXITCODE -ne 0) {
        Write-Error "Failed to start LDAP trace. exit code:$LASTEXITCODE; stdout:`"$stdout`"; error:`"$err`""
        return
    }
}

function Stop-LdapTrace {
    [CmdletBinding()]
    param(
        $SessionName = 'LdapTrace',
        [Parameter(Mandatory = $true)]
        $TargetProcess
    )

    Write-Log "Stopping $SessionName"
    Stop-EtwSession $SessionName | Out-Null

    # Remove a registry key under HKLM\SYSTEM\CurrentControlSet\Services\ldap\tracing (ignore any errors)

    # Process name must contain the extension such as "Outlook.exe", instead of "Outlook"
    if ([IO.Path]::GetExtension($TargetProcess) -ne 'exe') {
        $TargetProcess = [IO.Path]::GetFileNameWithoutExtension($TargetProcess) + ".exe"
    }

    $keypath = "HKLM:\SYSTEM\CurrentControlSet\Services\ldap\tracing\$TargetProcess"
    Remove-Item $keypath -ErrorAction SilentlyContinue | Out-Null
}

function Save-OfficeModuleInfo {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        $Path,
        # filter items by their name using -match (e.g. 'outlook.exe','mso\d\d.*\.dll'). These are treated as "OR".
        [string[]]$Filters,
        [Threading.CancellationToken]$CancellationToken
    )

    if (-not (Test-Path $Path)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    # If MS Office is not installed, bail.
    $officeInfo = Get-OfficeInfo -ErrorAction SilentlyContinue
    if (-not $officeInfo) {
        Write-Error "MS Office is not installed."
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

    Write-Log "officePaths are $officePaths"

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # Get exe and dll
    # It's slightly faster to run gci twice with -Filter than running once with -Include *exe, *.dll
    $items = @(
        foreach ($officePath in $officePaths) {
            Get-ChildItem -Path $officePath -Filter *.exe -Recurse -ErrorAction SilentlyContinue
            Get-ChildItem -Path $officePath -Filter *.dll -Recurse -ErrorAction SilentlyContinue
        }
    )

    $listingFinished = $sw.Elapsed
    Write-Log "Listing $($items.Count) items took $($listingFinished.TotalMilliseconds) ms."

    # Apply filters
    if ($Filters.Count) { # This is for PowerShell v2. PSv2 iterates a null collection.
        $items = $items | Where-Object {
            foreach ($filter in $Filters) {
                if ($_.Name -match $filter) {
                    $true
                    break
                }
            }
        }
    }

    $cmdletName = $PSCmdlet.MyInvocation.MyCommand.Name
    $name = $cmdletName.Substring($cmdletName.IndexOf('-') + 1)

    @(
    foreach ($item in $items) {
        if ($item.VersionInfo.FileVersionRaw) {
            $fileVersion = $item.VersionInfo.FileVersionRaw
        }
        else {
            $fileVersion = $item.VersionInfo.FileVersion
        }

        [PSCustomObject]@{
            Name = $item.Name
            FullName = $item.FullName
            #VersionInfo = $item.VersionInfo # too much info and FileVersionRaw is harder to find
            FileVersion = $fileVersion
        }

        if ($CancellationToken.IsCancellationRequested) {
            Write-Log "Cancel request acknowledged."
            break
        }

    }) | Export-Clixml -Depth 4 -Path $(Join-Path $Path -ChildPath "$name.xml") #-Encoding UTF8

    $sw.Stop()
    Write-Log "Enumerating items took $(($sw.Elapsed - $listingFinished).TotalMilliseconds) ms."
}

<#
This is an old implementation using a PowerShell Job.
Not used currently but I'm keeping it for a reference in future development.
This uses a named kernel event object for inter PS process (i.e. Job) communication.
#>
function Start-SavingOfficeModuleInfo_PSJob {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path,
        # filter items by their name using -match (e.g. 'outlook.exe','mso\d\d.*\.dll'). These are treated as "OR".
        [string[]]$Filters
    )

    if (-not (Test-Path $Path)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path

    # If MS Office is not installed, bail.
    $officeInfo = Get-OfficeInfo -ErrorAction SilentlyContinue
    if (-not $officeInfo) {
        Write-Error "MS Office is not installed."
        return
    }

    # Create a named event for inter PS process communication
    $eventName = [Guid]::NewGuid().ToString()
    $namedEvent = New-Object System.Threading.EventWaitHandle($false, [Threading.EventResetMode]::ManualReset, $eventName)

    $job =
    Start-Job -ScriptBlock {
        param (
            $Path,
            $OfficeInfo,
            $Filters,
            $OutputFileName,
            $EventName
        )

        $namedEvent = [System.Threading.EventWaitHandle]::OpenExisting($EventName)

        $officePaths = @(
            $OfficeInfo.InstallPath

            if ($env:CommonProgramFiles) {
                Join-Path $env:CommonProgramFiles 'microsoft shared'
            }

            if (${env:CommonProgramFiles(x86)}) {
                Join-Path ${env:CommonProgramFiles(x86)} 'microsoft shared'
            }
        )

        # Get exe and dll
        # It's slightly faster to run gci twice with -Filter than running once with -Include *exe, *.dll
        $items = @(
            foreach ($officePath in $officePaths) {
                Get-ChildItem -Path $officePath -Filter *.exe -Recurse -ErrorAction SilentlyContinue
                Get-ChildItem -Path $officePath -Filter *.dll -Recurse -ErrorAction SilentlyContinue
            }
        )

        # Apply filters
        if ($Filters.Count) { # This is for PowerShell v2. PSv2 iterates a null collection.
            $items = $items | Where-Object {
                foreach ($filter in $Filters) {
                    if ($_.Name -match $filter) {
                        $true
                        break
                    }
                }
            }
        }

        @(
        foreach ($item in $items) {
            if ($item.VersionInfo.FileVersionRaw) {
                $fileVersion = $item.VersionInfo.FileVersionRaw
            }
            else {
                $fileVersion = $item.VersionInfo.FileVersion
            }

            [PSCustomObject]@{
                Name = $item.Name
                FullName = $item.FullName
                #VersionInfo = $item.VersionInfo # too much info and FileVersionRaw is harder to find
                FileVersion = $fileVersion
            }

            #  If event is signaled, finish
            if ($namedEvent.WaitOne(0)) {
                break
            }
        }) | Export-Clixml -Depth 4 -Path $(Join-Path $Path -ChildPath "$OutputFileName.xml") -Encoding UTF8

        $namedEvent.Close()

    } -ArgumentList $Path, $officeInfo, $Filters, 'OfficeModuleInfo', $eventName

    [PSCustomObject]@{
        Job = $job
        Event = $namedEvent # To be closed by Stop-SavingOfficeModuleInfo_PSJob
    }

    Write-Log "Job (ID: $($job.Id)) has started. A Named Event (Handle: $($namedEvent.Handle), Name: '$eventName') is created"
}

<#
This is an old implementation using a PowerShell Job. Counterpart of Start-SavingOfficeModuleInfo_PSJob
Not used currently but I'm keeping it for a reference in future development.
#>
function Stop-SavingOfficeModuleInfo_PSJob {
    [CmdletBinding()]
    param(
        # Returned from Start-SavingOfficeModuleInfo_PSJob
        [Parameter(Mandatory = $true, ValueFromPipeline=$true)]
        $JobDescriptor,

        # Number of seconds to wait for the job.
        # Default value is -1 and this will block till the job completes
        [int]$TimeoutSecond = -1
    )

    process {
        $job = $JobDescriptor.job
        $namedEvent = $JobDescriptor.Event

        # Wait for the job up to timeout
        Write-Log "Waiting for the job (ID: $($job.Id)) up to $TimeoutSecond seconds."
        if (Wait-Job -Job $job -Timeout $TimeoutSecond) {
            # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/wait-job
            # > This cmdlet returns job objects that represent the completed jobs. If the wait ends because the value of the Timeout parameter is exceeded, Wait-Job does not return any objects.
            Write-Log "Job was completed."
        }
        else {
            Write-Log "Job did not complete. It will be stopped by event signal."
        }

        # Signal the event and close
        try {
            $namedEvent.Set() | Out-Null
            $namedEvent.Close()
            Write-Log "Event (Handle: $($namedEvent.Handle)) was closed."
        }
        catch {
            Write-Error -ErrorRecord $_
        }

        # Let the job finish
        Wait-Job -Job $job | Out-Null
        Stop-Job -Job $job
        # Receive-Job -Job $job
        Remove-Job -Job $job
        Write-Log "Job (ID: $($job.Id)) was removed."
    }
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
        Write-Error -Message "Failed to start $processName.`n$_" -Exception $_.Exception
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
    Write-Log "Starting a CAPI trace"
    $logmanResult = Invoke-Expression "logman create trace $SessionName -ow -o `"$traceFile`" -p `"Security: SChannel`" 0xffffffffffffffff 0xff -bs 1024 -mode `"$logFileMode`" -max 256 -ets"

    if ($LASTEXITCODE -ne 0) {
        Write-Error "logman failed to create a session. exit code = $LASTEXITCODE. $logmanResult"
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

    Write-Log "Stopping $SessionName"
    Stop-EtwSession $SessionName | Out-Null
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

    #  FiddlerCap is not available.
    if (-not (Test-Path $fiddlerExe)) {
        $fiddlerCapUrl = "https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe"
        $fiddlerSetupFile = Join-Path $Path -ChildPath 'FiddlerCapSetup.exe'

        # Check if FiddlerCapSetup.exe is already available locally; Otherwize download the setup file and extract it.
        if (-not (Test-Path $fiddlerSetupFile)) {
            # If it's not connected to internet, bail.
            $connectivity = Get-NLMConnectivity
            if (-not $connectivity.IsConnectedToInternet) {
                Write-Error "It seems there is no connectivity to Internet. Please download FiddlerCapSetup.exe from `"$fiddlerCapUrl`" and place it `"$Path`". Then run again."
                return
            }

            Write-Log "Downloading FiddlerCapSetup.exe"
            $webClient = $null
            try {
                $webClient = New-Object System.Net.WebClient
                Write-Progress -Activity "Downloading FiddlerCap" -Status "Please wait" -PercentComplete -1
                $webClient.DownloadFile($fiddlerCapUrl, $fiddlerSetupFile)
            }
            catch {
                Write-Error -Message "Failed to download FiddlerCapSetup from $fiddlerCapUrl. $_" -Exception $_.Exception
                return
            }
            finally {
                Write-Progress -Activity "Downloading FiddlerCap" -Status "Done" -Completed

                if ($webClient) {
                    $webClient.Dispose()
                }
            }
        }

        # Silently extract. Path must be absolute.
        $process = $null
        try {
            Write-Log "Extracting from FiddlerCapSetup"
            Write-Progress -Activity "Extracting from FiddlerCapSetup" -Status "This may take a while. Please wait" -PercentComplete -1

            # To redirect & capture error even when this cmdlet is called with ErrorAction:SilentlyContinue, need "Continue" error action.
            # Usually you can simply specify ErrorAction:Continue to the cmdlet. However, Start-Process does not respect that. So, I need to manually set $ErrorActionPreference here.
            $err = $($process = Invoke-Command {
                $ErrorActionPreference = "Continue"
                Start-Process $fiddlerSetupFile -ArgumentList "/S /D=$fiddlerPath" -Wait -PassThru
            }) 2>&1

            if ($process.ExitCode -ne 0) {
                Write-Error "Failed to extract $fiddlerExe. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
                return
            }
        }
        finally {
            Write-Progress -Activity "Extracting from FiddlerCapSetup" -Status "Done" -Completed
            if ($process) {
                $process.Dispose()
            }
        }
    }

    # Start FiddlerCap.exe
    $process = $null
    try {
        Write-Log "Starting FiddlerCap"
        $err = $($process = Invoke-Command {
            $ErrorActionPreference = "Continue"
            try {
                Start-Process $fiddlerExe -PassThru
            }
            catch {
                Write-Error -ErrorRecord $_
            }
        }) 2>&1

        if (-not $process -or $process.HasExited) {
            Write-Error "FiddlerCap failed to start or prematurely exited. $(if ($null -ne $process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
            return
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    [PSCustomObject]@{
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
    $procmonFile = $null

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

    if (-not ($procmonFile -and (Test-Path $procmonFile))) {
        # If 'ProcessMonitor.zip' isn't there, download.
        $procmonDownloadUrl = 'https://download.sysinternals.com/files/ProcessMonitor.zip'
        $procmonFolderPath = Join-Path $Path -ChildPath 'procmon_temp'
        $procmonZipFile = Join-Path $procmonFolderPath -ChildPath 'ProcessMonitor.zip'

        # If it's not connected to internet, bail.
        $connectivity = Get-NLMConnectivity
        if (-not $connectivity.IsConnectedToInternet) {
            Write-Error "It seems there is no connectivity to Internet. Please download the ProcessMonitor from `"$procmonDownloadUrl`""
            return
        }

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
            Write-Log "Downloading procmon"
            Write-Progress -Activity "Downloading procmon from $procmonDownloadUrl" -Status "Please wait" -PercentComplete -1
            $webClient = $null
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($procmonDownloadUrl, $procmonZipFile)
                $procmonZipDownloaded = $true
            }
            catch {
                Write-Error -Message "Failed to download procmon from $procmonDownloadUrl. $_" -Exception $_.Exception
                return
            }
            finally {
                if ($webClient) {
                    $webClient.Dispose()
                }

                Write-Progress -Activity "Downloading procmon from $procmonDownloadUrl" -Status "Done" -Completed
            }
        }

        # Unzip ProcessMonitor.zip
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            $NETFileSystemAvailable = $true
        }
        catch {
            Write-Log "System.IO.Compression.FileSystem isn't found. Using alternate method"
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
    Write-Log "Starting procmon"
    $process = $null
    $err = $($process = Invoke-Command {
        $ErrorActionPreference = "Continue"
        try {
            Start-Process $procmonFile -ArgumentList "/AcceptEula /Minimized /Quiet /NoFilter /BackingFile `"$pmlFile`"" -PassThru
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }) 2>&1

    try {
        if (-not $process -or $process.HasExited) {
            Write-Error "procmon failed to start or prematurely exited. $(if ($process.ExitCode) {"exit code = $($process.ExitCode)."}) $err"
            return
        }
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
    }

    Write-Log "Procmon successfully started"
    [PSCustomObject]@{
        ProcmonPath = $procmonFile
        ProcmonProcessId = $process.Id
        PMLFile = $pmlFile
        ProcmonZipDownloaded = $procmonZipDownloaded
        ProcmonFolderPath = $procmonFolderPath
    }
}

function Stop-Procmon {
    [CmdletBinding()]
    param()

    $process = @(Get-Process -Name Procmon*)
    if ($process.Count -eq 0) {
        Write-Error "Cannot find procmon or procmon64"
        return
    }

    $procmonFile = $process[0].Path
    foreach ($p in $process) {
        $p.Dispose()
    }

    # Stop procmon
    Write-Log "Stopping procmon"
    Write-Progress -Activity "Stopping procmon" -Status "Please wait" -PercentComplete -1
    $process = $null
    try {
        $err = $($process = Invoke-Command {
            $ErrorActionPreference = "Continue"
            try {
                Start-Process $procmonFile -ArgumentList "/Terminate" -Wait -PassThru
            }
            catch {
                Write-Error -ErrorRecord $_
            }
        }) 2>&1

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
        [string]$User
    )

    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $majorVersion = $officeInfo.Version.Split('.')[0]

    # Create registry key & values. Ignore errors (might fail due to existing values)
    $userRegRoot = Get-UserRegistryRoot -User $User -ErrorAction Stop
    $keypath = Join-Path $userRegRoot "Software\Microsoft\Office\$majorVersion.0\Common\Debug"

    Write-Log "Using $keypath."

    if (-not (Test-Path $keypath)) {
        New-Item $keypath -ErrorAction Stop | Out-Null
    }

    Write-Log "Starting a TCO trace by setting up $keypath"
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
        $Path,
        [string]$User
    )

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    $officeInfo = Get-OfficeInfo -ErrorAction Stop
    $majorVersion = $officeInfo.Version.Split('.')[0]

    # Remove registry values
    $userRegRoot = Get-UserRegistryRoot -User $User -ErrorAction Stop
    $keypath = Join-Path $userRegRoot "Software\Microsoft\Office\$majorVersion.0\Common\Debug"

    if (-not (Test-Path $keypath)) {
        Write-Warning "$keypath does not exist"
        return
    }

    Write-Log "Stopping a TCO trace by removing TCOTrace & MsoHttpVerbose from $keypath"
    Remove-ItemProperty $keypath -Name 'TCOTrace' -ErrorAction SilentlyContinue | Out-Null
    Remove-ItemProperty $keypath -Name 'MsoHttpVerbose' -ErrorAction SilentlyContinue | Out-Null

    # TCO Trace logs are in %TEMP%
    foreach ($item in @(Get-ChildItem -Path "$env:TEMP\*" -Include "office.log", "*.exe.log")) {
        try {
            Copy-Item $item -Destination $Path
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
Start tttracer.exe to launch a given executable.
#>
function Start-TTD {
    [CmdletBinding()]
    param(
        # Folder to save to.
        [Parameter(Mandatory=$true)]
        $Path,
        # Executable path (e.g. C:\Windows\System32\notepad.exe)
        [Parameter(Mandatory=$true)]
        $Executable
    )

    # Check if tttracer.exe is available (Win10 RS5 and above should include it)
    if (-not ($tttracer = Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "tttracer.exe is not available."
        return
    }

    # Make sure $Executable exists.
    if (-not (Test-Path $Executable)) {
        Write-Error "Cannot find $Executable."
        return
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    # Form the output file name.
    $targetName = [IO.Path]::GetFileNameWithoutExtension($Executable)
    $outPath = Join-Path $Path "$($targetName)_$(Get-Date -Format "yyyyMMdd_HHmmss")"

    $stdout = Join-Path $Path 'stdout.txt'
    $stderr = Join-Path $Path 'stderr.txt'

    Write-Log "TTD launching $Executable."

    $err = $($process = Invoke-Command {
        $ErrorActionPreference = 'Continue'
        Start-Process $tttracer -ArgumentList "-out `"$outPath`"", "`"$Executable`"" -PassThru  -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    }) 2>&1

    if (-not $process -or $process.HasExited) {
        Write-Error "tttracer.exe failed to start. ExitCode: $($process.ExitCode); Error: $err."
        $process.Dispose()
        return
    }

    # Find out the new process instantiated by tttracer.exe. This might take a bit.
    # The new process starts as a child process of tttracer.exe.
    $targetProcess = $null
    $maxRetry = 3
    foreach ($i in 1..$maxRetry) {
        if ($newProcess = Get-WmiObject Win32_Process -Filter "Name='$targetName.exe' AND ParentProcessId='$($process.Id)'")  {
            $targetProcess = Get-Process -Id $newProcess.ProcessId
            break
        }

        Start-Sleep -Seconds $i
    }

    if (-not $targetProcess) {
        Write-Error "Cannot find the new instance of $targetName."
        return
    }

    Write-Log "Target process $($targetProcess.Name) (PID: $($targetProcess.Id)) has started."

    [PSCustomObject]@{
        TTTracerProcess = $process
        TargetProcess = $targetProcess
        OutputFile = "$outPath.run"
    }
}

function Stop-TTD {
    [CmdletBinding()]
    param(
        # The returned object of Start-TTD
        [Parameter(Mandatory=$true)]
        $Descriptor,
        [switch]$KeepTargetProcess
    )

    $tttracerProcess = $Descriptor.TTTracerProcess
    $targetProcess = $Descriptor.TargetProcess

    if (-not (Get-Process -Id $targetProcess.Id -ErrorAction SilentlyContinue)) {
        Write-Log "Target process $($targetProcess.Name) (PID: $($targetProcess.Id)) does not exist."
        return
    }

    if (-not ($tttracer = Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "tttracer.exe is not available."
        return
    }

    if (-not ($tttracerProcess.ID -and $targetProcess.ID)) {
        Write-Error "Invalid input. tttracer PID: $($tttracerProcess.ID), target process PID: $($targetProcess.ID)"
        return
    }

    Write-Log "Stopping tttracer.exe. tttracer PID: $($tttracerProcess.ID), $($targetProcess.Name) PID: $($targetProcess.ID)."

    $message = & $tttracer -stop $targetProcess.ID
    $exitCode = $LASTEXITCODE

    # Wait-Process writes a non-terminating error when the process has exited. Ignore this error.
    $(Wait-Process -InputObject $tttracerProcess -ErrorAction SilentlyContinue) 2>&1 | Out-Null

    # Non zero exitcode indicates an error.
    if ($exitCode -ne 0 -or -not $tttracerProcess.HasExited) {
        Write-Error $("`"tttracer -stop`" failed. ExitCode: 0x{0:x}" -f $exitCode)
    }

    [PSCustomObject]@{
        ExitCode = $exitCode  # This is the exit code of "tttracer -stop"
        TTTracerExitCode = $tttracerProcess.ExitCode # This is the exit code of tttracer that has been attached to the target process. This may not be available.
        Message = $message
    }

    $tttracerProcess.Dispose()
    if (-not $KeepTargetProcess) {
        $targetProcess.Dispose()
    }
}

function Attach-TTD {
    [CmdletBinding()]
    param(
        # Folder to save to.
        [Parameter(Mandatory=$true)]
        $Path,
        # ProcessID of the process to attach to.
        [Parameter(Mandatory=$true)]
        $ProcessID
    )

    # Check if tttracer.exe is available (Win10 RS5 and above should include it)
    if (-not ($tttracer = Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        Write-Error "tttracer.exe is not available."
        return
    }

    if ($targetProcess = Get-Process -Id $ProcessID -ErrorAction SilentlyContinue) {
        $targetName = $targetProcess.Name
    }
    else {
        Write-Error "Cannot find a process with PID $ProcessID."
        return
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    # Form the output file name.
    $outPath = Join-Path $Path "$($targetName)_$(Get-Date -Format "yyyyMMdd_HHmmss")"

    # If a folder path is used, it must not end with "\". If that's the case remove it.
    # $outPath = $Path
    # if ($outPath.EndsWith([IO.Path]::DirectorySeparatorChar)) {
    #     $outPath = $outPath.Substring(0, $outPath.Length - 1)
    # }

    $stdout = Join-Path $Path 'stdout.txt'
    $stderr = Join-Path $Path 'stderr.txt'

    $err = $($process = Invoke-Command {
        $ErrorActionPreference = 'Continue'
        Start-Process $tttracer -ArgumentList "-out `"$outPath`"", "-attach $ProcessID" -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    }) 2>&1

    # Must wait for a little to see if tttracer succeeded.
    # Wait-Process writes a non-terminating error when timeout occurs. Note that timeout here is a good thing; tttracer is still running.
    $timeout = $(Wait-Process -InputObject $process -Timeout 3 <#seconds#> -ErrorAction Continue) 2>&1

    if ($timeout) {
        # Seems successful
        [PSCustomObject]@{
            TTTracerProcess = $process
            TargetProcess = $targetProcess
            OutputFile = "$outPath.run"
        }

        return
    }

    $stderrContent = Get-Content $stderr
    $exitCodeHex = "0x{0:x}" -f $process.ExitCode
    Write-Error "tttracer.exe failed to attach. ExitCode: $exitCodeHex; Error: $err.`n$stderrContent"
}

function Get-OfficeInfo {
    [CmdletBinding()]
    param(
        [switch]$IgnoreCache
    )

    # Use the cache if it's available
    if ($Script:OfficeInfoCache -and -not $IgnoreCache.IsPresent) {
        Write-Log "Returning a cache"
        return $Script:OfficeInfoCache
    }

    $officeInstallations = @(
        $hklm = $null
        try {
            if ('Microsoft.Win32.RegistryView' -as [type]) {
                Write-Log "Using OpenBaseKey with [Microsoft.Win32.RegistryView]::Registry64"
                $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64);
            }
            elseif (-not $env:PROCESSOR_ARCHITEW6432) {
                # RegistryView is not available, but it's OK because no WOW64.
                Write-Log "Using OpenRemoteBaseKey"
                $hklm = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[string]::Empty);
            }
            else {
                # This is the case where registry rediction takes place (32bit PowerShell on 64bit OS). Bail.
                Write-Error "32bit PowerShell 2.0 is running on 64bit OS. Please run 64bit PowerShell."
                return
            }

            $keysToSearch = @(
                'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                # 32bit MSI is under Wow6432Node.
                'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
            )

            foreach ($key in $keysToSearch) {
                $uninstallKey = $hklm.OpenSubKey($key)

                if ($null -eq $uninstallKey) {
                    continue
                }

                foreach ($subKeyName in $uninstallKey.GetSubKeyNames()) {
                    if ($null -eq $subKeyName) {
                        continue
                    }

                    $subKey = $uninstallKey.OpenSubKey($subKeyName)
                    $displayName = $subKey.GetValue('DisplayName')
                    $displayIcon = $subKey.GetValue('DisplayIcon')
                    $modifyPath =  $subKey.GetValue('ModifyPath')

                    if (($displayName -like "Microsoft Office*" -or $displayName  -like "Microsoft 365 Apps*") -and $displayIcon -and $modifyPath -notlike "*MUI*") {
                        [PSCustomObject]@{
                            Version = $subKey.GetValue('DisplayVersion')
                            Location = $subKey.GetValue('InstallLocation')
                            DisplayName = $displayName
                            ModifyPath = $modifyPath
                            DisplayIcon = $displayIcon
                        }
                    }
                    $subKey.Close()
                }

                $uninstallKey.Close()
            }
        }
        finally {
            if ($hklm) {
                $hklm.Close()
            }
        }
    )

    $displayName = $version = $installPath = $null

    if ($officeInstallations.Count -gt 0) {
        # There might be more than one version of Office installed.
        # Use the latest
        $latestOffice = $officeInstallations | Sort-Object -Property {[System.Version]$_.Version} -Descending | Select-Object -First 1
        $displayName = $latestOffice.DisplayName
        $version = $latestOffice.Version
        $installPath = $latestOffice.Location
    }
    else {
        Write-Log "Cannot find the Office installation from HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall. Fall back to HKLM:\SOFTWARE\Microsoft\Office"
        $keys =  @(Get-ChildItem HKLM:\SOFTWARE\Microsoft\Office\ | Where-Object {[RegEx]::IsMatch($_.PSChildName,'\d\d\.0') -or $_.PSChildName -eq 'ClickToRun' })

        # If 'ClickToRun' exists, use its "InstallPath" & "VersionToReport".
        $clickToRun = $keys | Where-Object {$_.PSChildName -eq 'ClickToRun'}
        if ($clickToRun) {
            $installPath = Get-ItemProperty $clickToRun.PSPath | Select-Object -ExpandProperty 'InstallPath'
            $version = Get-ItemProperty (Join-Path $clickToRun.PSPath 'Configuration')| Select-Object -ExpandProperty 'VersionToReport'
        }
        else {
            # Otherwise, check "Common\InstallRoot" key's "Path"
            foreach ($key in ($keys | Sort-Object -Property PSChildName -Descending)) {
                $installPath = Get-ItemProperty (Join-Path $key.PSPath 'Common\InstallRoot') -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'Path'
                if ($installPath) {
                    $version = $key.PSChildName
                    break
                }
            }
        }
    }

    if (-not $installPath){
        Write-Error "Microsoft Office is not installed"
        return
    }

    $outlookReg = Get-ItemProperty 'HKLM:\SOFTWARE\Clients\Mail\Microsoft Outlook' -ErrorAction SilentlyContinue
    if ($outlookReg) {
        $mapiDll = Get-ItemProperty $outlookReg.DLLPathEx -ErrorAction SilentlyContinue
    }

    $Script:OfficeInfoCache =
    [PSCustomObject]@{
        DisplayName = $displayName
        Version = $version
        InstallPath = $installPath
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
        Write-Log "$TargetProcess does not end with '.exe'.  Adding '.exe'"
        $TargetProcess += '.exe'
    }

    # Create a key 'LocalDumps' under HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps, if it doesn't exist
    $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'
    if (-not (Test-Path (Join-Path $werKey 'LocalDumps'))) {
        New-Item $werKey -Name 'LocalDumps' -ErrorAction Stop | Out-Null
    }

    # Create a ProcessName key under LocalDumps, if it doesn't exist
    $localDumpsKey = Join-Path $werKey 'LocalDumps'
    if (-not (Test-Path (Join-Path $localDumpsKey $TargetProcess))) {
        New-Item $localDumpsKey -Name $TargetProcess -ErrorAction Stop | Out-Null
    }

    # Create "CustomDumpFlags", "DumpType", and "DumpFolder" values in ProcessName key
    $ProcessKey = Join-Path $localDumpsKey $TargetProcess
    Write-Log "Setting up $ProcessKey with CustomDumpFlags:0x61826, DumpType:0, DumpFolder:$Path"
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
            Write-Log "Temporarily renaming $installedKey to `"_Installed`""
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
        Write-Log "$TargetProcess does not end with '.exe'.  Adding '.exe'"
        $TargetProcess += '.exe'
    }

    $werKey = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'
    $localDumpsKey = Join-Path $werKey 'LocalDumps'
    $ProcessKey = Join-Path $localDumpsKey $TargetProcess

    if (Test-Path $ProcessKey) {
        Write-Log "Removing $ProcessKey"
        Remove-Item $ProcessKey
    }
    else {
        Write-Error "$ProcessKey does not exist"
    }

    # Rename DW "_Installed" keys back to "Installed"
    $pcHealth = @(
        # For C2R
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
        'HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\_Installed'

        # For MSI
        'HKLM:\Software\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
        'HKLM:\Software\Wow6432Node\Microsoft\PCHealth\ErrorReporting\DW\_Installed'
    )

    foreach ($installedKey in $pcHealth) {
        if (Test-Path $installedKey) {
            Write-Log "Renaming $installedKey back to `"Installed`""
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
    [int]$IntervalSeconds,
    [TimeSpan]$MaxDuration = [TimeSpan]::FromHours(1)  # Just for safety, make sure to stop after a period
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType directory $Path -ErrorAction Stop | Out-Null
    }
    $Path = Resolve-Path $Path

    Write-Log "Starting a WFP job"
    $job = Start-Job -ScriptBlock {
        param($Path, $IntervalSeconds, $MaxDuration)

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
            netsh wfp show netevents file="$eventFilePath" <#timewindow=$IntervalSeconds#> | Out-Null
            Start-Sleep -Seconds $IntervalSeconds

        }
    } -ArgumentList $Path, $IntervalSeconds, $MaxDuration

    $job
}

function Stop-WfpTrace {
    [CmdletBinding()]
    [Parameter(Mandatory = $true)]
    param (
    $WfpJob
    )

    Write-Log "Stopping a WFP job"
    Stop-Job -Job $WfpJob
    Remove-Job -Job $WfpJob
}


function Save-Dump {
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)]
        $Path, # Folder to save a dump file
        [parameter(Mandatory = $true)]
        [int]$ProcessId
    )

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
    $DbgHelp = @'
    [DllImport("Dbghelp.dll", SetLastError=true)]
    public static extern bool MiniDumpWriteDump(
        IntPtr hProcess,
        uint ProcessId,
        IntPtr hFile,
        uint DumpType,
        IntPtr ExceptionParam,
        IntPtr UserStreamParam,
        IntPtr CallbackParam);
'@

    if (-not ('DbgHelp' -as [type])) {
        Add-type -MemberDefinition $DbgHelp -Name DbgHelp -Namespace Win32
    }

    if (-not (Test-Path $Path)) {
        New-Item $Path -ItemType Directory -ErrorAction Stop | Out-Null
    }

    $Path = Resolve-Path $Path
    $process = $null

    try{
        $process = Get-Process -Id $ProcessId -ErrorAction Stop
        if (-not $process.Handle) {
            # This scenario is possible for a system process.
            Write-Error "Cannot obtain the process handle of $($process.Name)."
            return
        }

        $dumpFile = Join-Path $Path "$($process.Name)_$(Get-Date -Format 'yyyy-MM-dd-HHmmss').dmp"
        $dumpFileStream = [System.IO.File]::Create($dumpFile)
        $writeDumpSuccess = $false

        Write-Log "Calling Win32 MiniDumpWriteDump"
        # Note: 0x61826 = MiniDumpWithTokenInformation | MiniDumpIgnoreInaccessibleMemory | MiniDumpWithThreadInfo (0x1000) | MiniDumpWithFullMemoryInfo (0x800) |MiniDumpWithUnloadedModules (0x20) | MiniDumpWithHandleData (4) | MiniDumpWithFullMemory (2)
        if ([Win32.DbgHelp]::MiniDumpWriteDump($process.Handle, $ProcessId, $dumpFileStream.Handle, 0x61826, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero)) {
            [PSCustomObject]@{
                ProcessID = $process.Id
                ProcessName = $process.Name
                DumpFile = $dumpFile
            }
            $writeDumpSuccess = $true
        }
        else {
            Write-Error ("Failed to save a memory dump of $Process. Error = 0x{0:x}" -f [System.Runtime.InteropServices.Marshal]::GetLastWin32Error())
        }
    }
    finally {
        if ($dumpFileStream) {
            $dumpFileStream.Close()

            if (-not $writeDumpSuccess) {
                Remove-Item $dumpFile -Force -ErrorAction SilentlyContinue
            }
        }

        if ($process) {
            $process.Dispose()
        }
    }
}

function Save-MSIPC {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Path,
        $User
    )

    # MSIPC info is in %LOCALAPPDATA%\Microsoft\MSIPC
    if ($localAppdata = Get-UserShellFolder -User $User -ShellFolderName 'Local AppData') {
        $msipcPath = Join-Path $localAppdata 'Microsoft\MSIPC'
    }
    else {
        return
    }

    if (-not (Test-Path $msipcPath)) {
        Write-Error "$msipcPath does not exist"
        return
    }

    if (-not (Test-Path $Path -ErrorAction Stop)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    # Copy only folders (i.e. skip drm files)
    # gci -Directory is only available for PowerShell V3 and above. To support PowerShell V2 clients, Where-Object is used here.
    foreach ($folder in @(Get-ChildItem $msipcPath | Where-Object {$_.PSIsContainer})) {
        $dest = Join-Path $Path $folder.Name

        if (-not (Test-Path $dest -ErrorAction Stop)){
            New-Item -ItemType Directory $dest -ErrorAction Stop | Out-Null
        }

        Write-Log "Copying $($folder.FullName) to $dest"
        try {
            # Copy-Item could throw a terminating error
            Copy-Item (Join-Path $folder.FullName '*') -Destination $dest -Recurse -ErrorAction SilentlyContinue
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
}

<#
.SYNOPSIS
This function returns an instance of Microsoft.Identity.Client.LogCallback delegate which calls the given scriptblock when LogCallback is invoked.
#>
function New-LogCallback {
    [CmdletBinding()]
    param (
    # Scriptblock to be called when MSAL invokes LogCallback
    [Parameter(Mandatory=$true)]
    [scriptblock]$Callback,

    # Remaining arguments to be passd to Callback scriptblock via $Event.MessageData
    [Parameter(ValueFromRemainingArguments = $true)]
    [object[]]$ArgumentList
    )

    # Class that exposes an event of type Microsoft.Identity.Client.LogCallback that Register-ObjectEvent can register to.
    $LogCallbackProxyType = @"
        using System;
        using System.Threading;
        using Microsoft.Identity.Client;

        public sealed class LogCallbackProxy
        {
            // This is the exposed event. The sole purpose is for Register-ObjectEvent to hook to.
            public event LogCallback Logging;

            // This is the LogCallback delegate instance.
            public LogCallback Callback
            {
                get { return new LogCallback(OnLogging); }
            }

            // Raise the event
            private void OnLogging(LogLevel level, string message, bool containsPii)
            {
                LogCallback temp = Volatile.Read(ref Logging);
                if (temp != null) {
                    temp(level, message, containsPii);
                }
            }
        }
"@

    if (-not ("LogCallbackProxy" -as [type])) {
        Add-Type $LogCallbackProxyType -ReferencedAssemblies (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.dll')
    }

    $proxy = New-Object LogCallbackProxy
    Register-ObjectEvent -InputObject $proxy -EventName Logging -Action $Callback -MessageData $ArgumentList | Out-Null

    $proxy.Callback
}

<#
.SYNOPSIS
Obtains a modern auth token (maybe from a cached one if available).

.NOTES
You need the following MSAL.NET modules under "modules" sub folder:

 [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)
 [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)

 Folder structure should look like this:

    SomeFolder
    |  OutlookTrace.psm1
    |
    |- modules
          Microsoft.Identity.Client.dll
          Microsoft.Identity.Client.Extensions.Msal.dll

Note about proxy:
MSAL.NET uses System.Net.Http.HttpClient when calling RequestBase.ResolveAuthorityEndpointsAsync(), which reaches "/common/discovery/instance?api-version=1.1&authorization_endpoint=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fv2.0%2Fauthorize".
(And this data is cached by Microsoft.Identity.Client.Instance.Discovery.NetworkCacheMetadataProvider in memory. And it won't be fetched next time).
And it also uses System.Windows.Forms.WebBrowser-derived class (named "CustomWebBrowser") when calling InteractiveRequest.GetTokenResponseAsync() to reach "authorize" endpoint
e.g. "/common/oauth2/v2.0/authorize?scope=openid+profile+offline_access&response_type=code&client_id=d3590ed6-52b3-4102-aeff-aad2292ab01d&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient...
I can provide the builder's WithHttpClientFactory() with a IMsalHttpClientFactory of a HttpClient with a specific proxy. However I don't think I can do the same for the CustomWebBrowser.
Thus, in order to use a consistent proxy, it's best to configure the user's default proxy in WinInet.

.LINK
[AzureAD/microsoft-authentication-library-for-dotnet](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet)
[AzureAD/microsoft-authentication-extensions-for-dotnet](https://github.com/AzureAD/microsoft-authentication-extensions-for-dotnet)

#>
function Get-Token {
    [CmdletBinding()]
    param(
    # Client ID (Application ID) of the registered application.
    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    # Tenant ID. By default, it uses '/common' endpoint for multi-tenant app. For a single-tenant app, specify the tenant name or GUID (e.g. "contoso.com", "contoso.onmicrosoft.com", "333b3ed5-0ac4-4e75-a1cd-db9e8f593ff3")
    [string]$TenantId = 'common',

    # Array of scopes to request.  By default, "openid", "profile", and "offline_access" are included.
    [string[]]$Scopes,

    # Refirect URI for the application. When this is not given, "https://login.microsoftonline.com/common/oauth2/nativeclient" will be used.
    # Make sure to use the same URI as the one registered for the application.
    [string]$RedirectUri,

    # Clear the cached token and force to get a new token.
    [switch]$ClearCache,

    # Enable MSAL logging. Log file will be msal.log under the script folder.
    [switch]$EnableLogging
    )

    # Need MSAL.NET DLL under modules
    # https://github.com/AzureAD/microsoft-authentication-library-for-dotnet
    # [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)
    if (-not ('Microsoft.Identity.Client.AuthenticationResult' -as [type])) {
        try {
            Add-Type -Path (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.dll')
        }
        catch {
            Write-Error -ErrorRecord $_
            return
        }
    }

    # [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)
    if (-not ('Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper' -as [type])) {
        try {
            Add-Type -Path (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.Extensions.Msal.dll')
        }
        catch {
            Write-Error -ErrorRecord $_
            return
        }
    }

    # Configure & create a PublicClientApplication
    $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithAuthority((New-Object System.Uri "https://login.microsoftonline.com/$TenantId/"), <#validateAuthority#> $false)

    if ($RedirectUri) {
        $builder.WithRedirectUri($RedirectUri) | Out-Null
    }
    else {
        # WithDefaultRedirectUri() makes the redirect_uri "https://login.microsoftonline.com/common/oauth2/nativeclient".
        # Without it, redirect_uri would be "urn:ietf:wg:oauth:2.0:oob".
        $builder.WithDefaultRedirectUri() | Out-Null
    }

    $writer = $null

    if ($EnableLogging) {
        $logFile = Join-Path (Split-Path $PSCommandPath) 'msal.log'
        [IO.StreamWriter]$writer = [IO.File]::AppendText($logFile)
        Write-Verbose "MSAL Loggin is enabled. Log file: $logFile"

        # Add a CSV header line
        $writer.WriteLine("datetime,level,containsPii,message");

        $builder.WithLogging(
            # Microsoft.Identity.Client.LogCallback
            (New-LogCallback {
                param([Microsoft.Identity.Client.LogLevel]$level, [string]$message, [bool]$containsPii)

                $writer = $Event.MessageData[0]
                $writer.WriteLine("$((Get-Date).ToString('o')),$level,$containsPii,`"$message`"")

            } -ArgumentList $writer),

            [Microsoft.Identity.Client.LogLevel]::Verbose,
            # enablePiiLogging
            $true,
            # enableDefaultPlatformLogging
            $false
        ) | Out-Null
    }

    $publicClient = $builder.Build()

    # Configure caching
    $cacheFileName = "msalcache.bin"
    $cacheDir = Split-Path $PSCommandPath
    $storagePropertiesBuilder = New-Object Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder($cacheFileName, $cacheDir, $ClientId)
    $storageProperties = $storagePropertiesBuilder.Build()
    $cacheHelper = [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::CreateAsync($storageProperties).GetAwaiter().GetResult()
    $cacheHelper.RegisterCache($publicClient.UserTokenCache)

    if ($ClearCache) {
        $cacheHelper.Clear()
    }

    # Get an account
    $firstAccount = $publicClient.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1

    # By default, MSAL asks for scopes: openid, profile, and offline_access.
    try {
        $publicClient.AcquireTokenSilent($Scopes, $firstAccount).ExecuteAsync().GetAwaiter().GetResult()
    }
    catch [Microsoft.Identity.Client.MsalUiRequiredException] {
        try {
            $publicClient.AcquireTokenInteractive($Scopes).ExecuteAsync().GetAwaiter().GetResult()
        }
        catch {
            Write-Error -ErrorRecord $_
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
    finally {
        if ($writer){
            $writer.Dispose()
        }
    }
}

<#
.SYNOPSIS
This function makes an Autodiscover request.
#>
function Test-Autodiscover {
    [CmdletBinding()]
    param(
    # Server to send an Autodiscover request. For Exchange Online, use 'outlook.office365.com'
    # When not specified, "autodiscover.{SMTP domain}" will be tried.
    [string]$Server,

    # Target Email address for Autodiscover
    [Parameter(Mandatory=$true)]
    [string]$EmailAddress,

    # Legacy auth credential.
    [Parameter(ParameterSetName='LegacyAuth', Mandatory=$true)]
    [System.Management.Automation.PSCredential]
    $Credential,

    # Modern auth access token.
    # To mock an Office client, use ClientId 'd3590ed6-52b3-4102-aeff-aad2292ab01c' and Scope 'https://outlook.office.com/.default'
    # e.g. Get-Token -ClientId 'd3590ed6-52b3-4102-aeff-aad2292ab01c' -Scopes 'https://outlook.office.com/.default' -RedirectUri 'urn:ietf:wg:oauth:2.0:oob'
    [Parameter(ParameterSetName='ModernAuth', Mandatory=$true)]
    [string]$Token,

    # Proxy Server
    # e.g. "http://myproxy:8080"
    [string]$Proxy,

    # Skip adding "X-MapiHttpCapability: 1" to the header
    [switch]$SkipMapiHttpCapability,

    # Force Basic auth
    [switch]$ForceBasicAuth
    )

    $body = @"
<?xml version="1.0" encoding="utf-8"?>
<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
    <Request>
    <EMailAddress>$EmailAddress</EMailAddress>
    <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
    </Request>
</Autodiscover>
"@

    $mailDomain = $EmailAddress.Substring($EmailAddress.IndexOf("@") + 1)

    # These are the URL to try (+ redirect URLs).
    # Note that the urls are tried in the reverse order because this is a stack.
    $urls = New-Object System.Collections.Generic.Stack[string](,[string[]]@(
        "http://autodiscover.$mailDomain/autodiscover/autodiscover.xml"
        "https://autodiscover.$mailDomain/autodiscover/autodiscover.xml"
        "https://$Server/autodiscover/autodiscover.xml"
    ))

    $step = 1

    while ($urls.Count -gt 0) {
        $url = $urls.Pop()

        # Check if URL is valid (it could be invalid if $Server is not provided).
        $uri = $null
        if (-not [Uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri)) {
            Write-Log "Skipping $url because it's invalid."
            continue
        }

        # Arguments for Invoke-WebRequest paramters
        if ($uri.Scheme -eq 'https') {
            $arguments = @{
                Method = 'POST'
                Uri = $uri
                Headers =  @{'Content-Type'='text/xml'}
                Body = $body
                UseBasicParsing = $true
            }

            switch -Wildcard ($PSCmdlet.ParameterSetName) {
                'LegacyAuth' {
                    Write-Verbose "Credential is provided. Use it for legacy auth"

                    if ($ForceBasicAuth) {
                        $base64Cred = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$($Credential.UserName):$($Credential.GetNetworkCredential().Password)"))
                        $arguments['Headers'].Add('Authorization',"Basic $base64Cred")
                    }
                    else {
                        $arguments['Credential'] = $Credential
                    }
                    break
                }

                'ModernAuth' {
                    Write-Verbose "Token is provided. Use it for modern auth"
                    $arguments['Headers'].Add('Authorization',"Bearer $Token")
                    break
                }
            }

            if (-not $SkipMapiHttpCapability) {
                $arguments['Headers'].Add('X-MapiHttpCapability','1')
            }
        }
        else {
            $arguments = @{
                Method = 'GET'
                Uri = $uri
                MaximumRedirection = 0 # Just get 302 and don't follow the redirect.
                UseBasicParsing = $true
            }
        }

        if ($Proxy) {
            $arguments['Proxy'] = $Proxy
        }

        # Make a web request.
        Write-Log "Trying $($arguments.Method) $($arguments.Uri)"
        $result = $null
        $err = $($result = Invoke-WebRequest @arguments) 2>&1

        # Check result
        if ($result.StatusCode -eq 200) {
            [PSCustomObject]@{
                Step = $step++
                URI = $uri
                Success = $true
                Result = $result
            }
            return
        }
        elseif ($uri.Scheme -eq 'http' -and $result.StatusCode -eq 302) {
            # See if we got 302 with Location header
            $redirectUrl = $null
            if ($result.StatusCode -eq 302) {
                $redirectUrl = $result.Headers['Location']
            }

            if ($redirectUrl) {
                $result | Add-Member -MemberType ScriptMethod -Name 'ToString' -Force -Value {"Received a redirect URL $($this.Headers['Location'])"}
                [PSCustomObject]@{
                    Step = $step++
                    URI = $uri
                    Success = $true
                    Result = $result
                }

                # Try the given redirect uri next
                Write-Log "Received a redirect URL: $redirectUrl"
                $urls.Push($redirectUrl)
            }
            else {
                [PSCustomObject]@{
                    Step = $step++
                    URI = $uri
                    Success = $false
                    Result = $err
                }
            }
        }
        else {
            [PSCustomObject]@{
                Step = $step++
                URI = $uri
                Success = $false
                Result = $err
            }
        }
    }
}

<#
.SYNOPSIS
Convert a ProgID to CLSID
#>
function ConvertTo-CLSID {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$ProgID,
        [string]$User
    )

    $def = @'
    [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern uint CLSIDFromProgID(string progOd, out Guid clsid);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern uint StringFromCLSID([MarshalAs(UnmanagedType.LPStruct)] Guid refclsid, out IntPtr pClsidString);
'@

    if (-not ('Win32.OLE' -as [type])) {
        Add-Type -MemberDefinition $def -Namespace Win32 -Name OLE -ErrorAction Stop
    }

    [uint32]$S_OK = 0

    [Guid]$CLSID = [Guid]::Empty
    [uint32]$hr = [Win32.OLE]::CLSIDFromProgID($ProgID, [ref]$CLSID)

    if ($hr -ne $S_OK) {
        Write-Verbose -Message $("CLSIDFromProgID for `"$ProgID`" failed with 0x{0:x}. Trying ClickToRun registry." -f $hr)

        $userRegRoot = Get-UserRegistryRoot -User $User

        $locations = @(
            # ClickToRun Registry & the user's Classes
            "Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\"
            (Join-Path $userRegRoot "SOFTWARE\Classes\")
        )

        foreach ($loc in $locations) {
            $clsidProp = Get-ItemProperty (Join-Path $loc "$ProgID\CLSID") -ErrorAction SilentlyContinue
            $curVerProp = Get-ItemProperty (Join-Path $loc "$ProgID\CurVer") -ErrorAction SilentlyContinue

            if ($clsidProp) {
                $CLSID = $clsidProp.'(default)'
                break
            }
            elseif ($curVerProp) {
                $curProgID = $curVerProp.'(default)'
                $clsidProp = Get-ItemProperty (Join-Path $loc "$curProgID\CLSID") -ErrorAction SilentlyContinue
                $CLSID = $clsidProp.'(default)'
                break
            }
        }

        if ($CLSID -eq [Guid]::Empty) {
            Write-Error -Message $("CLSIDFromProgID for `"$ProgID`" failed with 0x{0:x}. Also, it was found in the ClickToRun & user registry" -f $hr)
            return
        }
    }

    [IntPtr]$pClsIdString = [IntPtr]::Zero
    $hr = [Win32.OLE]::StringFromCLSID($CLSID, [ref]$pCLSIDString)

    if ($hr -eq $S_OK -and $pCLSIDString) {
        $CLSIDString = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($pCLSIDString)
        [System.Runtime.InteropServices.Marshal]::FreeCoTaskMem($pCLSIDString)
        $pCLSIDString = [IntPtr]::Zero
    }

    [PSCustomObject]@{
        GUID = $CLSID
        String = $CLSIDString
    }
}

<#
.SYNOPSIS
Get Outlook's COM addins
#>
function Get-OutlookAddin {
    [CmdletBinding()]
    param(
        # User name or SID
        [string]$User
    )

    $userRegRoot = Get-UserRegistryRoot $User
    if (-not $userRegRoot) {
        return
    }

    # Get keys under "Addins"
    $addinKeys = @(
        @(
            'Registry::HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins'
            'Registry::HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins'
            'Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins'
            'Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\Outlook\AddIns'
            Join-Path $userRegRoot 'Software\Microsoft\Office\Outlook\Addins'
        ) |
        ForEach-Object {
            Get-ChildItem $_ -ErrorAction SilentlyContinue
        }
     )

     $LoadBehavior = @{
        0 = 'None'
        1 = 'NoneLoaded'
        2 = 'StartupUnloaded'
        3 = 'Startup'
        8 = 'LoadOnDemandUnloaded'
        9 = 'LoadOnDemand'
        16 = 'LoadAtNextStartupOnly'
    }

     $cache = @{}

     foreach ($addin in $addinKeys) {
        $props = @{}
        $props['Path'] = $addin.Name

        $props['ProgID'] = $addin.PSChildName

        if ($cache.ContainsKey($props['ProgID'])) {
            Write-Log "Skipping $($props['ProgID']) because it's already found."
            continue
        }
        else {
            $cache.Add($props['ProgID'], $null)
        }

        $($clsid = ConvertTo-CLSID $props['ProgID'] -User $User -ErrorAction Continue) 2>&1 | Write-Log

        if ($clsid) {
            $props['CLSID'] = $clsid.String
        }
        else {
            continue
        }

        # ToDo: text might get garbled in DBCS environment.
        $props['Description'] = $addin.GetValue('Description')
        $props['FriendlyName'] = $addin.GetValue('FriendlyName')
        $loadBehaviorValue = $addin.GetValue('LoadBehavior')

        if (-not $loadBehaviorValue) {
            Write-Log "Skipping $($props['ProgID']) because its LoadBehavior is null."
            continue
        }
        else {
            $props['LoadBehavior'] = $LoadBehavior[$loadBehaviorValue]
        }

        $inproc32 = Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$($props['CLSID'])\InprocServer32" -ErrorAction SilentlyContinue
        if (-not $inproc32) {
            $inproc32 = Get-ItemProperty "Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes\CLSID\$($props['CLSID'])\InprocServer32" -ErrorAction SilentlyContinue
        }

        if ($inproc32) {
            $props['DLL'] = $inproc32.'(default)'
            $props['ThreadingModel'] = $inproc32.ThreadingModel
        }

        [PSCustomObject]$props
     }

    # Close all the keys
    $addinKeys | ForEach-Object {$_.Close()}
}

function Get-ClickToRunConfiguration {
    [CmdletBinding()]
    param()

    # Registry path is the for 32bit Office on 64bit OS.
    Get-ItemProperty Registry::HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration
}

function Get-DeviceJoinStatus {
    [CmdletBinding()]
    param()

    $dsregcmd = 'dsregcmd.exe'

    if (Get-Command $dsregcmd -ErrorAction SilentlyContinue) {
        & $dsregcmd /status
    }
    else {
        Write-Log "$dsregcmd is not available."
    }
}

<#
This function just starts C:\Windows\System32\gatherNetworkInfo.vbs and returns a process
#>
function Start-GatherNetworkInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
    }

    Invoke-Command {
        $ErrorActionPreference = 'Continue'
        Start-Process cscript.exe -ArgumentList "C:\windows\system32\gatherNetworkInfo.vbs" -WorkingDirectory $Path -PassThru
    }
}

function Stop-GatherNetworkInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Process,
        [TimeSpan]$Timeout
    )

    Write-Log "Waiting for gatherNetworkInfo.vbs to finish $(if ($Timeout) {"with timeout $Timeout"})"
    if ($null -eq $Timeout) {
        $completed = $process.WaitForExit()
    }
    else {
        $completed = $process.WaitForExit($Timeout.TotalMilliseconds)
    }

    if (-not $completed) {
        Write-Log "gatherNetworkInfo reached timeout $Timeout"
        $process.Kill()
    }

    $process.Dispose()
}

<#
Get processes and its user (only for Outlook.exe).
PowerShell 4's Get-Process has -IncludeUserName, but I'm using WMI here for now.
#>
function Save-Process {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        $Path
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
    }

    Write-Log "Saving Win32_Process"
    Get-WmiObject -Class Win32_Process | ForEach-Object {
        if ($_.ProcessName -eq 'Outlook.exe') {
            $owner = $_.GetOwner()
            $_ | Add-Member -MemberType NoteProperty -Name 'User' -Value "$($owner.Domain)\$($owner.User)"
        }
        $_
    } | Export-Clixml -Path (Join-Path $Path "Win32_Process_$(Get-Date -Format "yyyyMMdd_HHmmss").xml")
}

function Collect-OutlookInfo {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory=$true)]
        $Path,
        [Parameter(Mandatory=$true)]
        [ValidateSet('Outlook', 'Netsh', 'PSR', 'LDAP', 'CAPI', 'Configuration', 'Fiddler', 'TCO', 'Dump', 'CrashDump', 'Procmon', 'WAM', 'WFP', 'TTD')]
        [array]$Component,
        [ValidateSet('None', 'Mini', 'Full')]
        $NetshReportMode = 'Mini',
        [int]$DumpCount = 3,
        [int]$DumpIntervalSeconds = 60,
        [switch]$SkipZip
    )

    # Explicitly check admin rights
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Please run as administrator."
        return
    }

    if ($env:PROCESSOR_ARCHITEW6432 -and $PSVersionTable.PSVersion.Major -eq 2) {
        Write-Error "32bit PowerShell 2.0 is running on 64bit OS. Please use 64bit PowerShell."
        return
    }

    # MS Office must be installed to collect Outlook, TCO, or TTD.
    # This is just a fail fast. Start-OutlookTrace/TCOTrace fail anyway.
    if ($Component -contains 'Outlook' -or $Component -contains 'TCO' -or $Component -contains 'TTD') {
        $err = $(Get-OfficeInfo -ErrorAction Continue | Out-Null) 2>&1
        if ($err) {
            Write-Error "Component `"Outlook`" and/or `"TCO`" is specified, but installation of Microsoft Office is not found. $err"
            return
        }
    }

    if ($Component -contains 'TTD' -and -not (Get-Command 'tttracer.exe' -ErrorAction SilentlyContinue)) {
        $os = Get-WmiObject -Class 'Win32_OperatingSystem'
        Write-Error "tttracer is not available on this machine. $($os.Caption) ($($os.Version))."
        return
    }

    if (-not (Test-Path $Path -ErrorAction Stop)){
        New-Item -ItemType Directory $Path -ErrorAction Stop | Out-Null
    }

    Write-Log "Running as $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"

    # Create a temporary folder to store data.
    $Path = Resolve-Path $Path
    $tempPath = Join-Path $Path -ChildPath $([Guid]::NewGuid().ToString())
    New-Item $tempPath -ItemType directory -ErrorAction Stop | Out-Null

    # Start logging.
    Open-Log -Path (Join-Path $tempPath 'Log.txt') -ErrorAction Stop
    Write-Log "Script Version: $Script:Version (Module Version $($MyInvocation.MyCommand.Module.Version.ToString()))"
    Write-Log "PSVersion: $($PSVersionTable.PSVersion); CLRVersion: $($PSVersionTable.CLRVersion)"
    Write-Log "PROCESSOR_ARCHITECTURE: $env:PROCESSOR_ARCHITECTURE; PROCESSOR_ARCHITEW6432: $env:PROCESSOR_ARCHITEW6432"

    $sb = New-Object System.Text.StringBuilder
    foreach ($paramName in $PSBoundParameters.Keys) {
        $var = Get-Variable $paramName -ErrorAction SilentlyContinue
        if ($var) {
            $sb.Append("$($var.Name):$($var.Value); ") | Out-Null
        }
    }
    Write-Log "Parameters $($sb.ToString())"

    # To use Start-Task, make sure to open runspaces first and close it when finished.
    # Open-TaskRunspace -Variables (Get-Variable 'logWriter')
    Open-TaskRunspace -IncludeScriptVariables

    Write-Log "Starting traces"
    try {
        if ($Component -contains 'Configuration') {
            # Sub directories
            $ConfigDir = Join-Path $tempPath 'Configuration'
            $OSDir = Join-Path $ConfigDir 'OS'
            $OfficeDir =  Join-Path $ConfigDir 'Office'
            $RegistryDir = Join-Path $ConfigDir 'Registry'
            $NetworkDir = Join-Path $ConfigDir 'Network'
            $MSIPCDir = Join-Path $ConfigDir 'MSIPC'
            $EventDir = Join-Path $ConfigDir 'EventLog'

            Write-Progress -Activity "Saving configuration" -Status "Please wait" -PercentComplete 0
            New-Item -Path $ConfigDir -ItemType directory | Out-Null

            # First start tasks that might take a while.

            # MSInfo32 takes a long time. Currently disabled.
            # $msinfo32Task = Start-Task -Command 'Save-MSInfo32' -Parameters @{Path = $OSDir}

            Write-Log "Starting officeModuleInfoTask."
            $cts = New-Object System.Threading.CancellationTokenSource
            $officeModuleInfoTask = Start-Task {param($path, $token) Save-OfficeModuleInfo -Path $path -CancellationToken $token} -ArgumentList $OfficeDir, $cts.Token
            # $officeModuleInfoTask = Start-Task -Command 'Save-OfficeModuleInfo' -Parameters @{Path = $OfficeDir; CancellationToken = $cts.Token}
            # $Filters = 'outlook\.exe', 'umoutlookaddin\.dll', 'mso\.dll', 'mso\d\d.+\.dll', 'olmapi32\.dll', 'emsmdb32\.dll', 'wwlib\.dll'
            # Save-OfficeModuleInfo -Path (Join-Path $tempPath 'Configuration') -ErrorAction SilentlyContinue -Timeout 00:00:30

            Write-Log "Starting networkInfoTask."
            $networkInfoTask = Start-Task {param($path) Save-NetworkInfo -Path $path} -ArgumentList $NetworkDir
            # $networkInfoTask = Start-Task -Command 'Save-NetworkInfo' -Parameters @{Path = $NetworkDir}
            # Save-NetworkInfo -Path (Join-Path $tempPath 'Configuration\NetworkInfo') -ErrorAction SilentlyContinue
            # Save-NetworkInfoMT -Path (Join-Path $tempPath 'Configuration\NetworkInfo_MT') -ErrorAction SilentlyContinue

            $LogonUser = Get-LogonUser -ErrorAction SilentlyContinue

            Write-Log "Starting officeRegistryTask."
            $officeRegistryTask = Start-Task {param($path, $sid) Save-OfficeRegistry -Path $path -User $sid} -ArgumentList $RegistryDir, $LogonUser.SID
            # $officeRegistryTask = Start-Task -Command 'Save-OfficeRegistry' -Parameters @{Path = $RegistryDir; User = $LogonUser.SID}
            # Save-OfficeRegistry -Path (Join-Path $tempPath 'Configuration') -User $LogonUser.SID -ErrorAction SilentlyContinue

            Write-Log "Starting oSConfigurationTask."
            $oSConfigurationTask = Start-Task {param($path) Save-OSConfiguration -Path $path} -ArgumentList $OSDir
            # $oSConfigurationTask = Start-Task -Command 'Save-OSConfiguration' -Parameters @{Path = $OSDir}
            # Save-OSConfiguration -Path (Join-Path $tempPath 'Configuration')

            Write-Progress -Activity "Saving configuration" -Status "Please wait" -PercentComplete 20
            Run-Command {Get-OfficeInfo} -Path $OfficeDir
            Run-Command {param($LogonUser) Get-OutlookProfile -User $LogonUser.SID} -ArgumentList $LogonUser -Path $OfficeDir
            Run-Command {param($LogonUser) Get-OutlookAddin -User $LogonUser.SID} -ArgumentList $LogonUser -Path $OfficeDir
            Run-Command {Get-ClickToRunConfiguration} -Path $OfficeDir

            # $(Get-OfficeInfo | Export-Clixml -Path (Join-Path $OfficeDir 'OfficeInfo.xml')) 2>&1 | Write-Log
            # $(Get-OutlookProfile -User $LogonUser.SID | Export-Clixml -Path (Join-Path $OfficeDir 'OutlookProfile.xml')) 2>&1 | Write-Log
            # $(Get-OutlookAddin -User $LogonUser.SID | Export-Clixml -Path (Join-Path $OfficeDir 'OutlookAddin.xml')) 2>&1 | Write-Log
            # $(if ($o = Get-ClickToRunConfiguration) {$o | Export-Clixml -Path (Join-Path $OfficeDir 'ClickToRunConfiguration.xml')}) 2>&1 | Write-Log

            Write-Progress -Activity "Saving configuration" -Status "Please wait" -PercentComplete 40
            Run-Command {param($LogonUser, $OfficeDir) Save-CachedAutodiscover -User $LogonUser.Name -Path $(Join-Path $OfficeDir 'Cached AutoDiscover')} -ArgumentList $LogonUser, $OfficeDir
            #$(Save-CachedAutodiscover -User $LogonUser.Name -Path (Join-Path $OfficeDir 'Cached AutoDiscover')) 2>&1 | Write-Log

            Write-Progress -Activity "Saving configuration" -Status "Please wait" -PercentComplete 60
            Run-Command {param($LogonUser, $MSIPCDir) Save-MSIPC -Path $MSIPCDir -User $($LogonUser.SID)} -ArgumentList $LogonUser, $MSIPCDir
            #$(Save-MSIPC -Path $MSIPCDir -User $LogonUser.SID) 2>&1 | Write-Log

            Write-Progress -Activity "Saving configuration" -Status "Please wait" -PercentComplete 80
            Run-Command {param($OSDir) Save-Process -Path $OSDir} -ArgumentList $OSDir
            #$(Save-Process -Path $OSDir) 2>&1 | Write-Log

            if ($LogonUser) {
                $LogonUser | Export-Clixml -Path (Join-Path $OSDir 'LogonUser.xml')
            }

            Write-Progress -Activity "Saving configuration" -Status "Done" -Completed
        }

        if ($Component -contains 'Fiddler') {
            Start-FiddlerCap -Path $Path -ErrorAction Stop | Out-Null
            $fiddlerCapStarted = $true

            Write-Warning "FiddlerCap has started. Please manually configure and start capture."
        }

        if ($Component -contains 'Netsh') {
            # When netsh trace is run for the first time, it does not capture packets (even with "capture=yes").
            # To workaround, netsh is started and stopped immediately.
            $tempNetshName = 'netsh_test'
            Start-NetshTrace -Path (Join-Path $tempPath $tempNetshName) -FileName "$tempNetshName.etl" -RerpotMode 'None'
            Stop-NetshTrace
            Remove-Item (Join-Path $tempPath $tempNetshName) -Recurse -Force -ErrorAction SilentlyContinue

            Start-NetshTrace -Path (Join-Path $tempPath 'Netsh') -RerpotMode $NetshReportMode
            $netshTraceStarted = $true
        }

        if ($Component -contains 'Outlook') {
            # Stop a lingering session if any.
            Stop-OutlookTrace -ErrorAction SilentlyContinue
            Start-OutlookTrace -Path (Join-Path $tempPath 'Outlook')
            $outlookTraceStarted = $true
        }

        if ($Component -contains 'PSR') {
            Start-PSR -Path $tempPath #-ShowGUI
            $psrStarted = $true
        }

        if ($Component -contains 'LDAP') {
            Start-LDAPTrace -Path (Join-Path $tempPath 'LDAP') -TargetProcess 'Outlook.exe'
            $ldapTraceStarted = $true
        }

        if ($Component -contains 'CAPI') {
            Start-CAPITrace -Path (Join-Path $tempPath 'CAPI')
            $capiTraceStarted = $true
        }

        if ($Component -contains 'TCO') {
            Start-TCOTrace
            $tcoTraceStarted = $true
        }

        if ($Component -contains 'WAM') {
            Stop-WamTrace -ErrorAction SilentlyContinue
            Start-WamTrace -Path (Join-Path $tempPath 'WAM')
            $wamTraceStarted = $true
        }

        if ($Component -contains 'Procmon') {
            $procmonResult = Start-Procmon -Path (Join-Path $tempPath 'Procmon') -ProcmonSearchPath $Path -ErrorAction Stop
            $procmonStared = $true
        }

        if ($Component -contains 'WFP') {
            $wfpJob = Start-WfpTrace -Path (Join-Path $tempPath 'WFP') -IntervalSeconds 15
            $wfpStarted = $true
        }

        if ($Component -contains 'CrashDump') {
            Add-WerDumpKey -Path (Join-Path $tempPath 'WerDump') -TargetProcess 'Outlook.exe'
            $crashDumpStarted = $true
        }

        if ($Component -contains 'Dump') {
            $process = Get-Process -Name 'Outlook' -ErrorAction Stop

            for ($i = 0; $i -lt $DumpCount; $i++) {
                Write-Progress -Activity "Saving a memory dump of Outlook ($i/$DumpCount)." -Status "Please wait." -PercentComplete -1
                $dumpResult = Save-Dump -Path (Join-Path $tempPath 'Dump') -ProcessId $process.Id
                Write-Progress -Activity "Saving a memory dump of Outlook ($i/$DumpCount)." -Status "Done" -Completed
                Write-Log "Saved dump file: $($dumpResult.DumpFile)"

                # If there are more dumps to save, wait.
                if ($i -lt ($DumpCount - 1)) {
                    $secondsRemaining = $DumpIntervalSeconds
                    while ($secondsRemaining -gt 0) {
                        Write-Progress -Activity "Waiting $DumpIntervalSeconds seconds till next dump ($($i + 1)/$DumpCount done)." -Status "Please wait." -SecondsRemaining $secondsRemaining
                        Start-Sleep -Seconds 1
                        $secondsRemaining-=1
                    }
                }
            }
        }

        if ($Component -contains 'TTD') {
            # If Outlook is already running, attach to it. Otherwise, let tttracer launch it.
            if ($outlookProcess = Get-Process -Name 'Outlook' -ErrorAction SilentlyContinue) {
                Write-Log "TTD attaching to Outlook (PID $($outlookProcess.Id))."
                $ttd = Attach-TTD -Path (Join-Path $tempPath 'TTD')  -ProcessID  $outlookProcess.Id -ErrorAction Stop
            }
            else {
                $outlookExe = $null
                $officeInfo = Get-OfficeInfo
                $executables = @(Get-ChildItem -Path $officeInfo.InstallPath -Filter 'Outlook.exe' -File -Recurse)

                if ($executables.Count -eq 1) {
                    $outlookExe = $executables[0]
                }
                else {
                    # For ClickToRun, there might be more than one Outlook.exe; downloaded Outlook.exe under the Office installation path.
                    # e.g. C:\Program Files\Microsoft Office\Updates\Download\PackageFiles\7ABA93E3-58C2-4BEE-AB49-3438C9F29D70\root\Office16\Outlook.exe
                    # Pick the one without 'PackageFiles' in the path.
                    $outlookExe = $executables | Where-Object {$_.FullName -notlike '*PackageFiles*'} | Select-Object -First 1
                }

                Write-Log "TTD launching Outlook"
                $ttd = Start-TTD -Path (Join-Path $tempPath 'TTD') -Executable $outlookExe.FullName -ErrorAction Stop
                Write-Host "Outlook has started (PID: $($ttd.TargetProcess.Id)) . It might take some time for Outlook to appear." -ForegroundColor Green
            }

            $ttdStarted = $true
        }

        if ($netshTraceStarted -or $outlookTraceStarted -or $psrStarted -or $ldapTraceStarted -or $capiTraceStarted -or $tcoTraceStarted -or $fiddlerCapStarted -or $crashDumpStarted -or $procmonStared -or $wamTraceStarted -or $wfpStarted -or $ttdStarted) {
            Write-Log "Waiting for the user to stop"
            Read-Host "Hit enter to stop"
        }
    }
    catch {
        # Log & save the exception so that I can analyze later. Then rethrow.
        Write-Log "Exception occured. $_"
        $_ | Export-CliXml (Join-Path $tempPath 'Exception.xml')
        throw
    }
    finally {
        Write-Progress -Activity 'Stopping traces' -Status "Please wait." -PercentComplete -1

        if ($ttdStarted) {
            $(Stop-TTD $ttd -KeepTargetProcess | Out-Null) 2>&1 | Write-Log

            # Outlook might be holding the TTD file.
            # Tell the user to stop Outlook and wait for the process to shutdown.
            if (-not $ttd.TargetProcess.HasExited) {
                Write-Log "Waiting for the user to shutdown Outlook."
                Write-Host "TTD Tracing is stopped. Please shutdown Outlook" -ForegroundColor Green
                Write-Progress -Activity 'Stopping traces' -Status "Please shutdown Outlook." -PercentComplete -1

                # Wait for Outlook to be stopped. Nudge the user once in a while.
                while ($true) {
                    $timeout = $(Wait-Process -InputObject $ttd.TargetProcess -Timeout 30 -ErrorAction Continue) 2>&1
                    if ($timeout) {
                        Write-Host "Please shutdown Outlook." -ForegroundColor Yellow
                    }
                    else {
                        Write-Host "Outlook is closed. Moving on." -ForegroundColor Green
                        break
                    }
                }
            }

            $ttd.TargetProcess.Dispose()
        }

        if ($netshTraceStarted) {
            Stop-NetshTrace
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

        if ($psrStarted) {
            Stop-PSR
        }

        Write-Progress -Activity 'Stopping traces' -Status "Please wait." -Completed

        # Save the event logs after tracing is done and wait for the tasks started earlier.
        if ($Component -contains 'Configuration') {
            Write-Progress -Activity 'Saving event logs.' -Status 'Please wait.' -PercentComplete -1
            $(Save-EventLog -Path $EventDir) 2>&1 | Write-Log
            Write-Progress -Activity 'Saving event logs.' -Status 'Please wait.' -Completed

            if ($oSConfigurationTask) {
                Write-Progress -Activity 'Saving OS configuration' -Status "Please wait." -PercentComplete -1
                $($oSConfigurationTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving OS configuration' -Status "Please wait." -Completed
                Write-Log "oSConfigurationTask is complete."
            }

            if ($officeRegistryTask) {
                Write-Progress -Activity 'Saving Office Registry' -Status "Please wait." -PercentComplete -1
                $($officeRegistryTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving Office Registry' -Status "Please wait." -Completed
                Write-Log "officeRegistryTask is complete."
            }

            if ($networkInfoTask) {
                Write-Progress -Activity 'Saving network info' -Status "Please wait." -PercentComplete -1
                $($networkInfoTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving network info' -Status "Please wait." -Completed
                Write-Log "networkInfoTask is complete."
            }

            if ($officeModuleInfoTask) {
                [TimeSpan]$timeout = [TimeSpan]::FromSeconds(30)
                Write-Progress -Activity 'Saving Office module info' -Status "Please wait up to $timeout" -PercentComplete -1

                if (Wait-Task $officeModuleInfoTask -Timeout $timeout)  {
                    Write-Log "officeModuleInfoTask is complete before timeout."
                }
                else {
                    Write-Log "officeModuleInfoTask timed out after $($timeout.TotalSeconds) seconds. Task will be canceled."
                    $cts.Cancel()
                }

                $($officeModuleInfoTask | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving Office module info' -Status 'Please wait.' -Completed
                Write-Log "officeRegistryTask is complete."
            }

            if ($msinfo32Task) {
                Write-Progress -Activity 'Saving MSInfo32' -Status 'Please wait.' -PercentComplete -1
                $($msinfo32Task | Receive-Task -AutoRemoveTask) 2>&1 | Write-Log
                Write-Progress -Activity 'Saving MSInfo32' -Status 'Please wait.' -Completed
            }

            # Save process list again after traces
            if ($Component.Count -gt 1) {
                Run-Command {param($OSDir) Save-Process -Path $OSDir} -ArgumentList $OSDir
                # $(Save-Process -Path $OSDir) 2>&1 | Write-Log
            }
        }

        Close-TaskRunspace
        Close-Log
    }

    $zipFileName = "Outlook_$($env:COMPUTERNAME)_$(Get-Date -Format "yyyyMMdd_HHmmss")"

    if ($SkipZip) {
        Rename-Item -Path $tempPath  -NewName $zipFileName
        return
    }

    Compress-Folder -Path $tempPath -ZipFileName $zipFileName -Destination $Path | Out-Null

    if (Test-Path $tempPath) {
        # Removing temp files might take a while. Do it in a background.
        $job = Start-Job -ScriptBlock {
            Remove-Item $using:tempPath -Recurse -Force
        }
        Write-Host "Temporary folder `"$tempPath`" will be removed by a background job (Job ID: $($job.Id))"
    }

    Write-Host "The collected data is `"$(Join-Path $Path $zipFileName).zip`"" -ForegroundColor Green
    Invoke-Item $Path
}

# Configure Export-Clixml & Out-File to use UTF8 by default.
if ($PSDefaultParameterValues -ne $null -and -not $PSDefaultParameterValues.Contains("Export-CliXml:Encoding")) {
    $PSDefaultParameterValues.Add("Export-Clixml:Encoding", 'UTF8')
}

if ($PSDefaultParameterValues -ne $null -and -not $PSDefaultParameterValues.Contains("Out-File:Encoding")) {
    $PSDefaultParameterValues.Add("Out-File:Encoding", 'utf8')
}

Export-ModuleMember -Function Start-WamTrace, Stop-WamTrace, Start-OutlookTrace, Stop-OutlookTrace, Start-NetshTrace, Stop-NetshTrace, Start-PSR, Stop-PSR, Save-EventLog, Get-MicrosoftUpdate, Save-MicrosoftUpdate, Get-InstalledUpdate,  Save-OfficeRegistry, Get-ProxySetting, Save-OSConfiguration, Get-ProxySetting, Get-NLMConnectivity, Get-WSCAntivirus, Save-CachedAutodiscover, Remove-CachedAutodiscover, Start-LdapTrace, Stop-LdapTrace, Save-OfficeModuleInfo, Start-SavingOfficeModuleInfo, Stop-SavingOfficeModuleInfo, Save-MSInfo32, Start-CAPITrace, Stop-CapiTrace, Start-FiddlerCap, Start-Procmon, Stop-Procmon, Start-TcoTrace, Stop-TcoTrace, Get-OfficeInfo, Add-WerDumpKey, Remove-WerDumpKey, Start-WfpTrace, Stop-WfpTrace, Save-Dump, Save-MSIPC, Get-EtwSession, Stop-EtwSession, Get-Token, Test-Autodiscover, Get-LogonUser, Get-JoinInformation, Get-OutlookProfile, Get-OutlookAddin, Get-ClickToRunConfiguration, Get-DeviceJoinStatus, Save-NetworkInfo, Start-TTD, Stop-TTD, Attach-TTD, Collect-OutlookInfo