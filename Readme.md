[![ja](https://img.shields.io/badge/Japanese-日本語-green)](https://github.com/jpmessaging/OutlookTrace/blob/master/Readme_ja.md)

## Overview

OutlookTrace.psm1 is a PowerShell script to collect several traces related to Microsoft Outlook

[Download](https://github.com/jpmessaging/OutlookTrace/releases/download/v2025-05-06/OutlookTrace.psm1)

SHA256: `C98D22FECE3D870D1B01C73E6FFD4618117D45A2460830E346D43F5E61350EA3`

You can get the file hash with `Get-FileHash`:

  ```PowerShell
  Get-FileHash <Path to .psm1 file> -Algorithm SHA256
  ```

## How to use

1.  Shutdown Outlook if it's running.
2.  Download OutlookTrace.psm1 and place it on the target machine.
3.  Start PowerShell as administrator.
4.  Run the following command to unblock the file.

    ```PowerShell
    Unblock-File <Path to OutlookTrace.psm1>
    ```

    e.g.

    ```PowerShell
    Unblock-File C:\temp\OutlookTrace.psm1
    ```

5. Temporarily set ExecutionPolicy to `RemoteSigned`.

   ```PowerShell
   Set-ExecutionPolicy RemoteSigned -Scope Process
   ```

    Press `Y` when asked for a confirmation.

6.  Import OutlookTrace.psm1.

    ```PowerShell
    Import-Module <Path to OutlookTrace.psm1> -DisableNameChecking
    ```

    e.g.

    ```PowerShell
    Import-Module C:\temp\OutlookTrace.psm1 -DisableNameChecking
    ```

7.  Run `Collect-OutlookInfo`.

    Note: Follow Microsoft engineer's instruction regarding which components to trace.

    ```
    Collect-OutlookInfo -Path <Output folder> -Component <components to trace>
    ```

    e.g.

    ```
    Collect-OutlookInfo -Path C:\temp -Component Configuration, Outlook, Netsh, PSR, WAM
    ```

8.  When traces have started successfully, it shows `Press enter to stop`.

    Note: When `Dump` is included in Component parameter, you are prompted with `Press enter to save a process dump of Outlook. To quit, enter q:`. Press enter key to save a dump file. For a hang issue, repeat the process to collect 3 dump files, with interval of about 30 seconds between saves. When finished, press `q`.

    Note: When `Fiddler` is included in Component parameter, an application called "Fiddler Everywhere Reporter" starts. Use the following instructions to start capture, and then reproduce the issue

    ⚠️ When the target user is different from the one running the script, Fiddler Everywhere Reporter does not start. The target user needs to start it manually.

    <details>
        <summary>How to start Fiddler capture</summary>
        
    1. Check [I agree to the Terms of Service and Privacy Policy] and click [Proceed].
    2. Select [Start Capturing Everything] in the box 1. on the top.
    3. If a dialog [Trust Certificate and Enable HTTPS] appears, click [Trust and Enable HTTPS].
    4. Click [Yes] in the following security warning.

        ```
        You are about to install a certificate from a certification authority (CA) claiming to represent:

        Fiddler Root Certificate Authority

        Windows cannot validate that the certificate is actually from "Fiddler Root Certificate Authority". You should confirm its origin by contacting "Fiddler Root Certificate Authority". The following number will assist you in this process:

        Thumbprint (sha1): ***

        Warning:
        If you install this root certificate, Windows will automatically trust any certificate issued by this CA. Installing a certificate with an unconfirmed thumbprint is a security risk. If you click "Yes" you acknowledge this risk.

        Do you want to install this certificate?
        ```
    </details>

9.  Start Outlook and reproduce the issue.
10. When `Fiddler` is included, stop and save the capture.

    <details>
    <summary>How to stop Fiddler capture</summary>

    1. Click [2. Stop Capture].
    2. Click [3. Save Capture].
    3. Save the capture in the folder with GUID name created under "Path" parameter you specified in Collect-OutloookInfo.  
       ⚠️ Password must be at least 8 characters.
    4. Click a menu item [Certificate]-[Remove Root Certificate]  

        In the following dialog, click [Yes].

        ```
        Do you want to DELETE the following certificate from the Root Store?

        Subject : Fiddler Root Certificate Authority, Progress Telerik Fiddler, Created by http://www.fiddler2.com
        Issuer : Self Issued
        Time Validity : ***
        Serial Number : ***
        Thumbprint (sha1) : ***
        Thumbprint (md5) : ***
        ```

    5. Close the Fiddler Everywhere Reporter
    </details>

11. Press enter key in the console to stop.

Send the zip file `"Outlook_<MachineName>_<DateTime>.zip"` in the output folder specified in step 6.  
If you captured a Fiddler trace, send the password used in step 9 too.

## Parameters

### Mandatory parameters

| Name      | Description                                                                             |
| --------- | --------------------------------------------------------------------------------------- |
| Path      | Folder path where gathered data will be placed. It will be created if it does not exist |
| Component | Diagnostics data to collect (see below)                                                 |

### Possible values for `Component` parameter

| Name          | Description                                                                                                        |
| ------------- | ------------------------------------------------------------------------------------------------------------------ |
| Configuration | OS config, Registry, Event logs, Proxy settings, etc.                                                              |
| Outlook       | Outlook ETW                                                                                                        |
| Netsh         | Netsh ETW                                                                                                          |
| PSR           | Problem Steps Recorder                                                                                             |
| WAM           | WAM (Web Account Manager) ETW                                                                                      |
| Fiddler       | [Fiddler](https://api.getfiddler.com/reporter/win/latest) trace (Fiddler trace must be manually started & stopped) |
| Procmon       | [Process monitor](https://docs.microsoft.com/en-us/sysinternals/downloads/procmon)                                 |
| LDAP          | LDAP ETW                                                                                                           |
| CAPI          | CAPI (Crypt API) ETW                                                                                               |
| TCO           | TCO trace                                                                                                          |
| Dump          | Outlook's process dump                                                                                             |
| CrashDump     | Crash dump for any process (see `CrashDumpTargets` below)                                                          |
| HungDump      | Outlook's hung dump (When a window hung is detected, a dump file is generated)                                     |
| WPR           | WPR (Windows Performance Recorder) ETW (OS must be Windows 10 or above)                                            |
| WFP           | Windows Firewall diagnostic log                                                                                    |
| Performance   | Performance counter log (Process, Memory, LogicalDisk etc.)                                                        |
| TTD           | Time Travel Debugging trace (OS must be Windows 10 or above)                                                       |
| Recording     | Screen recording by [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip)                                   |
| NewOutlook    | New Outlook for Windows logs                                                                                       |
| WebView2      | WebView2 NetLog                                                                                                    |

>[!IMPORTANT]
>`Collect-OutlookInfo` tries to download Fiddler Everywhere Reporter, Procmon, TTD, and ZoomIt when `Component` parameter includes `Fiddler`, `Procmon`, `TTD` and `Recording` respectively.  
> If the target machine does not have access to the Internet, please download from the links below and place them in the folder specified by `Path` parameter:
> 
> - [Fiddler Everywhere Reporter](https://api.getfiddler.com/reporter/win/latest)
> - [Procmon](https://download.sysinternals.com/files/ProcessMonitor.zip)
> - [TTD](https://windbg.download.prss.microsoft.com/dbazure/prod/1-11-481-0/TTD.msixbundle)
> - [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip)

### Optional parameters

| Name                 | Description                                                                                                                                    |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| User                 | Target user whose configuration data is collected. By default, it's the logon user (Note: Not necessarily the current user running the script) |
| LogFileMode          | ETW trace's mode. Valid values: `NewFile`, `Circular` (Default: `NewFile`)                                                                     |
| MaxFileSizeMB        | Max file size for ETW trace files. By default, 256 MB when `NewFile` and 2048 MB when `Circular`                                               |
| NetshReportMode      | Netsh trace's report mode. Valid values: `None`, `Mini`, `Full` (Default: `None`)                                                              |
| ArchiveType          | Valid values: `Zip` or `Cab`. Zip is faster, but Cab is smaller (Default: `Zip`)                                                               |
| SkipArchive          | Switch to skip archiving (zip or cab)                                                                                                          |
| SkipAutoUpdate       | Switch to skip auto update                                                                                                                     |
| AutoFlush            | Switch to flush log data every time it's written (This is just for troubleshooting the script)                                                 |
| PsrRecycleInterval   | PSR recycle interval. A new instance of PSR is created after this interval (Default: `00:10:00`, Min: `00:01:00`, Max: `01:00:00`)             |
| HungTimeout          | TimeSpan used to detect a hung window when `HungDump` is in `Component` parameter (Default: `00:00:05`, Min: `00:00:01`, Max: `00:01:00`)      |
| MaxHungDumpCount     | Max number of hung dump files to be saved per process instance (Default: `3`, Min: `1`, Max: `10`)                                             |
| TargetProcessName    | Target process name (such as `Outlook` or `olk`). By default `Outlook`. `olk` when `NewOutlook` is in `Component` parameter                    |
| CrashDumpTargets     | Names of the target processes for crash dumps. When not specified, all processes are the targets                                               |
| RemoveIdentityCache  | Switch to remove identity cache                                                                                                                |
| EnablePageHeap       | Switch to enable full page heap for Outlook.exe (With page heap, Outlook will consume a lot of memory and slow down)                           |
| EnableLoopbackExempt | Switch to add Microsoft.AAD.BrokerPlugin to Loopback Exempt                                                                                    |
| SkipVersionCheck     | Switch to skip script version check                                                                                                            |
| TTDCommandlineFilter | Command line filter for TTD monitor                                                                                                            |
| TTDModules           | Restrict TTD trace to specified modules                                                                                                        |
| TTDShowUI            | Switch to show TTD UI                                                                                                                          |
| WprProfiles          | WPR profiles to capture (Default: `GeneralProfile`, `CPU`, `DiskIO`, `FileIO`, `Registry`, `Network`)                                          |

## License

Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
