## Overview

OutlookTrace.psm1 is a PowerShell script to collect several traces related to Microsoft Outlook

[Download](https://github.com/jpmessaging/OutlookTrace/releases/download/v2023-11-29/OutlookTrace.psm1)

## How to use

1.  Shutdown Outlook if it's running.
2.  Download OutlookTrace.psm1 and place it on the target machine.
3.  Start cmd as administrator.
4.  Start PowerShell as follow.

    ```PowerShell
    powershell -ExecutionPolicy Bypass
    ```

5.  Import OutlookTrace.psm1

    ```
    Import-Module <path to OutlookTrace.psm1> -DisableNameChecking
    ```

    e.g.

    ```
    Import-Module C:\temp\OutlookTrace.psm1 -DisableNameChecking
    ```

6.  Run `Collect-OutlookInfo`

    Note: Follow Microsoft engineer's instruction regarding which components to trace.

    ```
    Collect-OutlookInfo -Path <output folder> -Component <components to trace>
    ```

    e.g.

    ```
    Collect-OutlookInfo -Path C:\temp -Component Configuration, Outlook, Netsh, PSR, WAM
    ```

7.  When traces have started successfully, it shows "Hit enter to stop".

    Note: When `Dump` is included in Component parameter, you are prompted with `Hit enter to save a process dump of Outlook. To quit, enter q:`. Hit enter key to save a dump file. For a hang issue, repeat the process to collect 3 dump files, with interval of about 30 seconds between saves. When finished, hit `q`.

    Note: When `Fiddler` is included in Component parameter, a dialog box [FiddlerCap Web Recorder] appears. Use the following instructions to start capture, and then reproduce the issue

    ⚠️ When the target user is different from the one running the script, FiddlerCap does not start. The target user needs to start FiddlerCap.exe manually.

    <details>
        <summary>How to start Fiddler capture</summary>
        
    1. Check [Decrypt HTTPS traffic]
    2. When the following explanation appears, read it and click [OK].

        ```
        HTTPS decryption will enable your debugging buddy to see the raw traffic sent via the HTTPS protocol.

        This feature works by decrypting SSL traffic and reencrypting it using a locally generated certificate. FiddlerCap will generate this certificate and remove it when you close this tool.
        You may choose to temporarily install this certificate in the Trusted store to avoid warnings from your browser or client application.
        ```

    3.  Click [Yes] on the following security warning.

        ```
        You are about to install a certificate from a certification authority (CA) claiming to represent:

        DO_NOT_TRUST_FiddlerRoot

        Windows cannot validate that the certificate is actually from "DO_NOT_TRUST_FiddlerRoot". You should confirm its origin by contacting "DO_NOT_TRUST_FiddlerRoot". The following number will assist you in this process:

        Thumbprint (sha1): ***

        Warning:
        If you install this root certificate, Windows will automatically trust any certificate issued by this CA. Installing a certificate with an unconfirmed thumbprint is a security risk. If you click "Yes" you acknowledge this risk.

        Do you want to install this certificate?
        ```

    4.  Click [1. Start capture].

        If a web browser starts automatically, you can close the browser.

    </details>

8.  Start Outlook and reproduce the issue.
9.  When "Fiddler" is included, stop and save the capture.

    <details>
    <summary>How to stop Fiddler capture</summary>

    1.  Click [2. Stop Capture].
    2.  Click [3. Save Capture].
    3.  In [Save as type], select `Password-Protected Capture (*.saz)`.
    4.  Save the capture in the folder with GUID name created under "Path" parameter you specified in Collect-OutloookInfo.
    5.  Close the [FiddlerCap Web Recorder] dialog box.

        If the following dialog appears, click [Yes].

        ```
        Do you want to DELETE the following certificate from the Root Store?

        Subject : DO_NOT_TRUST_FiddlerRoot, DO_NOT_TRUST, Created by http://www.fiddler2.com
        Issuer : Self Issued
        Time Validity : ***
        Serial Number : ***
        Thumbprint (sha1) : ***
        Thumbprint (md5) : ***
        ```

    </details>

10. Hit enter key in the console to stop.

Send the zip file `"Outlook_<MachineName>_<DateTime>.zip"` in the output folder specified in step 6.  
If you captured a Fiddler trace, send the password used in step 9 too.

## Parameters

### Mandatory parameters

| Name      | Description                                                                             |
| --------- | --------------------------------------------------------------------------------------- |
| Path      | Folder path where gathered data will be placed. It will be created if it does not exist |
| Component | Diagnostics data to collect (see below)                                                 |

### Possible values for `Component` parameter

| Name          | Description                                                                                                    |
| ------------- | -------------------------------------------------------------------------------------------------------------- |
| Configuration | OS config, Registry, Event logs, Proxy settings, etc.                                                          |
| Outlook       | Outlook ETW                                                                                                    |
| Netsh         | Netsh ETW                                                                                                      |
| PSR           | Problem Steps Recorder                                                                                         |
| WAM           | WAM (Web Account Manager) ETW                                                                                  |
| Fiddler       | [Fiddler](https://www.telerik.com/fiddler/fiddlercap) trace (Fiddler trace must be manually started & stopped) |
| Procmon       | [Process monitor](https://docs.microsoft.com/en-us/sysinternals/downloads/procmon)                             |
| LDAP          | LDAP ETW                                                                                                       |
| CAPI          | CAPI (Crypt API) ETW                                                                                           |
| TCO           | TCO trace                                                                                                      |
| Dump          | Outlook's process dump                                                                                         |
| CrashDump     | Crash dump for any process (see `CrashDumpTargets` above)                                                      |
| HungDump      | Outlook's hung dump (When a window hung is detected, a dump file is generated)                                 |
| WPR           | WPR (Windows Performance Recorder) ETW (OS must be Windows 10 or above)                                        |
| WFP           | Windows Firewall diagnostic log                                                                                |
| Performance   | Performance counter log (Process, Memory, LogicalDisk etc.)                                                    |
| TTD           | Time Travel Debugging trace (OS must be Windows 10 or above)                                                   |
| Recording     | Screen recording by [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip)                               |

💡 `Collect-OutlookInfo` tries to download FiddlerCap, Procmon, and ZoomIt when `Component` parameter includes `Fiddler`, `Procmon`, and `Recording` respectively. If Internet access is not available, please download [FiddlerCap](https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe), [Procmon](https://download.sysinternals.com/files/ProcessMonitor.zip), [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip) and place them in `Path` folder before running the command.

### Optional parameters

| Name                 | Description                                                                                                                                    |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| NetshReportMode      | Netsh trace's report mode. Valid values: `None`, `Mini`, `Full` (Default: `None`)                                                              |
| LogFileMode          | ETW trace's mode. Valid values: `NewFile`, `Circular` (Default: `NewFile`)                                                                     |
| MaxFileSizeMB        | Max file size for ETW trace files. By default, 256 MB when `NewFile` and 2048 MB when `Circular`                                               |
| ArchiveType          | Valid values: `Zip` or `Cab`. Zip is faster, but Cab is smaller (Default: `Zip`)                                                               |
| SkipArchive          | Switch to skip archiving (zip or cab)                                                                                                          |
| SkipAutoUpdate       | Switch to skip auto update                                                                                                                     |
| AutoFlush            | Switch to flush log data every time it's written (This is just for troubleshooting the script)                                                 |
| PsrRecycleInterval   | PSR recycle interval. A new instance of PSR is created after this interval (Default: `00:10:00`, Min: `00:01:00`, Max: `01:00:00`)             |
| User                 | Target user whose configuration data is collected. By default, it's the logon user (Note: Not necessarily the current user running the script) |
| HungTimeout          | Timespan used to detect a hung window when `HungDump` is requested in Component (Default: `00:00:05`, Min: `00:00:01`, Max: `00:01:00`)        |
| HungMonitorTarget    | Name of the target process to monitor a hung window (Default: `Outlook`)                                                                       |
| MaxHungDumpCount     | Max number of hung dump files to be saved per process instance (Default: `3`, Min: `1`, Max: `10`)                                             |
| CrashDumpTargets     | Names of the target processes for crash dumps. When not specified, all processes are the targets                                               |
| RemoveIdentityCache  | Switch to remove identity cache                                                                                                                |
| EnablePageHeap       | Switch to enable full page heap for Outlook.exe (With page heap, Outlook will consume a lot of memory and slow down)                           |
| EnableLoopbackExempt | Switch to add Microsoft.AAD.BrokerPlugin to Loopback Exempt                                                                                    |
| SkipVersionCheck     | Switch to skip script version check                                                                                                            |
| TTDCommandlineFilter | Command line filter for TTD monitor                                                                                                            |
| TTDShowUI            | Switch to show TTD UI                                                                                                                          |

## License

Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
