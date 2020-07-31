## Overview
OutlookTrace.ps1 is a PowerShell script to collect several traces related to Microsoft Outlook

[Download](https://github.com/jpmessaging/OutlookTrace/releases/download/v2020-07-31/OutlookTrace.ps1)

## How to use
1. Download OutlookTrace.ps1 and unblock the file.

    1. Right-click the ps1 file and click [Property].
    2. In the [General] tab, if you see "This file came from another computer and might be blocked to help protect this computer]", check [Unblock].

2. Place OutlookTrace.ps1 on the target machine.
3. Start PowerShell as administrator.

    Run Get-ExecutionPolicy and if it's not RemoteSigned, set it RemoteSigned as follows.

    ```PowerShell
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    ```
       
4. Dot source OutlookTrace.ps1

    ```
    . <path to OutlookTrace.ps1>
    ```

    e.g.
    ```
    . C:\temp\OutlookTrace.ps1
    ```

5. Run Collect-OutlookInfo

    Note: Follow Microsoft engineer's instruction regarding which components to trace.

    ```
    Collect-OutlookInfo -Path <output folder> -Component <components to trace>
    ```

    例:
    ```
    Collect-OutlookInfo -Path c:\temp -Component Configuration, Netsh, Outlook
    ```

6. When traces have started successfully, it shows "Hit enter to stop". Reproduce the issue.
   
    Note: When "Fiddler" is included in Component parameter, a dialog box [FiddlerCap Web Recorder] appears. Use the following instructions to start capture, and then reproduce the issue.

    1. Check [Decrypt HTTPS traffic] 
    2. When the following explanation appears, read it and click [OK].

        ```
        HTTPS decryption will enable your debugging buddy to see the raw traffic sent via the HTTPS protocol. 

        This feature works by decrypting SSL traffic and reencrypting it using a locally generated certificate. FiddlerCap will generate this certificate and remove it when you close this tool.
        You may choose to temporarily install this certificate in the Trusted store to avoid warnings from your browser or client application.
        ```

    3. Click [Yes] on the following security warning.

        ```
        You are about to install a certificate from a certification authority (CA) claiming to represent:

        DO_NOT_TRUST_FiddlerRoot

        Windows cannot validate that the certificate is actually from "DO_NOT_TRUST_FiddlerRoot". You should confirm its origin by contacting "DO_NOT_TRUST_FiddlerRoot". The following number will assist you in this process:

        Thumbprint (sha1): ***

        Warning:
        If you install this root certificate, Windows will automatically trust any certificate issued by this CA. Installing a certificate with an unconfirmed thumbprint is a security risk. If you click "Yes" you acknowledge this risk.

        Do you want to install this certificate?
        ```

    4. Click [1. Start capture].

        If a web browser starts automatically, you can close the browser.

7. After reproducing the issue, hit Enter in the console to stop traces.
8. When "Fiddler" is included, stop and save the capture.

    1. Click [2. Stop Capture].
    2. Click [3. Save Capture].
    3. Save the capture.
    4. Close the [FiddlerCap Web Recorder] dialog box.

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

9. If ExecutionPolicy is changed in the step 3, set it back to original value

    ```PowerShell
    Set-ExecutionPolicy -ExecutionPolicy <original value>
    ```
    
Send the following files:

- A zip file `"Outlook_<MachineName>_<DateTime>.zip"` in the output folder specified in step 5.
- If Fiddler is included, capture file (`"FiddlerCap_***.saz"`) saved in step 8.


## License
Copyright (c) 2020 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

