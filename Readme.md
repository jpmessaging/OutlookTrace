## Overview
OutlookTrace.psm1 is a PowerShell script to collect several traces related to Microsoft Outlook

[Download](https://github.com/jpmessaging/OutlookTrace/releases/download/v2020-09-24/OutlookTrace.psm1)

## How to use
1. Download OutlookTrace.psm1 and place it on the target machine.
2. Start cmd as administrator.
3. Start PowerShell as follow.

    ```PowerShell
    powershell -ExecutionPolicy Bypass
    ```

4. Import OutlookTrace.psm1

    ```
    Import-Module <path to OutlookTrace.psm1> -DisableNameChecking
    ```

    e.g.
    ```
    Import-Module C:\temp\OutlookTrace.psm1 -DisableNameChecking
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

6. When traces have started successfully, it shows "Hit enter to stop".

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

7. Reproduce the issue.
8. When "Fiddler" is included, stop and save the capture.

    1. Click [2. Stop Capture].
    2. Click [3. Save Capture].
    3. In [Save as type], select `Password-Protected Capture (*.saz)`.
    4. Save the capture in the folder with GUID name created under "Path" parameter you specified in Collect-OutloookInfo.
    5. Close the [FiddlerCap Web Recorder] dialog box.

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

9. Hit enter key in the console to stop.

Send the zip file `"Outlook_<MachineName>_<DateTime>.zip"` in the output folder specified in step 5.  
If you captured a Fiddler trace, send the password used in step 8 too.

## License
Copyright (c) 2020 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

