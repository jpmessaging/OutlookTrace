[![en](https://img.shields.io/badge/English-英語-red)](https://github.com/jpmessaging/OutlookTrace/blob/master/Readme.md)

## 概要

OutlookTrace.psm1 は Outlook に関する情報採取用の PowerShell スクリプトです。

[ダウンロード](https://github.com/jpmessaging/OutlookTrace/releases/download/v2024-12-09/OutlookTrace.psm1)

SHA256: `1C971F65E45DAE1385A35E58ED781E0BEE228140E7A72C243CADD0BFC82FC340`

`Get-FileHash` コマンドでファイル ハッシュを取得できます:

  ```PowerShell
  Get-FileHash <.psm1 ファイルのパス> -Algorithm SHA256
  ```

Fiddler トレースや Process Monitor ログ、ZoomIt によるスクリーン レコーディング、そして TTD トレースも含めて採取する場合には以下から事前にダウンロードできます:

- [FiddlerCapSetup](https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe)
- [Process Monitor](https://download.sysinternals.com/files/ProcessMonitor.zip)
- [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip)
- [TTD](https://windbg.download.prss.microsoft.com/dbazure/prod/1-11-429-0/TTD.msixbundle)

いずれも `Collect-OutlookInfo` の `-Path` パラメータで指定するフォルダ配下に配置ください。  

## 利用方法

1. Outlook を実行している場合には終了します。
2. OutlookTrace.psm1 をダウンロードして対象のマシン上にコピーします。
3. 管理者権限で Windows PowerShell を起動します。
4. PowerShell で以下を実行して OutlookTrace.psm1 のブロックを解除します

    ```PowerShell
    Unblock-File <OutlookTrace.psm1 のパス>
    ```

    例:  
    ```PowerShell
    Unblock-File C:\temp\OutlookTrace.psm1
    ```

5. 一時的に ExecutionPolicy を `RemoteSigned` へ変更します

   ```PowerShell
   Set-ExecutionPolicy RemoteSigned -Scope Process
   ```

   確認を求められるので、`Y` を入力します。

6. OutlookTrace.psm1 をインポートします。

    ```PowerShell
    Import-Module <OutlookTrace.psm1 のパス> -DisableNameChecking
    ```

    例:

    ```PowerShell
    Import-Module C:\temp\OutlookTrace.psm1 -DisableNameChecking
    ```

    💡 もし上記が失敗する場合には、`Get-ExecutionPolicy -List` を実行して、その結果をお寄せください。

7. `Collect-OutlookInfo` を実行します。

    ※ 採取するコンポーネントについてはエンジニアからの案内をご確認ください。

    ※ Fiddler、Procmon、または ZoomIt によるスクリーン レコーディングを採取する場合、スクリプト内で自動的にダウンロードを試みます。インターネットへのアクセスに制限がある環境で実行する場合には、事前にダウンロードした [FiddlerCapSetup.exe](https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe)、[ProcessMonitor.zip](https://download.sysinternals.com/files/ProcessMonitor.zip)、そして [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip) を、下記 `Path` パラメータで指定するフォルダに配置ください。

    ```
    Collect-OutlookInfo -Path <出力先フォルダ> -Component <採取するコンポーネント>
    ```

    例:

    ```
    Collect-OutlookInfo -Path C:\temp -Component Configuration, Outlook, Netsh, PSR, WAM
    ```

8. 正常にトレースが開始されると、`Press enter to stop` と表示されます。

    ※ 採取するコンポーネントに `Dump` を含めた場合、`Press enter to save a process dump of Outlook. To quit, enter q:` とプロンプトされます。ダンプ ファイルを取得したいタイミングで Enter を入力してください。ハング事象の場合、およそ 30 秒間隔で 3 回程度採取ください。ダンプ ファイルの採取が終了したら `q` を入力します。

    ※ 採取するコンポーネントに `Fiddler` を含めた場合、[FiddlerCap Web Recorder] ダイアログボックスが表示されます。以下の手順に従って手動で、キャプチャを開始ください。キャプチャ開始後に事象を再現します。

    ⚠️ スクリプト実行ユーザーと、情報採取対象ユーザーが異なる場合には FiddlerCap は自動的に開始されません。情報採取対象ユーザーが FiddlerCap.exe を開始する必要があります。

    <details>
        <summary>Fiddler 開始方法</summary>

    1. [HTTPS 通信を解読] にチェックを入れます。
    2. 以下の説明が表示されたら、内容を確認して [OK] をクリックします。

       ```
       HTTPS の解読は、HTTPS プロトコル経由で送られる Raw トラフィックを見るためにデバッグしやすくしてくれます。
       この機能は SSL トラフィックを解読し、ローカルに生成された証明書を用いて再度暗号化します。よって、この機能を使うと、不明な発行元からの証明書を使っているリモートサイトであること表示する、赤い警告ページが Web ブラウザーに表示されることを意味します。
       このトラフィックをキャプチャすることに限定して、このブラウザーに表示される警告を無視してください。
       ```

    3. 以下の内容のセキュリティ警告が表示されたら、[はい] をクリックします。

       ```
       発行者が次であると主張する証明機関 (CA) から証明書をインストールしようとしています:

       DO_NOT_TRUST_FiddlerRoot

       証明書が実際に "DO_NOT_TRUST_FiddlerRoot" からのものであるかどうかを検証できません。"DO_NOT_TRUST_FiddlerRoot" に連絡して発行者を確認する必要があります。 次の番号はこの過程で役立ちます:

       拇印 (sha1): ***

       警告:
       このルート証明書をインストールすると、この CA によって発行された証明書は自動的に信頼されます。確認されていない拇印付きの証明書をインストールすることは、セキュリティ上、危険です。 [はい] をクリックすると、この危険を認識したことになります。

       この証明書をインストールしますか?
       ```

    4. [1. キャプチャ開始] をクリックします。

        自動的にブラウザが起動されたら、そのブラウザはクローズいただいて結構です。
    </details>

9.  Outlook を起動して、現象を再現させます。
10. 採取するコンポーネントに Fiddler を含めた場合、以下の手順で停止して保存します。

    <details>
        <summary>Fiddler 停止方法</summary>
        
    1. [2. キャプチャ停止] をクリックします。
    2. [3. キャプチャ保存] をクリックします。
    3. [ファイルの種類] で `Password-Protected Capture (*.saz)` を選択します。
    4. ファイルを Collect-OutlookInfo の "Path" パラメータに指定したフォルダ配下に作成された GUID 名のフォルダに保存します。
    5. [FiddlerCap Web Recorder] ダイアログボックスをクローズします。
        この時以下の内容が表示されたら、[はい] をクリックします。

        ```
        次の証明書をルート ストアから削除しますか?

        サブジェクト: DO_NOT_TRUST_FiddlerRoot, DO_NOT_TRUST, Created by http://www.fiddler2.com
        発行者: 自己発行
        有効期間: ***
        シリアル番号 : ***
        拇印 (sha1): ***
        拇印 (md5):
        ```

    </details>

11. コンソールに Enter キーを入力しトレースを停止します。

手順 6 で出力先に指定したフォルダに作成された `"Outlook_<マシン名>_<取得日時>.zip"` という名前の ZIP ファイルをお寄せください。  
Fiddler トレースを採取した場合には、手順 9 で指定したパスワードも併せてお寄せください。

## ライセンス

Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
