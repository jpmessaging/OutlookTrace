[![en](https://img.shields.io/badge/English-英語-red)](https://github.com/jpmessaging/OutlookTrace/blob/master/Readme.md)

## 概要

OutlookTrace.psm1 は Outlook に関する情報採取用の PowerShell スクリプトです。

[ダウンロード](https://github.com/jpmessaging/OutlookTrace/releases/download/v2025-10-07/OutlookTrace.psm1)

SHA256: `CE2C51D98041403BB22C13277122F0B475306FED4E90517BF316E76FBE1FD439`

`Get-FileHash` コマンドでファイル ハッシュを取得できます:

  ```PowerShell
  Get-FileHash <.psm1 ファイルのパス> -Algorithm SHA256
  ```

Fiddler トレースや Process Monitor ログ、ZoomIt によるスクリーン レコーディング、そして TTD トレースも含めて採取する場合には以下から事前にダウンロードできます:

- [Fiddler Everywhere Reporter](https://api.getfiddler.com/reporter/win/latest)
- [Process Monitor](https://download.sysinternals.com/files/ProcessMonitor.zip)
- [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip)
- [TTD](https://windbg.download.prss.microsoft.com/dbazure/prod/1-11-532-0/TTD.msixbundle)

いずれも `Collect-OutlookInfo` の `-Path` パラメータで指定するフォルダ配下に配置ください。  

## 利用方法

1. Outlook を実行している場合には終了します。
2. OutlookTrace.psm1 をダウンロードして対象のマシン上にコピーします。
3. 管理者権限で Windows PowerShell を起動します ([管理者として実行] で開始します)。
4. PowerShell で以下を実行して OutlookTrace.psm1 のブロックを解除します。

    ```PowerShell
    Unblock-File <OutlookTrace.psm1 のパス>
    ```

    例:  
    ```PowerShell
    Unblock-File C:\temp\OutlookTrace.psm1
    ```

5. 一時的に ExecutionPolicy を `RemoteSigned` へ変更します。

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

    ※ Fiddler、Procmon、または ZoomIt によるスクリーン レコーディングを採取する場合、スクリプト内で自動的にダウンロードを試みます。インターネットへのアクセスに制限がある環境で実行する場合には、事前にダウンロードした [Fiddler Everywhere Reporter](https://api.getfiddler.com/reporter/win/latest)、[ProcessMonitor.zip](https://download.sysinternals.com/files/ProcessMonitor.zip)、そして [ZoomIt](https://download.sysinternals.com/files/ZoomIt.zip) を、下記 `Path` パラメータで指定するフォルダに配置ください。

    ```
    Collect-OutlookInfo -Path <出力先フォルダ> -Component <採取するコンポーネント>
    ```

    例:

    ```
    Collect-OutlookInfo -Path C:\temp -Component Configuration, Outlook, Netsh, PSR, WAM
    ```

8. 正常にトレースが開始されると、`Press enter to stop` と表示されます。

    ※ 採取するコンポーネントに `Dump` を含めた場合、`Press enter to save a process dump of Outlook. To quit, enter q:` とプロンプトされます。ダンプ ファイルを取得したいタイミングで Enter を入力してください。ハング事象の場合、およそ 30 秒間隔で 3 回程度採取ください。ダンプ ファイルの採取が終了したら `q` を入力します。

    ※ 採取するコンポーネントに `Fiddler` を含めた場合、Fiddler Everywhere Reporter というアプリケーションが開始されます。以下の手順に従って手動で、キャプチャを開始ください。キャプチャ開始後に事象を再現します。

    ⚠️ スクリプト実行ユーザーと、情報採取対象ユーザーが異なる場合には Fiddler Everywhere Reporter は自動的に開始されません。情報採取対象ユーザーが Fiddler Everywhere Reporter-***.exe を開始する必要があります。

    <details>
        <summary>Fiddler 開始方法</summary>

    1. [I agree to the Terms of Service and Privacy Policy] にチェックを入れ、[Proceed] をクリックします
    2. 上部のボックス 1 で、[Start Capturing Everything] を選択します。
    3. [Trust Certificate and Enable HTTPS] というダイアログが表示された場合は、[Trust and Enable HTTPS] をクリックします
    4. 以下の内容のセキュリティ警告が表示されたら、[はい] をクリックします。

       ```
       発行者が次であると主張する証明機関 (CA) から証明書をインストールしようとしています:

       Fiddler Root Certificate Authority

       証明書が実際に "Fiddler Root Certificate Authority" からのものであるかどうかを検証できません。"Fiddler Root Certificate Authority" に連絡して発行者を確認する必要があります。 次の番号はこの過程で役立ちます:

       拇印 (sha1): ***

       警告:
       このルート証明書をインストールすると、この CA によって発行された証明書は自動的に信頼されます。確認されていない拇印付きの証明書をインストールすることは、セキュリティ上、危険です。 [はい] をクリックすると、この危険を認識したことになります。

       この証明書をインストールしますか?
       ```
    </details>

9.  Outlook を起動して、現象を再現させます。
10. 採取するコンポーネントに Fiddler を含めた場合、以下の手順で停止して保存します。

    <details>
        <summary>Fiddler 停止方法</summary>
        
    1. [2. Stop Capture] をクリックします。
    2. [3. Save Capture] をクリックします。
    3. ファイルを Collect-OutlookInfo の "Path" パラメータに指定したフォルダ配下に作成された GUID 名のフォルダに保存します。

       ⚠️ パスワードの長さは 8 文字以上にする必要があります。

    4. メニュー アイテムの  [Certificate]-[Remove Root Certificate] をクリックします。

        この時以下の内容が表示されたら、[はい] をクリックします。

        ```
        次の証明書をルート ストアから削除しますか?

        サブジェクト: Fiddler Root Certificate Authority, Progress Telerik Fiddler, Created by http://www.fiddler2.com
        発行者: 自己発行
        有効期間: ***
        シリアル番号 : ***
        拇印 (sha1): ***
        拇印 (md5):
        ```
    5. Fiddler Everywhere Reporter を終了します。
    </details>

11. コンソールに Enter キーを入力しトレースを停止します。

手順 7 で出力先に指定したフォルダに作成された `"Outlook_<マシン名>_<取得日時>.zip"` という名前の ZIP ファイルをお寄せください。  
Fiddler トレースを採取した場合には、手順 10 で指定したパスワードも併せてお寄せください。

## ライセンス

Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
