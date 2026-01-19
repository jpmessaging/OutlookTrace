[![en](https://img.shields.io/badge/English-英語-red)](https://github.com/jpmessaging/OutlookTrace/blob/master/Readme.md)

## 概要

OutlookTrace.psm1 は Outlook に関する情報採取用の PowerShell スクリプトです。

[ダウンロード](https://github.com/jpmessaging/OutlookTrace/releases/download/v2026-01-19/OutlookTrace.psm1)

SHA256: `77DA505A88246EA2AB4CF4024134F6DCDCF461EE8799904E500A5EC9AB7D0D0C`

`Get-FileHash` コマンドでファイル ハッシュを取得できます:

  ```PowerShell
  Get-FileHash <.psm1 ファイルのパス> -Algorithm SHA256
  ```

## 利用方法

1. ログの出力先フォルダーを作成します。以下の説明では C:\temp を出力先フォルダーの例として使用します。
2. 以下のリンクをクリックし、Fiddler Everywhere Reporter をダウンロードします。

    [Fiddler Everywhere Reporter](https://api.getfiddler.com/reporter/win/latest)

3. ダウンロードしたファイルをログの出力先フォルダーにコピーします。
4. Outlook を実行している場合には終了します。
5. OutlookTrace.psm1 をダウンロードして対象のマシン上にコピーします。
6. 管理者権限で Windows PowerShell を起動します ([管理者として実行] で開始します)。
7. PowerShell で以下を実行して OutlookTrace.psm1 のブロックを解除します。

    ```PowerShell
    Unblock-File <OutlookTrace.psm1 のパス>
    ```

    例:  
    ```PowerShell
    Unblock-File C:\temp\OutlookTrace.psm1
    ```

8. 一時的に ExecutionPolicy を `RemoteSigned` へ変更します。

   ```PowerShell
   Set-ExecutionPolicy RemoteSigned -Scope Process
   ```

   確認を求められるので、`Y` を入力します。

9. OutlookTrace.psm1 をインポートします。

    ```PowerShell
    Import-Module <OutlookTrace.psm1 へのパス> -DisableNameChecking
    ```

    例:

    ```PowerShell
    Import-Module C:\temp\OutlookTrace.psm1 -DisableNameChecking
    ```

    💡 もし上記が失敗する場合には、`Get-ExecutionPolicy -List` を実行して、その結果をお寄せください。

10. `Collect-OutlookInfo` を実行します。

    ※ 採取するコンポーネントについてはエンジニアからの案内をご確認ください。

    ```
    Collect-OutlookInfo -Path <出力先フォルダ> -Component <採取するコンポーネント>
    ```

    例:

    ```
    Collect-OutlookInfo -Path C:\temp -Component Configuration, Outlook, Fiddler, Netsh, PSR, WAM
    ```

    正常にトレースが開始されると、`Press enter to stop` と表示されます。

11. Fiddler Everywhere Reporter が表示されない場合、**Outlook を開始するユーザー**にて出力先フォルダーにある Fiddler Everywhere Reporter-<バージョン>.exe を実行します。

    ※ スクリプトの実行ユーザーと、情報採取のターゲットとなるユーザーが異なる場合、ターゲット ユーザー自身が Fiddler Everywhere Reporter を実行する必要があります。

12. [I agree to the Terms of Service and Privacy Policy] にチェックを入れ、[Proceed] をクリックします。
13. 上部のボックス 1 で、[Start Capturing Everything] を選択します。
14. [Trust Certificate and Enable HTTPS] というダイアログが表示された場合は、[Trust and Enable HTTPS] をクリックします。
15. 以下の内容のセキュリティ警告が表示されたら、[はい] をクリックします。

    ```
    発行者が次であると主張する証明機関 (CA) から証明書をインストールしようとしています:

    Fiddler Root Certificate Authority

    証明書が実際に "Fiddler Root Certificate Authority" からのものであるかどうかを検証できません。"Fiddler Root Certificate Authority" に連絡して発行者を確認する必要があります。 次の番号はこの過程で役立ちます:

    拇印 (sha1): ***

    警告:
    このルート証明書をインストールすると、この CA によって発行された証明書は自動的に信頼されます。確認されていない拇印付きの証明書をインストールすることは、セキュリティ上、危険です。 [はい] をクリックすると、この危険を認識したことになります。

    この証明書をインストールしますか?
    ```

16. Outlook を起動して、現象を再現させます。
17. 以下の手順で Fiddler を停止して保存します。
18. [2. Stop Capture] をクリックします。
19. [3. Save Capture] をクリックします。
20. ファイルを Collect-OutlookInfo の "Path" パラメータに指定したフォルダ配下に作成された GUID 名のフォルダに保存します。

    ⚠️ パスワードの長さは 8 文字以上にする必要があります。

21. 左上の [...] から [Certificate]-[Remove CA Certificate] をクリックします。

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

22. Fiddler Everywhere Reporter を終了します。
23. コンソールに Enter キーを入力しトレースを停止します。

手順 10 で出力先に指定したフォルダに作成された `"Outlook_<マシン名>_<取得日時>.zip"` という名前の ZIP ファイルをお寄せください。
また、手順 20 で指定したパスワードも併せてお寄せください。

⚠️ もし Fiddler トレース (`.saz` ファイル) を別途保存した場合には、こちらのファイルも忘れずにお寄せください。

## ライセンス

Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。