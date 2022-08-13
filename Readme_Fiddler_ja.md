## 概要

OutlookTrace.psm1 は Outlook に関する情報採取用の PowerShell スクリプトです。

[ダウンロード](https://github.com/jpmessaging/OutlookTrace/releases/download/v2022-08-12/OutlookTrace.psm1)

## 利用方法

1. ログの出力先フォルダーを作成します。以下の説明では C:\temp を出力先フォルダーの例として使用します。
2. 以下のリンクをクリックし、FiddlerCapSetup.exe をダウンロードします。

    [FiddlerCapSetup](https://telerik-fiddler.s3.amazonaws.com/fiddler/FiddlerCapSetup.exe)

3. ダウンロードしたファイルをログの出力先フォルダーにコピーします。
4. Outlook を実行している場合には終了します。
5. OutlookTrace.psm1 をダウンロードして対象のマシン上にコピーします。
6. 管理者権限で cmd を起動します。
7. PowerShell を以下のように起動します。

    ```
    powershell -ExecutionPolicy Bypass
    ```

8. OutlookTrace.psm1 をインポートします。

    ```
    Import-Module <OutlookTrace.psm1 へのパス> -DisableNameChecking
    ```

    例:

    ```
    Import-Module C:\temp\OutlookTrace.psm1 -DisableNameChecking
    ```

9. Collect-OutlookInfo を実行します。

    ※ 採取するコンポーネントについてはエンジニアからの案内をご確認ください。

    ```
    Collect-OutlookInfo -Path <出力先フォルダ> -Component <採取するコンポーネント>
    ```

    例:

    ```
    Collect-OutlookInfo -Path C:\temp -Component Configuration, Outlook, Fiddler, Netsh, PSR, WAM
    ```

    正常にトレースが開始されると、`Hit enter to stop` と表示されます。

10. [FiddlerCap Web Recorder] ダイアログ ボックスが表示されない場合、**Outlook を開始するユーザー**にて出力先フォルダーの FiddlerCap フォルダーにある FiddlerCap.exe を実行します。

    ※ スクリプトの実行ユーザーと、情報採取のターゲットとなるユーザーが異なる場合、ターゲット ユーザー自身が FiddlerCap.exe を実行する必要があります。

11. [HTTPS 通信を解読] にチェックを入れます。
12. 以下の説明が表示されたら、内容を確認して [OK] をクリックします。

    ```
    HTTPS の解読は、HTTPS プロトコル経由で送られる Raw トラフィックを見るためにデバッグしやすくしてくれます。
    この機能は SSL トラフィックを解読し、ローカルに生成された証明書を用いて再度暗号化します。よって、この機能を使うと、不明な発行元からの証明書を使っているリモートサイトであること表示する、赤い警告ページが Web ブラウザーに表示されることを意味します。
    このトラフィックをキャプチャすることに限定して、このブラウザーに表示される警告を無視してください。
    ```

13. 以下の内容のセキュリティ警告が表示されたら、[はい] をクリックします。

    ```
    発行者が次であると主張する証明機関 (CA) から証明書をインストールしようとしています:

    DO_NOT_TRUST_FiddlerRoot

    証明書が実際に "DO_NOT_TRUST_FiddlerRoot" からのものであるかどうかを検証できません。"DO_NOT_TRUST_FiddlerRoot" に連絡して発行者を確認する必要があります。 次の番号はこの過程で役立ちます:

    拇印 (sha1): ***

    警告:
    このルート証明書をインストールすると、この CA によって発行された証明書は自動的に信頼されます。確認されていない拇印付きの証明書をインストールすることは、セキュリティ上、危険です。 [はい] をクリックすると、この危険を認識したことになります。

    この証明書をインストールしますか?
    ```

14. [1. キャプチャ開始] をクリックします。

    自動的にブラウザが起動されたら、そのブラウザはクローズいただいて結構です。

15. Outlook を起動して、現象を再現させます。
16. 以下の手順で Fiddler を停止して保存します。
17. [2. キャプチャ停止] をクリックします。
18. [3. キャプチャ保存] をクリックします。
19. [ファイルの種類] で `Password-Protected Capture (*.saz)` を選択します。
20. ファイルを Collect-OutlookInfo の "Path" パラメータに指定したフォルダ配下に作成された GUID 名のフォルダに保存します。
21. [Password-Protection Session Capture] のダイアログで任意のパスワードを入力し、[OK] をクリックします。
22. [FiddlerCap Web Recorder] ダイアログボックスをクローズします。
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

23. コンソールに Enter キーを入力しトレースを停止します。

手順 9 で出力先に指定したフォルダに作成された `"Outlook_<マシン名>_<取得日時>.zip"` という名前の ZIP ファイルをお寄せください。
また、手順 21 で指定したパスワードも併せてお寄せください。

## ライセンス

Copyright (c) 2021 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。