## 概要
OutlookTrace.ps1 は Outlook に関する情報採取用の PowerShell スクリプトです。

[ダウンロード](https://github.com/jpmessaging/OutlookTrace/releases/download/v2019-12-13/OutlookTrace.ps1)

## 利用方法
1. OutlookTrace.ps1 をダウンロードし、ブロックを解除します。

    1. ファイルを右クリックして、プロパティを開きます。  
    2. [全般] タブにて、「このファイルは他のコンピューターから取得したものです。このコンピューターを保護するため、このファイルへのアクセスはブロックされる可能性があります。」というメッセージが表示されている場合には、[許可する] にチェックを入れます。  

2. 対象のマシン上に OutlookTrace.ps1 をコピーします。
3. 管理者権限で PowerShell を起動します。

   Get-ExecutionPolicy を実行して RemoteSigned となっていない場合には以下のように設定します。

    ```PowerShell
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
    ```
       
4. ドット ソースで OutlookTrace.ps1 をインポートします。

    ```
    . <OutlookTrace.ps1 へのパス>
    ```

    例: 
    ```
    . C:\temp\OutlookTrace.ps1
    ```

5. Collect-OutlookInfo を実行します  

    ※ 採取するコンポーネントについてはエンジニアからの案内をご確認ください。

    ```
    Collect-OutlookInfo -Path <出力先フォルダ> -Component <採取するコンポーネント>
    ```

    例:
    ```
    Collect-OutlookInfo -Path c:\temp -Component Configuration, Netsh, Outlook
    ```

6. 正常にトレースが開始されると、"Hit enter to stop tracing" と表示されるので、事象を再現します。
   
    ※ 採取するコンポーネントに Fiddler を含めた場合、[FiddlerCap Web Recorder] ダイアログボックスが表示されます。以下の手順に従って手動で、キャプチャを開始ください。キャプチャ開始後に事象を再現します。

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

7. 再現後、コンソールに Enter キーを入力しトレースを停止します。
8. 採取するコンポーネントに Fiddler を含めた場合、以下の手順で停止して保存します。

    1. [2. キャプチャ停止] をクリックします。
    2. [3. キャプチャ保存] をクリックします。
    3. ファイルを任意の場所に保存します。
    4. [FiddlerCap Web Recorder] ダイアログボックスをクローズします。  
        この時以下の内容が表示されたら、[はい] をクリックします。

        ```
        次の証明書をルート ストアから削除しますか?

        サブジェクト: DO_NOT_TRUST_FiddlerRoot, DO_NOT_TRUST, Created by http://www.fiddler2.com
        発行者: 自己発行
        有効期間: ***
        シリアル番号 : ***
        拇印 (sha1): ***
        拇印 (md5):***
        ```

9. 手順 3 で Set-ExecutionPolicy で変更した場合には元の値へ戻します。

    ```PowerShell
    Set-ExecutionPolicy -ExecutionPolicy <元の値>
    ```
    
以下のファイルをお寄せください。

- 手順 5 で出力先に指定したフォルダに作成された "Outlook_<サーバー名>_<取得日時>.zip" という名前の ZIP ファイル
- 採取するコンポーネントに Fiddler を含めた場合には、手順 8 で保存したファイル (FiddlerCap_***.saz) 

