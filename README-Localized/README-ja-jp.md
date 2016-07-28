# イメージ吹き出し Word アドイン サンプル: イメージの読み込み、編集、挿入

**目次**

* [概要](#summary)
* [必要なツール](#required-tools)
* [証明書のインストール方法](#how-to-install-certificates)
* [アプリの設定および実行方法](#how-to-set-up-and-run-the-app)
* [Word 2016 for Windows でのアドインの実行方法](#how-to-run-the-add-in-in-Word-2016-for-Windows)
* [FAQ](#faq)
* [質問とコメント](#questions-and-comments)
* [詳細を見る](#learn-more)


## まとめ

この Word アドイン サンプルには、以下を実行する方法が示されています。

1. Typescript を使用して Word アドインを作成します。
2. ドキュメントからアドインにイメージを読み込みます。
3. HTML キャンバス API を使用してアドイン内のイメージを編集し、Word 文書にそのイメージを挿入します。
4. リボンからアドインを起動するアドイン コマンドと、リボンとコンテキスト メニューの両方からスクリプトを実行するアドイン コマンドを実装します。
5. Office UI ファブリックを使用して、アドインに対してネイティブの Word のようなエクスペリエンスを作成します。

![](/readme-images/Word-Add-in-TypeScript-Canvas.gif)

定義 - **アドイン コマンド**: Word UI への拡張機能。これにより、アドインを作業ウィンドウで起動するか、またはスクリプトをリボンかコンテキスト メニューのいずれかから実行することができます。

この拡張機能の動作の確認だけをしたい場合は、「[Word 2016 for Windows の設定](#word-2016-for-windows-set-up)」にスキップしてこの[マニフェスト](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/blob/deploy2Azure/manifest-word-add-in-canvas.xml)を使用してください。

## 必要なツール

イメージ吹き出し Word アドイン サンプルを使用するには、以下が必要になります。

* Word 2016 16.0.6326.0000 以降、または Word Javascript API をサポートする任意のクライアント。このサンプルでは、JavaScript API をサポートするホストで実行されているか確認する要件チェックが実行されます。
* npm (https://www.npmjs.com/)。依存関係をインストールします。[NodeJS](https://nodejs.org/en/) に付属しています。
* (Windows の場合) [Git Bash](http://www.git-scm.com/downloads)。
* ローカル コンピューターにこのリポジトリのクローンを作成します。

> 注: Word for Mac 2016 は、現時点でアドイン コマンドをサポートしていません。このサンプルは、アドイン コマンドを使用せずに Mac で実行できます。

## 証明書のインストール方法

アドイン コマンドは HTTPS を必要とするため、このサンプルを実行するには証明書が必要になります。また、アドイン コマンドには UI がないため、無効な証明書を受け付けることができません。[./gen-cert.sh](#gen-cert.sh) を実行して証明書を作成した後、信頼されたルート証明機関ストアに ca.crt をインストールする必要があります (Windows の場合)。

## アプリの設定および実行方法

1. コマンド ラインに ```npm install typings -g``` と入力して、TypeScript 定義マネージャーをインストールします。
2. コマンド ラインを使用してプロジェクトのルート ディレクトリで ```typings install``` を実行し、typings.json で識別される Typescript 定義をインストールします。
3. プロジェクトのルート ディレクトリで ```npm install``` を実行し、package.json で識別されるプロジェクトが依存関係をインストールします。
4. gulp ```npm install -g gulp``` をインストールします。
5. ```gulp copy:libs``` を実行して、ファブリックと JQuery ファイルをコピーします。(Windows の場合) ここで問題が発生する場合、*%APPDATA%\npm* が path 変数に設定されていることを確認します。
6. プロジェクトのルート ディレクトリから ```gulp``` を実行することによって、既定の gulp タスクを実行します。TypeScript 定義が更新されない場合は、ここでエラーが発生します。

この時点で、このサンプル アドインが配置されたことになります。次に、Word がアドインを検索する場所を認識できるようにする必要があります。

### Word 2016 for Windows の設定

1. (Windows のみ) この[レジストリ キー](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/AddInCommandsUndark)を解凍し、実行することによって、アドイン コマンド機能を有効にします。これは、アドイン コマンドが**プレビュー機能**である間、必要です。
2. ネットワーク共有を作成するか、[ネットワークでフォルダーを共有し](https://technet.microsoft.com/ja-jp/library/cc770880.aspx)、そのフォルダーに [manifest-word-add-in-canvas.xml](manifest-word-add-in-canvas.xml) マニフェスト ファイルを配置します。
3. Word を起動し、ドキュメントを開きます。
4. [**ファイル**] タブを選択し、[**オプション**] を選択します。
5. [**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。
6. **[信頼されているアドイン カタログ]** を選択します。
7. **[カタログの URL]** ボックスに、manifest-word-add-in-canvas.xml があるフォルダー共有へのネットワーク パスを入力して、**[カタログの追加]** を選択します。
8. [**メニューに表示する**] チェック ボックスをオンにし、[**OK**] を選択します。
9. これらの設定が Office を次回起動したときに適用されることを示すメッセージが表示されます。Word を終了して、再起動します。

## Word 2016 for Windows でのアドインの実行方法

1. Word 文書を開きます。
2. Word 2016 の**[挿入]** タブで、**[マイ アドイン]** を選択します。
3. **[共有フォルダー]** タブを選択します。
4. **[イメージ吹き出しアドイン]** を選択し、**[挿入]** を選択します。
5. ご使用の Word バージョンでアドイン コマンドがサポートされている場合、UI によってアドインが読み込まれたことが通知されます。**[吹き出しアドイン]** タブを使用して、UI にアドインを読み込み、ドキュメントにイメージを挿入します。イメージをドキュメントに挿入する右クリックコンテキスト メニューを使用することもできます。
6. アドイン コマンドがご使用の Word バージョンによってサポートされていない場合は、アドインが作業ウィンドウに読み込まれます。アドインの機能を使用するために、Word 文書に画像を挿入することが必要になります。
7. Word 文書でイメージを選択し、*[doc からイメージを読み込む] を選択して作業ウィンドウにそのイメージを読み込みます。これで、吹き出しをイメージに挿入することができるようになりました。[イメージを doc に挿入する]* を選択して、更新されたイメージを Word 文書に配置します。アドインは、それぞれの吹き出しに対してプレース ホルダーの説明を生成します。

## FAQ

* アドイン コマンドは Mac や iPad で動作しますか。いいえ。このリリース ノートの公開時点では、アドイン コマンドは Mac または iPad では動作しません。
* 個人用のアドインが **[個人用アドイン]** ウィンドウに表示されないのはなぜですか。アドイン マニフェストにエラーがある可能性があります。[マニフェスト スキーマ](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Tools/XSD)に対してマニフェストを検証することをお勧めします。
* 個人用アドイン コマンドで、関数ファイルを呼び出せないのはなぜですか。アドイン コマンドでは、HTTPS が必要になります。アドイン コマンドでは TLS を必要とし、UI がないため、証明書に問題があるかどうか確認できません。作業ウィンドウで無効な証明書を受け入れる必要がある場合、アドイン コマンドは機能しません。
* npm インストール コマンドがハングアップしてしまうのはなぜですか。ハングアップしていない可能性があります。Windows 上での処理に時間がかかります。

## 質問とコメント

イメージ吹き出し Word アドイン サンプルについて、Microsoft にフィードバックをお寄せください。質問や提案につきましては、このリポジトリの「[問題](https://github.com/OfficeDev/Word-Add-in-TypeScript-Canvas/issues)」セクションに送信できます。

アドイン開発全般の質問については、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/Office365+API)」に投稿してください。質問またはコメントには、[office-js]、[word-addins]、[API] のタグを付けてください。これらのタグに注目しています。

## 詳細を見る

Word Javascript API ベースのアドインを作成するのに役立つその他のリソースを以下に示します。

* [Office アドイン プラットフォームの概要](https://msdn.microsoft.com/ja-jp/library/office/jj220082.aspx)
* [Word アドイン](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins.md)
* [Word アドインのプログラミングの概要](https://github.com/OfficeDev/office-js-docs/blob/master/word/word-add-ins-programming-guide.md)
* [Word のスニペット エクスプローラー](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)
* [Word アドインの JavaScript API リファレンス](https://github.com/OfficeDev/office-js-docs/tree/master/word/word-add-ins-javascript-reference)
* [SillyStories サンプル](https://github.com/OfficeDev/Word-Add-in-SillyStories) - サービスから docx ファイルを読み込んで、開いている Word 文書にファイルを挿入する方法について説明します。
* [Node.js 用の Office アドイン サーバー認証サンプル](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) - アドイン ユーザーを認証するために Azure および Google OAuth プロバイダーを使用する方法について説明します。

## 著作権
Copyright (c) 2016 Microsoft. All rights reserved.
