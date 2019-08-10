## graph-client

このソリューションはSharePointのWebパーツにOutlookの予定表データを表示するためのサンプルです。  
yo @microsoft/sharepoint コマンドで作成した雛形に対して、以下の変更を加えています。  
 * パッケージ追加(npm install @microsoft/microsoft-graph-types --save-dev)
 * config > pakage-solution.jsonファイルにwebApiPermissionRequestsを追加
 * src > webparts > graphClient > GraphClientWebParts.tsファイルにコードを追加
 * 同フォルダ > GraphClientWebPart.module.scssファイルにレイアウト用スタイルを追加

### ビルド方法

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

* graph-clientフォルダをVisual Studio Codeで開く
* ターミナルで以下コマンドを順次実行
* npm i
* gulp build --ship
* gulp bundle --ship
* gulp package-solution --ship
* sharepointフォルダ > solutionフォルダ > graph-client.sppkgが出来れば成功

### デプロイ方法

* ビルド方法に従い作成したgraph-client.sppkgをSharePointのアプリカタログサイトにアップロード
* エラーが無く、展開済であることを確認
* SharePoint管理センター > APIの管理 画面で、Microsoft Graphのアクセス許可を承認
* 任意のSharePointサイトでアプリを追加(アプリ名：graph-client-client-side-solution)
* 同サイトの任意のページにWebパーツを追加(Webパーツ名：graph-client)