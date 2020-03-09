---
topic: sample
products:
- office-365
- office-outlook
languages:
- javascript
extensions:
  contentType: tools
  createdDate: 4/23/2015 8:16:04 AM
---
Office 365 Outlook サービス 用の Apache Cordova プラグイン
=============================
Microsoft Office 365 Outlook サービスと連携する JavaScript API を提供します。メール、カレンダー、連絡先、イベント。
<!--
TODO review api compliance to 
http://msdn.microsoft.com/en-us/office/office365/howto/common-mail-tasks-client-library
-->
####Supported Platforms####

- Android (cordova-android@>= 4.0.0 がサポートされています)
- iOS
- Windows (Windows 8.0、Windows 8.1、Windows Phone 8.1)

## サンプルの使用方法 ##
メール API にアクセスするには、アクセス トークンを取得し、Outlook サービス クライアントを取得する必要があります。その後、非同期クエリを送信してメール データを操作できます。注: アプリケーション ID、承認およびリダイレクト URI は、Microsoft Azure Active Directory にアプリを登録するときに割り当てられます。

```javascript
var resourceUrl = 'https://outlook.office365.com';
var officeEndpointUrl = 'https://outlook.office365.com/ews/odata';
var appId = '14b0c641-7fea-4e84-8557-25285eb86e43';
var authUrl = 'https://login.windows.net/common/';
var redirectUrl = 'http://localhost:4400/services/office365/redirectTarget.html';

var AuthenticationContext = Microsoft.ADAL.AuthenticationContext;

var outlookClient = new Microsoft.OutlookServices.Client(officeEndpointUrl,
    new AuthenticationContext(authUrl), resourceUrl, appId, redirectUrl);

outlookClient.me.folders.getFolder('Inbox').messages.getMessages().fetchAll().then(function (result) {
    result.forEach(function (msg) {
        console.log('Message "' + msg.Subject + '" received at "' + msg.DateTimeReceived.toString() + '"');
    });
}, function(error) {
    console.error(error);
});
```
完全な例は[こちら](https://github.com/MSOpenTech/cordova-office-samples/tree/master/outlook-services/mailbox)から入手できます。

## インストール手順 ##

[Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) を使用してアプリを作成し、プラグインを追加します。

1. Node.js の最新バージョンがインストールされていることを確認してから、次のコマンドを入力して [Cordova CLI](https://github.com/apache/cordova-cli) をインストールします。

        npm install -g cordova

2. プロジェクトを作成し、サポートするプラットフォームを追加します:

        cordova create outlookClientApp
        cd outlookClientApp
        cordova platform add windows <- Windows 8.0、Windows 8.1 および Windows Phone 8.1 のサポート
        cordova platform add android
        cordova platform add ios

3. プラグインをプロジェクトに追加します:

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-outlook

4. ビルドして実行します。例えば、以下のように操作します:

        cordova run android

詳細については、「[Apache Cordova CLI Usage Guide (Apache Cordova CLI 使用方法ガイド)](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html)」を参照してください。

## 著作権 ##
Copyright (c) Microsoft Open Technologies, Inc.All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License"); you may not use these files except in compliance with the License.You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.See the License for the specific language governing permissions and limitations under the License.


このプロジェクトでは、[Microsoft Open Source Code of Conduct (Microsoft オープン ソース倫理規定)](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
