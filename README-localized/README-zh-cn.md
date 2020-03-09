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
适用于 Office 365 Outlook Services 的 Apache Cordova 插件
=============================
提供 JavaScript API 以使用 Microsoft Office 365 Outlook 服务：“邮件”、“日历”、“联系人”和“事件”。
<!--
TODO review api compliance to 
http://msdn.microsoft.com/en-us/office/office365/howto/common-mail-tasks-client-library
-->
####Supported Platforms####

- Android（支持 cordova-android@>=4.0.0）
- iOS
- Windows（Windows 8.0、Windows 8.1 和 Windows Phone 8.1）

## 示例用法 ##
若要访问邮件 API，需要获取访问令牌并获取 Outlook 服务客户端。然后，可发送异步查询以便与邮件数据交互。注意：在 Microsoft Azure Active Directory 中注册应用时会分配应用程序 ID、授权和重定向 URI。

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
完整示例位于[此处](https://github.com/MSOpenTech/cordova-office-samples/tree/master/outlook-services/mailbox)。

## 安装说明 ##

使用 [Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) 创建应用并添加插件。

1. 确保已安装 Node.js 的最新版本，然后键入以下命令以安装 [Cordova CLI](https://github.com/apache/cordova-cli)：

        npm install -g cordova

2. 创建项目并添加希望支持的平台：

        cordova create outlookClientApp
        cd outlookClientApp
        cordova platform add windows <- support of Windows 8.0, Windows 8.1 and Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. 将插件添加到你的项目：

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-outlook

4. 生成并运行，例如：

        cordova run android

若要了解详细信息，请参阅 [Apache Cordova CLI 使用指南](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html)。

## 版权信息 ##
版权所有 (c) Microsoft Open Technologies, Inc.保留所有权利。

按照 Apache 许可 2.0 版本（称为“许可”）授予许可；要使用这些文件，必须遵循“许可”中的说明。你可以从以下网站获取许可的副本

http://www.apache.org/licenses/LICENSE-2.0

除非适用法律要求或书面同意，根据“许可”分配的软件“按原样”分配，不提供任何形式（无论是明示还是默示）的担保和条件。请参阅“许可证”了解“许可证”中管理权限和限制的特定语言。


此项目遵循 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
