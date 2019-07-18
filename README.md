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
Apache Cordova plugin for Office 365 Outlook Services
=============================
Provides JavaScript API to work with Microsoft Office 365 Outlook Services: Mail, Calandar, Contacts and Events.
<!--
TODO review api compliance to 
http://msdn.microsoft.com/en-us/office/office365/howto/common-mail-tasks-client-library
-->
####Supported Platforms####

- Android (cordova-android@>=4.0.0 is supported)
- iOS
- Windows (Windows 8.0, Windows 8.1 and Windows Phone 8.1)

## Sample usage ##
To access the Mail API you need to acquire an access token and get the Outlook Services client. Then, you can send async queries to interact with mail data. Note: application ID, authorization and redirect URIs are assigned when you register your app with Microsoft Azure Active Directory.

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
Complete example is available [here](https://github.com/MSOpenTech/cordova-office-samples/tree/master/outlook-services/mailbox).

## Installation Instructions ##

Use [Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) to create your app and add the plugin.

1. Make sure an up-to-date version of Node.js is installed, then type the following command to install the [Cordova CLI](https://github.com/apache/cordova-cli):

        npm install -g cordova

2. Create a project and add the platforms you want to support:

        cordova create outlookClientApp
        cd outlookClientApp
        cordova platform add windows <- support of Windows 8.0, Windows 8.1 and Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. Add the plugin to your project:

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-outlook

4. Build and run, for example:

        cordova run android

To learn more, read [Apache Cordova CLI Usage Guide](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

## Copyrights ##
Copyright (c) Microsoft Open Technologies, Inc. All rights reserved.

Licensed under the Apache License, Version 2.0 (the "License"); you may not use these files except in compliance with the License. You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
