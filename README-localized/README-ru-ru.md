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
Подключаемый модуль Apache Cordova для служб Outlook в Microsoft Office 365.
=============================
Предоставляет API на языке JavaScript для работы со службами Outlook в Microsoft Office 365 "Почта", "Календарь", "Контакты" и "События".
<!--
TODO review api compliance to 
http://msdn.microsoft.com/en-us/office/office365/howto/common-mail-tasks-client-library
-->
####Supported Platforms####

- Android (поддерживается cordova-android@>=4.0.0)
- iOS
- Windows (Windows 8.0, Windows 8.1 и Windows Phone 8.1)

## Пример использования ##
Для доступа к API почты необходимо получить маркер доступа и клиент служб Outlook. После этого можно отправлять асинхронные запросы для взаимодействия с данными почты. Примечание. Идентификатор приложения и URI перенаправления и авторизации назначаются при регистрации приложения в Microsoft Azure Active Directory.

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
Полный пример см. [здесь](https://github.com/MSOpenTech/cordova-office-samples/tree/master/outlook-services/mailbox).

## Инструкции по установке ##

Для создания своего приложения и добавления подключаемого модуля используйте [интерфейс командной строки Apache Cordova](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

1. Убедитесь, что установлена последняя версия Node.js, а затем введите следующую команду, чтобы установить [интерфейс командной строки Cordova](https://github.com/apache/cordova-cli):

        npm install -g cordova

2. Создайте проект и добавьте в него нужные платформы:

        cordova create outlookClientApp
        cd outlookClientApp
        cordova platform add windows <- support of Windows 8.0, Windows 8.1 and Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. Добавьте в проект подключаемый модуль:

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-outlook

4. Соберите и запустите, например:

        cordova run android

Дополнительные сведения см. в статье [Рекомендации по использованию интерфейса командной строки Apache Cordova](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

## Авторские права ##
(c) Microsoft Open Technologies, Inc. Все права защищены.

Предоставляется по лицензии Apache версии 2.0 ("Лицензия"); эти файлы можно использовать только в соответствии с Лицензией. Копию Лицензии можно получить по адресу:

http://www.apache.org/licenses/LICENSE-2.0

Программное обеспечение, распространяемое по Лицензии, распространяется на условиях «КАК ЕСТЬ», БЕЗ ГАРАНТИЙ ИЛИ УСЛОВИЙ ЛЮБОГО РОДА, явно выраженных или подразумеваемых, если такие гарантии или условия не требуются действующим законодательством или не согласованы в письменной форме. Конкретные юридические формулировки, регулирующие связанные с Лицензией разрешения и ограничения, содержатся в тексте Лицензии.


Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).
