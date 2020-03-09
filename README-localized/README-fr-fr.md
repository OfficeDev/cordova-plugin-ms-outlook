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
Plug-in Apache Cordova pour Office 365 Services Outlook
=============================
Fournit l'API JavaScript pour l'utiliser avec Microsoft Office 365 Services Outlook : Courrier, calendriers, contacts et événements.
<!--
TODO review api compliance to 
http://msdn.microsoft.com/en-us/office/office365/howto/common-mail-tasks-client-library
-->
####Supported Platforms####

- Android (cordova-android@>=4.0.0 est pris en charge)
- iOS
- Windows (Windows 8.0, Windows 8.1 et Windows Phone 8.1)

## Exemple d’utilisation ##
Vous devez obtenir un jeton d'accès et Services Outlook client pour accéder à l'API Courrier. Vous pouvez ensuite envoyer des requêtes asynchrones pour interagir avec les données de courrier. Remarque : les ID d’application, les URI d’autorisation et de redirection sont attribués lorsque vous enregistrez votre application avec Microsoft Azure Active Directory.

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
Un exemple entier est disponible [ici](https://github.com/MSOpenTech/cordova-office-samples/tree/master/outlook-services/mailbox).

## Instructions d’installation ##

Utilisez [Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) pour créer votre application et ajouter le plug-in.

1. Assurez-vous qu'une version à jour de Node.js est installée, puis tapez la commande suivante pour installer [Cordova CLI](https://github.com/apache/cordova-cli) :

        npm install -g cordova

2. Créez un projet et ajoutez les plateformes que vous souhaitez prendre en charge :

        cordova create outlookClientApp
        cd outlookClientApp
        cordova platform add windows <-support de Windows 8.0, Windows 8.1 et Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. Ajoutez le plug-in à votre projet :

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-outlook

4. Créez et exécutez, par exemple :

        cordova run android

Pour en savoir plus, consultez le [Guide d’utilisation de CLI Apache Cordova](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

## Copyrights ##
Copyright (c) Microsoft Open Technologies, Inc. Tous droits réservés.

Sous licence Apache, version 2.0 (la « License »); vous devez utiliser ces fichiers conformément à la Licence. Vous pouvez obtenir une copie de la Licence sur 

http://www.apache.org/licenses/LICENSE-2.0

Sauf exigence par une loi applicable ou accord écrit, tout logiciel distribué dans le cadre de la Licence est fourni « EN L'ÉTAT », SANS GARANTIE OU CONDITION D'AUCUNE SORTE, explicite ou implicite. Consultez la Licence pour les dispositions spécifiques régissant les autorisations et limitations dans le cadre de la License.


Ce projet a adopté le [Code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
