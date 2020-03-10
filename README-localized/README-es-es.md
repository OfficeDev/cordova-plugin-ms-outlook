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
Complemento de Apache Cordova para Servicios de Outlook de Office 365
=============================
Proporciona la API de JavaScript para funcionar con los Servicios de Outlook de Microsoft Office 365: Correo, Calendario, Contactos y Eventos.
<!--
TODO review api compliance to 
http://msdn.microsoft.com/en-us/office/office365/howto/common-mail-tasks-client-library
-->
####Supported Platforms####

- Android (compatible con cordova-android@>=4.0.0)
- iOS
- Windows (Windows 8.0, Windows 8.1 y Windows Phone 8.1)

## Ejemplo de uso ##
Para acceder a la API de correo, debe adquirir un token de acceso y obtener el cliente de Servicios de Outlook. Después, puede enviar consultas asincrónicas para interactuar con los datos de correo. Nota: los URI de redirección, autorización y del identificador de la aplicación se asignan al registrar la aplicación en Microsoft Azure Active Directory.

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
El ejemplo completo se encuentra disponible [aquí](https://github.com/MSOpenTech/cordova-office-samples/tree/master/outlook-services/mailbox).

## Instrucciones de instalación ##

Use [CLI de Apache Cordova](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) para crear su aplicación y agregar el complemento.

1. Asegúrese de que tiene instalada una versión actualizada de Node.js y después, escriba el siguiente comando para instalar el [CLI de Cordova](https://github.com/apache/cordova-cli):

        npm install -g cordova

2. Cree un proyecto y agregue las plataformas que desea admitir:

        cordova create outlookClientApp
        cd outlookClientApp
        cordova platform add windows <- support of Windows 8.0, Windows 8.1 and Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. Agregue el complemento a su proyecto:

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-outlook

4. Compile y ejecute, por ejemplo:

        cordova run android

Para obtener más información, consulte [Guía de uso de CLI Apache Cordova](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

## Derechos de autor ##
Copyright (c) Microsoft Open Technologies, Inc. Todos los derechos reservados.

Con licencia bajo la Licencia de Apache, Versión 2.0 (la "Licencia"); es posible que no pueda usar estos archivos excepto en cumplimiento con la Licencia. Puede obtener una copia de la Licencia en

http://www.apache.org/licenses/LICENSE-2.0

Excepto si lo requiere la legislación vigente o es acordado por escrito, el software distribuido bajo la Licencia se distribuye "TAL CUAL", SIN GARANTÍAS O CONDICIONES DE NINGÚN TIPO, ya sea de forma explícita o implícita. Vea la Licencia para el idioma específico que rige los permisos y las limitaciones de la Licencia.


Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
