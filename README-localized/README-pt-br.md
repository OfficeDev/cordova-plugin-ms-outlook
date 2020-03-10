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
Plug-in Apache Cordova para o Office 365 Outlook Services
=============================
Fornece a API JavaScript para funcionar com o Microsoft Office 365 Outlook Services: email, calendário, contatos e eventos.
<!--
TODO review api compliance to 
http://msdn.microsoft.com/en-us/office/office365/howto/common-mail-tasks-client-library
-->
####Supported Platforms####

- Android (compatível com cordova-android 4.0.0 ou mais recente)
- iOS
- Windows (Windows 8.0, Windows 8.1 e Windows Phone 8.1)

## Exemplo de uso ##
Para acessar a API Mail, você precisa adquirir um token de acesso e obter o cliente do Outlook Services. Em seguida, você pode enviar consultas assíncronas para interagir com os dados do correio. Observação: ID do aplicativo, autorização e URIs de redirecionamento são atribuídos quando você registra seu aplicativo no Microsoft Azure Active Directory.

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
Veja o [exemplo completo](https://github.com/MSOpenTech/cordova-office-samples/tree/master/outlook-services/mailbox).

## Instruções de instalação ##

Use [Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html) para criar seu aplicativo e adicionar o plug-in.

1. Verifique se uma versão atualizada do Node.js está instalada e digite o seguinte comando para instalar o [Cordova CLI](https://github.com/apache/cordova-cli):

        npm install -g cordova

2. Crie um projeto e adicione as plataformas que você quer usar:

        cordova create outlookClientApp
        cd outlookClientApp
        cordova platform add windows <- suporte do Windows 8.0, Windows 8.1 e Windows Phone 8.1
        cordova platform add android
        cordova platform add ios

3. Adicione o plug-in ao seu projeto:

        cordova plugin add https://github.com/OfficeDev/cordova-plugin-ms-outlook

4. Crie e execute, por exemplo:

        cordova run android

Saiba mais em [Guia de uso do Apache Cordova CLI](http://cordova.apache.org/docs/en/edge/guide_cli_index.md.html).

## Direitos autorais ##
Copyright (c) Microsoft Open Technologies, Inc. Todos os direitos reservados.

Licenciado nos termos da Licença do Apache, Versão 2.0 (a "Licença"); você não pode usar esses arquivos, exceto em conformidade com a Licença. Você encontra uma cópia da Licença em

http://www.apache.org/licenses/LICENSE-2.0

A menos que exigido pela lei aplicável ou acordado por escrito, o software distribuído nos termos da Licença é distribuído "COMO ESTÁ", SEM GARANTIAS OU CONDIÇÕES DE QUALQUER TIPO, expressas ou implícitas. Consulte a Licença para obter a linguagem específica que rege as permissões e limitações nos termos da Licença.


Este projeto adotou o [Código de Conduta de Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/).  Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
