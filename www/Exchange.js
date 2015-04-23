// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var Users = require('./Users').Users;
var UserFetcher = require('./Users').UserFetcher;
var Deferred = require('./utility').Utility.Deferred;

var Exchange = {
};

Exchange.Client = Client;

function DataContext(serviceRootUri, authContext, resourceUrl, appId, redirectUrl) {
    this.serviceRootUri = serviceRootUri;

    this.getAccessTokenFn = function () {
        var d = new Deferred();

        authContext.tokenCache.readItems().then(function (tokenCacheItems) {
            var correspondingCacheItem = tokenCacheItems.filter(function (item) {
                return item.clientId === appId && item.resource === resourceUrl;
            })[0];

            if (correspondingCacheItem == null) {
                authContext.acquireTokenAsync(resourceUrl, appId, redirectUrl).then(function (authResult) {
                    d.resolve(authResult.accessToken);
                }, function (err) {
                    d.reject(err);
                });
            } else {
                authContext.acquireTokenSilentAsync(resourceUrl, appId, correspondingCacheItem.userInfo && correspondingCacheItem.userInfo.userId).then(function (authResult) {
                    d.resolve(authResult.accessToken);
                }, function (err) {
                    authContext.acquireTokenAsync(resourceUrl, appId, redirectUrl).then(function (authResult) {
                        d.resolve(authResult.accessToken);
                    }, function (err) {
                        d.reject(err);
                    });
                });
            }
        }, function (err) {
            d.reject(err);
        });

        return d;
    };
}

function Client(serviceRootUri, authContext, resourceUrl, clientId, redirectUrl) {
    this.context = new DataContext(serviceRootUri, authContext, resourceUrl, clientId, redirectUrl);
}

Client.prototype.getPath = function (prop) {
    return this.context.serviceRootUri + '/' + prop;
};

Object.defineProperty(Client.prototype, "me", {
    get: function () {
        if (this._me === undefined) {
            this._me = new UserFetcher(this.context, this.getPath("Me"), "me");
        }
        return this._me;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(Client.prototype, "users", {
    get: function () {
        if (this._users === undefined) {
            this._users = new Users(this.context, this.getPath("Users"));
        }
        return this._users;
    },
    enumerable: true,
    configurable: true
});

module.exports = Exchange;
