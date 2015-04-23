// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var exec = require('cordova/exec');
var Deferred = require('cordova-plugin-ms-adal.utility').Utility.Deferred;

function Entity(context, path) {
    this.path = path;
    this.context = context;
}

Entity.prototype.getPath = function (prop) {
    return this.path + '/' + prop;
};

Entity.prototype.executeNativeMethod = function (nativeMethodName, resultType, payload, appendResultId) {
    var _this = this;
    var deferred = new Deferred();

    this.context.getAccessTokenFn().then(
        function (token) {
            var win = function(entity){
                if (!!resultType) {
                    try {
                        var result = JSON.parse(entity);
                        var path = !!appendResultId ? _this.getPath(result.Id) : _this.path;
                        deferred.resolve(new resultType(_this.context, path, result));
                    } catch (e) {
                        deferred.reject(e);
                    }
                } else {
                    deferred.resolve(null);
                }
            };

            var fail = function(err){
                // in most cases error callback returns serialized error object so we need to parse it
                if (typeof err === "string") {
                    try {
                        err = JSON.parse(err);
                    } catch(ex) {}
                }
                deferred.reject(err);
            };

            var nativeArguments = [token, _this.context.serviceRootUri, _this.path];
            nativeArguments = !!payload ?nativeArguments.concat(payload) : nativeArguments;
            exec(win, fail, "OutlookServices", nativeMethodName, nativeArguments);
        }, function (err) {
            deferred.reject(err);
        }
    );

    return deferred;
};

module.exports = Entity;

