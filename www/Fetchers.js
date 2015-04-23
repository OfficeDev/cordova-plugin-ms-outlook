// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var exec = require('cordova/exec');
var utils = require('./utility');
var Entity = require('./Entity');

utils.extends(Fetcher, Entity);

function Fetcher(context, path) {
    Entity.call(this, context, path);
}

Fetcher.prototype.fetch = function () {
    // abstract method
};

utils.extends(CollectionFetcher, Fetcher);
function CollectionFetcher(context, path) {
    Fetcher.call(this, context, path);
    this.reset();
}

CollectionFetcher.prototype.executeNativeMethod = function (nativeMethodName, resultType, payload, appendResultId) {
    var _this = this;
    var deferred = new utils.Utility.Deferred();

    this.context.getAccessTokenFn().then(
        function (token) {
            // To support native ADAL plugin
            if (token.accessToken) {
                token = token.accessToken;
            }

            var win = function(res){
                try {
                    var result = JSON.parse(res);
                    var resultArray = [];
                    result.value.forEach(function (resItem) {
                        var path = !!appendResultId ? _this.getPath(resItem.Id) : _this.path;
                        resultArray.push(new resultType(_this.context, path, resItem));
                    });
                    deferred.resolve(resultArray);
                } catch (e) {
                    deferred.reject(e);
                }
            };
            
            var fail = function (err) {
                // in most cases error callback returns serialized error object so we need to parse it
                if (typeof err === "string") {
                    try {
                        err = JSON.parse (err);
                    } catch(ex) {}
                }
                deferred.reject(err);
            };

            var nativeArguments = [token, _this.context.serviceRootUri, _this.path].concat(payload);
            exec(win, fail, "OutlookServices", nativeMethodName, nativeArguments);
        }, function(err) {
            deferred.reject(err);
        }
    );

    return deferred;
};

CollectionFetcher.prototype.fetch = function (prop) {
    // abstract method
};

CollectionFetcher.prototype.fetchAll = function () {
    // abstract method
};

CollectionFetcher.prototype.reset = function () {
    this._top = -1;
    this._skip = -1;
    this._selectedId = null;
    this._select = null;
    this._expand = null;
    this._filter = null;
};

CollectionFetcher.prototype.top = function (top) {
    this._top = top;
    return this;
};

CollectionFetcher.prototype.skip = function(skip) {
    this._skip = skip;
    return this;
};

CollectionFetcher.prototype.select = function(select) {
    this._select = select;
    return this;
};

CollectionFetcher.prototype.expand = function(expand) {
    this._expand = expand;
    return this;
};

CollectionFetcher.prototype.filter = function(filter) {
    this._filter = filter;
    return this;
};

module.exports.Fetcher = Fetcher;
module.exports.CollectionFetcher = CollectionFetcher;

