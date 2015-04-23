// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Fetchers = require('./Fetchers');
var Fetcher = Fetchers.Fetcher;
var ItemHelpers = require('./ItemHelpers');
var ItemBody = ItemHelpers.ItemBody;
var Importance = ItemHelpers.Importance;

utils.extends(Item, Entity);
function Item(context, path, data) {
    Entity.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.Body = data.Body && new ItemBody(data.Body);
    this.BodyPreview = data.BodyPreview;
    this.DateTimeCreated = data.DateTimeCreated ? new Date(data.DateTimeCreated) : null;
    this.DateTimeLastModified = data.DateTimeLastModified ? new Date(data.DateTimeLastModified) : null;
    this.Categories = data.Categories;
    this.ChangeKey = data.ChangeKey;
    this.ClassName = data.ClassName;
    this.HasAttachments = data.HasAttachments;
    this.Importance = Importance[data.Importance];
    this.Subject = data.Subject;
}

Item.prototype.preparePayload = function () {
    var payload = {
        Body: this.Body ? ItemBody.prototype.preparePayload.call(this.Body) : undefined,
        Categories: this.Categories,
        Importance: Importance[this.Importance],
        Subject: this.Subject
    };
    return payload;
};

Object.defineProperty(Item.prototype, "attachments", {
    get: function () {
        if (this._attachments === undefined) {
            var Attachments = require('./Attachments').Attachments;
            this._attachments = new Attachments(this.context, this.getPath("Attachments"));
        }
        return this._attachments;
    },
    enumerable: true,
    configurable: true
});

utils.extends(ItemFetcher, Fetcher);
function ItemFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

ItemFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getAttachmentItem", Item, this._id);
};

module.exports.Item = Item;
module.exports.ItemFetcher = ItemFetcher;
