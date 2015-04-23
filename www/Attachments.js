// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Item = require('./Items').Item;
var ItemFetcher = require('./Items').ItemFetcher;
var Fetchers = require('./Fetchers');
var Fetcher = Fetchers.Fetcher;
var CollectionFetcher = Fetchers.CollectionFetcher;

var attachmentTypeChooser = function (context, path, data) {
    if (data['@odata.type']) {
        if (data['@odata.type'] === '#Microsoft.OutlookServices.FileAttachment')
            return new FileAttachment(context, path, data);
        if (data['@odata.type'] === '#Microsoft.OutlookServices.ItemAttachment')
            return new ItemAttachment(context, path, data);
        
        return new Attachment(context, path, data);
    }
};

utils.extends(Attachment, Entity);
function Attachment(context, path, data) {
    Entity.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.Name = data.Name;
    this.ContentType = data.ContentType;
    this.Size = data.Size;
    this.IsInline = data.IsInline;
    this.DateTimeLastModified = data.DateTimeLastModified && new Date(data.DateTimeLastModified);
}

Attachment.prototype.preparePayload = function () {
    // TODO: refine writeable properties for common Attachment type
    return {};
};

Attachment.prototype.update = function () {
    var payload = JSON.stringify(this.preparePayload());
    return this.executeNativeMethod("updateAttachment", attachmentTypeChooser, payload);
};

Attachment.prototype.delete = function () {
    return this.executeNativeMethod("deleteAttachment");
};

utils.extends(FileAttachment, Attachment);
function FileAttachment(context, path, data) {
    Attachment.call(this, context, path, data);

    if (!data) {
        return;
    }

    this['@odata.type'] = '#Microsoft.OutlookServices.FileAttachment';

    this.ContentId = data.ContentId;
    this.ContentLocation = data.ContentLocation;
    this.IsContactPhoto = data.IsContactPhoto;
    this.ContentBytes = data.ContentBytes;
}

FileAttachment.prototype.preparePayload = function () {
    var payload = {
        Name: this.Name,
        ContentBytes: this.ContentBytes,
        "@odata.type": "#Microsoft.OutlookServices.FileAttachment"
    };
    return payload;
};

utils.extends(ItemAttachment, Attachment);
function ItemAttachment(context, path, data) {
    Attachment.call(this, context, path, data);

    if (!data) {
        return;
    }

    this['@odata.type'] = '#Microsoft.OutlookServices.ItemAttachment';

    this.Item = data.Item;
}

ItemAttachment.prototype.preparePayload = function () {

    var item = new Item(null, null, this.Item);
    var itemPayload = item.preparePayload();

    var payload = {
        Name: this.Name,
        Item: itemPayload,
        "@odata.type": "#Microsoft.OutlookServices.ItemAttachment"
    };

    return payload;
};

Object.defineProperty(ItemAttachment.prototype, "item", {
    get: function () {
        if (this._item === undefined) {
            // TODO: ItemFetcher need to be updated with conditional logic (Item can be either Message or Event)
            this._item = new ItemFetcher(this.context, this.getPath("Item"));
        }
        return this._item;
    },
    enumerable: true,
    configurable: true
});

utils.extends(Attachments, Entity);
function Attachments(context, path) {
    Entity.call(this, context, path);
}

Attachments.prototype.getAttachment = function (id) {
    return new AttachmentFetcher(this.context, this.getPath(id), id);
};

Attachments.prototype.getAttachments = function () {
    return new AttachmentCollectionFetcher(this.context, this.path);
};

Attachments.prototype.addAttachment = function (item) {
    var payload = JSON.stringify(item.preparePayload());
    return this.executeNativeMethod("addAttachment", attachmentTypeChooser, payload, true);
};

utils.extends(AttachmentFetcher, Fetcher);
function AttachmentFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

AttachmentFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getAttachment", attachmentTypeChooser, this._id);
};

utils.extends(AttachmentCollectionFetcher, CollectionFetcher);
function AttachmentCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

AttachmentCollectionFetcher.prototype.fetch = function (count) {
    return this.fetchAll();
};

AttachmentCollectionFetcher.prototype.fetchAll = function () {
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });

    return this.executeNativeMethod("getAttachments", attachmentTypeChooser, queryParams, true);
};

module.exports.Attachment = Attachment;
module.exports.FileAttachment = FileAttachment;
module.exports.ItemAttachment = ItemAttachment;
module.exports.Attachments = Attachments;