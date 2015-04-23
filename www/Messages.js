// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Fetcher = require('./Fetchers').Fetcher;
var CollectionFetcher = require('./Fetchers').CollectionFetcher;

var Item = require('./Items').Item;
var ItemHelpers = require('./ItemHelpers');

var ItemBody = ItemHelpers.ItemBody;
var Recipient = ItemHelpers.Recipient;
var MeetingMessageType = ItemHelpers.MeetingMessageType;

utils.extends(Message, Item);
function Message(context, path, data) {
    Item.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.ParentFolderId = data.ParentFolderId;
    this.From = new Recipient(data.From);
    this.Sender = new Recipient(data.Sender);
    this.ToRecipients = data.ToRecipients && data.ToRecipients.map(function (recipient) {
        return new Recipient(recipient);
    });
    this.CcRecipients = data.CcRecipients && data.CcRecipients.map(function (recipient) {
        return new Recipient(recipient);
    });
    this.BccRecipients = data.BccRecipients && data.BccRecipients.map(function (recipient) {
        return new Recipient(recipient);
    });
    this.ReplyTo = data.ReplyTo && data.ReplyTo.map(function (recipient) {
        return new Recipient(recipient);
    });
    this.ConversationId = data.ConversationId;
    this.UniqueBody = new ItemBody(data.UniqueBody);
    this.DateTimeReceived = data.DateTimeReceived ? new Date(data.DateTimeReceived) : null;
    this.DateTimeSent = data.DateTimeSent ? new Date(data.DateTimeSent) : null;
    this.IsDeliveryReceiptRequested = !!data.IsDeliveryReceiptRequested;
    this.IsReadReceiptRequested = !!data.IsReadReceiptRequested;
    this.IsDraft = data.IsDraft;
    this.IsRead = data.IsRead;
    this.EventId = data.EventId;
    this.MeetingMessageType = MeetingMessageType[data.MeetingMessageType || "None"];
}

Message.prototype.preparePayload = function () {

    var payload = {
        Body: this.Body ? ItemHelpers.ItemBody.prototype.preparePayload.call(this.Body) : undefined,
        Categories: this.Categories,
        Importance: ItemHelpers.Importance[this.Importance],
        Subject: this.Subject,
        From: this.From || undefined,
        Sender: this.Sender || undefined,
        ToRecipients: this.ToRecipients || undefined,
        CcRecipients: this.CcRecipients || undefined,
        BccRecipients: this.BccRecipients || undefined,
        ReplyTo: this.ReplyTo || undefined
    };
    return payload;
};

Message.prototype.copy = function (destinationId) {
    return this.executeNativeMethod("copyMessage", Message, destinationId);
};

Message.prototype.move = function (destinationId) {
    return this.executeNativeMethod("moveMessage", Message, destinationId);
};

Message.prototype.update = function () {
    var payload = JSON.stringify(this.preparePayload());
    return this.executeNativeMethod("updateMessage", Message, payload);
};

Message.prototype.delete = function () {
    return this.executeNativeMethod("deleteMessage");
};

Message.prototype.createReply = function () {
    return this.executeNativeMethod("createReply", Message);
};

Message.prototype.createReplyAll = function () {
    return this.executeNativeMethod("createReplyAll", Message);
};

Message.prototype.createForward = function () {
    return this.executeNativeMethod("createForward", Message);
};

Message.prototype.reply = function (comment) {
    return this.executeNativeMethod("reply", null, comment);
};

Message.prototype.replyAll = function (comment) {
    return this.executeNativeMethod("replyAll", null, comment);
};

Message.prototype.forward = function (comment, toRecipients) {

    if (cordova.platformId == 'ios') { // ios proxy requires recipients as string
        toRecipients = JSON.stringify(toRecipients);
    }
    return this.executeNativeMethod("forward", null, [comment, toRecipients]);
};

Message.prototype.send = function () {
    return this.executeNativeMethod("send");

};

utils.extends(Messages, Entity);
function Messages(context, path) {
    Entity.call(this, context, path);
}

Messages.prototype.getMessage = function (id) {
    return new MessageFetcher(this.context, this.getPath(id), id);
};

Messages.prototype.getMessages = function () {
    return new MessageCollectionFetcher(this.context, this.path);
};

Messages.prototype.addMessage = function (item) {
    var payload = JSON.stringify(Message.prototype.preparePayload.call(item));
    return this.executeNativeMethod("addMessage", Message, payload, true);
};

utils.extends(MessageFetcher, Fetcher);
function MessageFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

MessageFetcher.prototype.copy = function (destinationId) {
    return this.executeNativeMethod("copyMessage", Message, destinationId);
};

MessageFetcher.prototype.move = function (DestinationId) {
    return this.executeNativeMethod("moveMessage", Message, destinationId);
};

MessageFetcher.prototype.update = function () {
    return this.executeNativeMethod("updateMessage", Message, JSON.stringify(this));
};

MessageFetcher.prototype.delete = function () {
    return this.executeNativeMethod("deleteMessage");
};

MessageFetcher.prototype.createReply = function () {
    return this.executeNativeMethod("createReply", Message, JSON.stringify(this));
};

MessageFetcher.prototype.createReplyAll = function () {
    return this.executeNativeMethod("createReplyAll", Message, JSON.stringify(this));
};

MessageFetcher.prototype.createForward = function () {
    return this.executeNativeMethod("createForward", Message, JSON.stringify(this));
};

MessageFetcher.prototype.reply = function (comment) {
    return this.executeNativeMethod("reply", null, comment);
};

MessageFetcher.prototype.replyAll = function (comment) {
    return this.executeNativeMethod("replyAll", null, comment);
};

MessageFetcher.prototype.forward = function (comment, toRecipients) {

    if (cordova.platformId == 'ios') { // ios proxy requires recipients as string
        toRecipients = JSON.stringify(toRecipients);
    }
    return this.executeNativeMethod("forward", null, [comment, toRecipients]);
};

MessageFetcher.prototype.send = function () {
    return this.executeNativeMethod("send");
};

MessageFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getMessage", Message, this._id);
};

utils.extends(MessageCollectionFetcher, CollectionFetcher);
function MessageCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

MessageCollectionFetcher.prototype.fetch = function () {
    return this.fetchAll();
};

MessageCollectionFetcher.prototype.fetchAll = function () {
    
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });

    return this.executeNativeMethod("getMessages", Message, queryParams, true);
};

module.exports.Message = Message;
module.exports.Messages = Messages;
