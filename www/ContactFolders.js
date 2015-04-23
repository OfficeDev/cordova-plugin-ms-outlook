// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Fetcher = require('./Fetchers').Fetcher;
var CollectionFetcher = require('./Fetchers').CollectionFetcher;

var Contacts = require('./Contacts').Contacts;

utils.extends(ContactFolder, Entity);
function ContactFolder(context, path, data) {
    Entity.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.ParentFolderId = data.ParentFolderId;
    this.DisplayName = data.DisplayName;
}
                                      
Object.defineProperty(ContactFolder.prototype, "childFolders", {
    get: function () {
        if (this._childFolders === undefined) {
            this._childFolders = new ContactFolders(this.context, this.getPath("ChildFolders"));
        }
        return this._childFolders;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(ContactFolder.prototype, "contacts", {
    get: function () {
        if (this._contacts === undefined) {
            this._contacts = new Contacts(this.context, this.getPath("Contacts"));
        }
        return this._contacts;
    },
    enumerable: true,
    configurable: true
});

ContactFolder.prototype.update = function () {
    return this.executeNativeMethod("updateContactFolder", ContactFolder, JSON.stringify(this));
};

ContactFolder.prototype.delete = function () {
    return this.executeNativeMethod("deleteContactFolder");
};

Object.defineProperty(ContactFolder.prototype, "contacts", {
    get: function () {
        if (this._contacts === undefined) {
            this._contacts = new Contacts(this.context, this.getPath('Contacts'));
        }
        return this._contacts;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(ContactFolder.prototype, "childFolders", {
    get: function () {
        if (this._childFolders === undefined) {
            this._childFolders = new ContactFolders(this.context, this.getPath('ChildFolders'));
        }
        return this._childFolders;
    },
    enumerable: true,
    configurable: true
});

utils.extends(ContactFolders, Entity);
function ContactFolders(context, path) {
    Entity.call(this, context, path);
}

ContactFolders.prototype.getContactFolder = function (id) {
    return new ContactFolderFetcher(this.context, this.getPath(id), id);
};

ContactFolders.prototype.getContactFolders = function () {
    return new ContactFolderCollectionFetcher(this.context, this.path);
};

ContactFolders.prototype.addContactFolder = function (item) {
    var payload = JSON.stringify(ContactFolder.prototype.preparePayload.call(item));
    return this.executeNativeMethod("addContactFolder", ContactFolder, payload, true);
};

utils.extends(ContactFolderFetcher, Fetcher);
function ContactFolderFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

Object.defineProperty(ContactFolderFetcher.prototype, "contacts", {
    get: function () {
        if (this._contacts === undefined) {
            this._contacts = new Contacts(this.context, this.getPath('Contacts'));
        }
        return this._contacts;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(ContactFolderFetcher.prototype, "childFolders", {
    get: function () {
        if (this._childFolders === undefined) {
            this._childFolders = new ContactFolders(this.context, this.getPath('ChildFolders'));
        }
        return this._childFolders;
    },
    enumerable: true,
    configurable: true
});

ContactFolderFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getContactFolder", ContactFolder, this._id);
};

utils.extends(ContactFolderCollectionFetcher, CollectionFetcher);
function ContactFolderCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

ContactFolderCollectionFetcher.prototype.fetch = function () {
    return this.fetchAll();
};

ContactFolderCollectionFetcher.prototype.fetchAll = function () {
    
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });

    return this.executeNativeMethod("getContactFolders", ContactFolder, queryParams, true);
};

module.exports.ContactFolder = ContactFolder;
module.exports.ContactFolders = ContactFolders;
module.exports.ContactFolderFetcher = ContactFolderFetcher;
module.exports.ContactFolderCollectionFetcher = ContactFolderCollectionFetcher;

