// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');

var Fetcher = require('./Fetchers').Fetcher;
var CollectionFetcher = require('./Fetchers').CollectionFetcher;

var Messages = require('./Messages').Messages;

utils.extends(Folder, Entity);
function Folder(context, path, data) {
    Entity.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.ParentFolderId = data.ParentFolderId;
    this.DisplayName = data.DisplayName;
    this.ChildFolderCount = data.ChildFolderCount;
}

Folder.prototype.preparePayload = function () {
    var payload = { DisplayName: this.DisplayName };
    return payload;
};

Object.defineProperty(Folder.prototype, "childFolders", {
    get: function () {
        if (this._childFolders === undefined) {
            this._childFolders = new Folders(this.context, this.getPath("ChildFolders"));
        }
        return this._childFolders;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(Folder.prototype, "messages", {
    get: function () {
        if (this._messages === undefined) {
            this._messages = new Messages(this.context, this.getPath("Messages"));
        }
        return this._messages;
    },
    enumerable: true,
    configurable: true
});

Folder.prototype.copy = function (destinationId) {
    return this.executeNativeMethod("copyFolder", Folder, destinationId);
};

Folder.prototype.move = function (destinationId) {
    return this.executeNativeMethod("moveFolder", Folder, destinationId);
};

Folder.prototype.update = function () {
    var payload = JSON.stringify(this.preparePayload());
    return this.executeNativeMethod("updateFolder", Folder, payload);
};

Folder.prototype.delete = function () {
    // TODO: Refine if native method returns any folder object
    return this.executeNativeMethod("deleteFolder");
};

utils.extends(Folders, Entity);
function Folders(context, path) {
    Entity.call(this, context, path);
}

Folders.prototype.getFolder = function (id) {
    return new FolderFetcher(this.context, this.getPath(id), id);
};

Folders.prototype.getFolders = function () {
    return new FolderCollectionFetcher(this.context, this.path);
};

Folders.prototype.addFolder = function (item) {
    var payload = JSON.stringify(Folder.prototype.preparePayload.call(item));
    return this.executeNativeMethod("addFolder", Folder, payload, true);
};

utils.extends(FolderFetcher, Fetcher);
function FolderFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

Object.defineProperty(FolderFetcher.prototype, "childFolders", {
    get: function () {
        if (this._childFolders === undefined) {
            this._childFolders = new Folders(this.context, this.getPath("ChildFolders"));
        }
        return this._childFolders;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(FolderFetcher.prototype, "messages", {
    get: function () {
        if (this._messages === undefined) {
            this._messages = new Messages(this.context, this.getPath("Messages"));
        }
        return this._messages;
    },
    enumerable: true,
    configurable: true
});

FolderFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getFolder", Folder, this._id);
};

FolderFetcher.prototype.copy = function (destinationId) {
    return this.executeNativeMethod("copyFolder", Folder, destinationId);
};

FolderFetcher.prototype.move = function (destinationId) {
    return this.executeNativeMethod("moveFolder", Folder, destinationId);
};

utils.extends(FolderCollectionFetcher, CollectionFetcher);
function FolderCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

FolderCollectionFetcher.prototype.fetch = function () {
    return this.fetchAll();
};

FolderCollectionFetcher.prototype.fetchAll = function () {
    
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });

    return this.executeNativeMethod("getFolders", Folder, queryParams, true);
};

module.exports.Folder = Folder;
module.exports.Folders = Folders;
module.exports.FolderFetcher = FolderFetcher;
module.exports.FolderCollectionFetcher = FolderCollectionFetcher;
