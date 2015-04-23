// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Fetcher = require('./Fetchers').Fetcher;
var CollectionFetcher = require('./Fetchers').CollectionFetcher;

var Contacts = require('./Contacts').Contacts;
var ContactFolders = require('./ContactFolders').ContactFolders;
var Calendars = require('./Calendars').Calendars;
var CalendarFetcher = require('./Calendars').CalendarFetcher;
var CalendarGroups = require('./CalendarGroups').CalendarGroups;
var Events = require('./Events').Events;
var Folders = require('./Folders').Folders;
var FolderFetcher = require('./Folders').FolderFetcher;
var Messages = require('./Messages').Messages;

utils.extends(User, Entity);
function User(context, path, data) {
    Entity.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = data.Id = data.Id;

    this.DisplayName = data.DisplayName;
    this.Alias = data.Alias;
    this.MailboxGuid = data.MailboxGuid;
}

Object.defineProperty(User.prototype, "contacts", {
    get: function () {
        if (this._contacts === undefined) {
            this._contacts = new Contacts(this.context, this.getPath('Contacts'));
        }
        return this._contacts;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "calendar", {
    get: function () {
        if (this._calendar === undefined) {
            this._calendar = new CalendarFetcher(this.context, this.getPath('Calendar'), "Calendar");
        }
        return this._calendar;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "calendars", {
    get: function () {
        if (this._calendars === undefined) {
            this._calendars = new Calendars(this.context, this.getPath('Calendars'));
        }
        return this._calendars;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "events", {
    get: function () {
        if (this._events === undefined) {
            this._events = new Events(this.context, this.getPath('Events'));
        }
        return this._events;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "messages", {
    get: function () {
        if (this._messages === undefined) {
            this._messages = new Messages(this.context, this.getPath('Messages'));
        }
        return this._messages;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "folders", {
    get: function () {
        if (this._folders === undefined) {
            this._folders = new Folders(this.context, this.getPath('Folders'));
        }
        return this._folders;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "rootFolder", {
    get: function () {
        if (this._rootFolder === undefined) {
            this._rootFolder = new FolderFetcher(this.context, this.getPath('Folders/RootFolder'));
        }
        return this._rootFolder;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "inbox", {
    get: function () {
        if (this._inbox === undefined) {
            this._inbox = new FolderFetcher(this.context, this.getPath('Folders/Inbox'));
        }
        return this._inbox;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "drafts", {
    get: function () {
        if (this._drafts === undefined) {
            this._drafts = new FolderFetcher(this.context, this.getPath('Folders/Drafts'));
        }
        return this._drafts;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "sentItems", {
    get: function () {
        if (this._sentItems === undefined) {
            this._sentItems = new FolderFetcher(this.context, this.getPath('Folders/SentItems'));
        }
        return this._sentItems;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "deletedItems", {
    get: function () {
        if (this._deletedItems === undefined) {
            this._deletedItems = new FolderFetcher(this.context, this.getPath('Folders/DeletedItems'));
        }
        return this._deletedItems;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "contactFolders", {
    get: function () {
        if (this._contactFolders === undefined) {
            this._contactFolders = new ContactFolders(this.context, this.getPath('ContactFolders/Contacts/childFolders'));
        }
        return this._contactFolders;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(User.prototype, "calendarGroups", {
    get: function () {
        if (this._calendarGroups === undefined) {
            this._calendarGroups = new CalendarGroups(this.context, this.getPath('CalendarGroups'));
        }
        return this._calendarGroups;
    },
    enumerable: true,
    configurable: true
});

User.prototype.update = function () {
    return this.executeNativeMethod("updateUser", User, JSON.stringify(this));
};

User.prototype.delete = function () {
    return this.executeNativeMethod("deleteUser");
};

utils.extends(Users, Entity);
function Users(context, path) {
    Entity.call(this, context, path);
}

Users.prototype.getUser = function (id) {
    return new UserFetcher(this.context, this.getPath(id), id);
};

Users.prototype.getUsers = function () {
    return new UserCollectionFetcher(this.context, this.path);
};

Users.prototype.addUser = function (item) {
    return this.executeNativeMethod("addUser", User, JSON.stringify(item), true);
};

utils.extends(UserFetcher, Fetcher);
function UserFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

Object.defineProperty(UserFetcher.prototype, "contacts", {
    get: function () {
        if (this._contacts === undefined) {
            this._contacts = new Contacts(this.context, this.getPath('Contacts'));
        }
        return this._contacts;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "calendar", {
    get: function () {
        if (this._calendar === undefined) {
            this._calendar = new CalendarFetcher(this.context, this.getPath('Calendar'));
        }
        return this._calendar;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "calendars", {
    get: function () {
        if (this._calendars === undefined) {
            this._calendars = new Calendars(this.context, this.getPath('Calendars'));
        }
        return this._calendars;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "events", {
    get: function () {
        if (this._events === undefined) {
            this._events = new Events(this.context, this.getPath('Events'));
        }
        return this._events;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "messages", {
    get: function () {
        if (this._messages === undefined) {
            this._messages = new Messages(this.context, this.getPath('Messages'));
        }
        return this._messages;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "folders", {
    get: function () {
        if (this._folders === undefined) {
            this._folders = new Folders(this.context, this.getPath('Folders/RootFolder/ChildFolders'));
        }
        return this._folders;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "rootFolder", {
    get: function () {
        if (this._rootFolder === undefined) {
            this._rootFolder = new FolderFetcher(this.context, this.getPath('Folders/RootFolder'), "RootFolder");
        }
        return this._rootFolder;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "inbox", {
    get: function () {
        if (this._inbox === undefined) {
            this._inbox = new FolderFetcher(this.context, this.getPath('Folders/Inbox'), "Inbox");
        }
        return this._inbox;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "drafts", {
    get: function () {
        if (this._drafts === undefined) {
            this._drafts = new FolderFetcher(this.context, this.getPath('Folders/Drafts'), "Drafts");
        }
        return this._drafts;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "sentItems", {
    get: function () {
        if (this._sentItems === undefined) {
            this._sentItems = new FolderFetcher(this.context, this.getPath('Folders/SentItems'), "SentItems");
        }
        return this._sentItems;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "deletedItems", {
    get: function () {
        if (this._deletedItems === undefined) {
            this._deletedItems = new FolderFetcher(this.context, this.getPath('Folders/DeletedItems'), "DeletedItems");
        }
        return this._deletedItems;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "contactFolders", {
    get: function () {
        if (this._contactFolders === undefined) {
            this._contactFolders = new ContactFolders(this.context, this.getPath('ContactFolders'));
        }
        return this._contactFolders;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(UserFetcher.prototype, "calendarGroups", {
    get: function () {
        if (this._calendarGroups === undefined) {
            this._calendarGroups = new CalendarGroups(this.context, this.getPath('CalendarGroups'));
        }
        return this._calendarGroups;
    },
    enumerable: true,
    configurable: true
});

UserFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getUser", User, this._id);
};

utils.extends(UserCollectionFetcher, CollectionFetcher);
function UserCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

UserCollectionFetcher.prototype.fetch = function () {
    return this.fetchAll();
};

UserCollectionFetcher.prototype.fetchAll = function () {
    
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });

    return this.executeNativeMethod("getUsers", User, queryParams, true);
};

module.exports.User = User;
module.exports.Users = Users;
module.exports.UserFetcher = UserFetcher;
module.exports.UserCollectionFetcher = UserCollectionFetcher;
