// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Fetchers = require('./Fetchers');
var Fetcher = Fetchers.Fetcher;
var CollectionFetcher = Fetchers.CollectionFetcher;
var Calendars = require('./Calendars').Calendars;

utils.extends(CalendarGroup, Entity);
function CalendarGroup(context, path, data) {
    Entity.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.Name = data.Name;
    this.ChangeKey = data.ChangeKey;
    this.ClassId = data.ClassId;
}

CalendarGroup.prototype.preparePayload = function () {
    var payload = { Name: this.Name };
    return payload;
};

Object.defineProperty(CalendarGroup.prototype, "calendars", {
    get: function () {
        if (this._calendars === undefined) {
            this._calendars = new Calendars(this.context, this.getPath('Calendars'));
        }
        return this._calendars;
    },
    enumerable: true,
    configurable: true
});

CalendarGroup.prototype.update = function () {
    var payload = JSON.stringify(this.preparePayload());
    return this.executeNativeMethod("updateCalendarGroup", CalendarGroup, payload);
};

CalendarGroup.prototype.delete = function () {
    return this.executeNativeMethod("deleteCalendarGroup");
};

utils.extends(CalendarGroups, Entity);
function CalendarGroups(context, path) {
    Entity.call(this, context, path);
}

CalendarGroups.prototype.getCalendarGroup = function (id) {
    return new CalendarGroupFetcher(this.context, this.getPath(id), id);
};

CalendarGroups.prototype.getCalendarGroups = function () {
    return new CalendarGroupCollectionFetcher(this.context, this.path);
};

CalendarGroups.prototype.addCalendarGroup = function (item) {
    var payload = JSON.stringify(CalendarGroup.prototype.preparePayload.call(item));
    return this.executeNativeMethod("addCalendarGroup", CalendarGroup, payload, true);
};

utils.extends(CalendarGroupFetcher, Fetcher);
function CalendarGroupFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

Object.defineProperty(CalendarGroupFetcher.prototype, "calendars", {
    get: function () {
        if (this._calendars === undefined) {
            this._calendars = new Calendars(this.context, this.getPath('Calendars'));
        }
        return this._calendars;
    },
    enumerable: true,
    configurable: true
});

CalendarGroupFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getCalendarGroup", CalendarGroup);
};

utils.extends(CalendarGroupCollectionFetcher, CollectionFetcher);
function CalendarGroupCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

CalendarGroupCollectionFetcher.prototype.fetch = function (count) {
    return this.fetchAll();
};

CalendarGroupCollectionFetcher.prototype.fetchAll = function () {
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });
    return this.executeNativeMethod("getCalendarGroups", CalendarGroup, queryParams, true);
};

module.exports.CalendarGroup = CalendarGroup;
module.exports.CalendarGroups = CalendarGroups;
module.exports.CalendarGroupFetcher = CalendarGroupFetcher;
module.exports.CalendarGroupCollectionFetcher = CalendarGroupCollectionFetcher;
