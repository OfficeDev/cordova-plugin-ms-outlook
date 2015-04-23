// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Fetchers = require('./Fetchers');
var Fetcher = Fetchers.Fetcher;
var CollectionFetcher = Fetchers.CollectionFetcher;

var Events = require('./Events').Events;
utils.extends(Calendar, Entity);
function Calendar(context, path, data) {
    Entity.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.Name = data.Name;
    this.ChangeKey = data.ChangeKey;
}

Object.defineProperty(Calendar.prototype, "events", {
    get: function () {
        if (this._events === undefined) {
            this._events = new Events(this.context, this.getPath("Events"));
        }
        return this._events;
    },
    enumerable: true,
    configurable: true
});

Calendar.prototype.preparePayload = function() {
    var payload = { Name: this.Name };
    return payload;
};

Calendar.prototype.update = function () {
    var payload = JSON.stringify(this.preparePayload());
    return this.executeNativeMethod("updateCalendar", Calendar, payload);
};

Calendar.prototype.delete = function () {
    return this.executeNativeMethod("deleteCalendar");
};

module.exports.Calendar = Calendar;

utils.extends(Calendars, Entity);
function Calendars(context, path) {
    Entity.call(this, context, path);
}

Calendars.prototype.getCalendar = function (id) {
    return new CalendarFetcher(this.context, this.getPath(id), id);
};

Calendars.prototype.getCalendars = function () {
    return new CalendarCollectionFetcher(this.context, this.path);
};

Calendars.prototype.addCalendar = function (item) {
    var payload = JSON.stringify(Calendar.prototype.preparePayload.call(item));
    return this.executeNativeMethod("addCalendar", Calendar, payload, true);
};

utils.extends(CalendarFetcher, Fetcher);
function CalendarFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

Object.defineProperty(CalendarFetcher.prototype, "events", {
    get: function () {
        if (this._events === undefined) {
            this._events = new Events(this.context, this.getPath("Events"));
        }
        return this._events;
    },
    enumerable: true,
    configurable: true
});

CalendarFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getCalendar", Calendar, this._id);
};

utils.extends(CalendarCollectionFetcher, CollectionFetcher);
function CalendarCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

CalendarCollectionFetcher.prototype.fetch = function (count) {
    return this.fetchAll();
};

CalendarCollectionFetcher.prototype.fetchAll = function () {
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });
    return this.executeNativeMethod("getCalendars", Calendar, queryParams, true);
};

module.exports.Calendar = Calendar;
module.exports.Calendars = Calendars;
module.exports.CalendarFetcher = CalendarFetcher;
module.exports.CalendarCollectionFetcher = CalendarCollectionFetcher;

