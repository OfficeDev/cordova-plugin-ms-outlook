// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Item = require('./Items').Item;
var Fetchers = require('./Fetchers');

var Fetcher = Fetchers.Fetcher;
var CollectionFetcher = Fetchers.CollectionFetcher;

var ItemHelpers = require('./ItemHelpers');

var ItemBody = ItemHelpers.ItemBody;
var Attendee = ItemHelpers.Attendee;
var Recipient = ItemHelpers.Recipient;
var Location = ItemHelpers.Location;
var EventType = ItemHelpers.EventType;
var FreeBusyStatus = ItemHelpers.FreeBusyStatus;
var PatternedRecurrence = ItemHelpers.PatternedRecurrence;

utils.extends(Event, Item);
function Event(context, path, data) {
    Item.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id  = this.Id = data.Id;

    this.Attendees = data.Attendees && data.Attendees.map(function (attendee) {
        return new Attendee(attendee);
    });
    this.End = data.End ? new Date(data.End) : null;
    this.IsAllDay = data.IsAllDay;
    this.IsCancelled = data.IsCancelled;
    this.IsOrganizer = data.IsOrganizer;
    this.Location = data.Location && new Location(data.Location);
    this.Organizer = data.Organizer && new Recipient(data.Organizer);
    this.Recurrence = data.Recurrence && new PatternedRecurrence(data.Recurrence);
    this.ResponseRequested = data.ResponseRequested;
    this.SeriesMasterId = data.SeriesMasterId;
    this.SeriesId = data.SeriesId;
    this.ShowAs = FreeBusyStatus[data.ShowAs];
    this.Start = data.Start ? new Date(data.Start) : null;
    this.Type = EventType[data.Type];
}

Event.prototype.preparePayload = function () {
    var payload = {
        Body: this.Body ? ItemBody.prototype.preparePayload.call(this.Body) : undefined,
        Categories: this.Categories,
        Importance: ItemHelpers.Importance[this.Importance],
        Subject: this.Subject,
        Attendees: this.Attendees,
        End: this.End,
        IsAllDay: this.IsAllDay,
        IsCancelled: this.IsCancelled,
        IsOrganizer: this.IsOrganizer,
        Location: this.Location,
        Organizer: this.Organizer || undefined,
        Recurrence: this.Recurrence,
        ResponseRequested: this.ResponseRequested,
        SeriesMasterId: this.SeriesMasterId,
        SeriesId: this.SeriesId,
        ShowAs: ItemHelpers.FreeBusyStatus[this.ShowAs],
        Start: this.Start,
        Type: ItemHelpers.EventType[this.Type]
    };
    return payload;
};

Object.defineProperty(Event.prototype, "calendar", {
    get: function () {
        if (this._calendar === undefined) {
            var CalendarFetcher = require('./Calendars').CalendarFetcher;
            this._calendar = new CalendarFetcher(this.context, this.getPath("Calendar"), "Calendar");
        }
        return this._calendar;
    },
    enumerable: true,
    configurable: true
});

Event.prototype.accept = function (comment) {
    return this.executeNativeMethod("accept", null, comment);
};

Event.prototype.decline = function (comment) {
    return this.executeNativeMethod("decline", null, comment);
};

Event.prototype.tentativelyAccept = function (comment) {
    return this.executeNativeMethod("tentativelyAccept", null, comment);
};

Event.prototype.update = function () {
    var payload = JSON.stringify(this.preparePayload());
    return this.executeNativeMethod("updateEvent", Event, payload);
};

Event.prototype.delete = function () {
    return this.executeNativeMethod("deleteEvent");
};

utils.extends(Events, Entity);
function Events(context, path) {
    Entity.call(this, context, path);
}

Events.prototype.getEvent = function (id) {
    return new EventFetcher(this.context, this.getPath(id), id);
};

Events.prototype.getEvents = function () {
    return new EventCollectionFetcher(this.context, this.path);
};

Events.prototype.addEvent = function (item) {
    var payload = JSON.stringify(Event.prototype.preparePayload.call(item));
    return this.executeNativeMethod("addEvent", Event, payload, true);
};

utils.extends(EventFetcher, Fetcher);
function EventFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

Object.defineProperty(EventFetcher.prototype, "calendar", {
    get: function () {
        if (this._calendar === undefined) {
            var CalendarFetcher = require('./Calendars').CalendarFetcher;
            this._calendar = new CalendarFetcher(this.context, this.getPath("Calendar"), "Calendar");
        }
        return this._calendar;
    },
    enumerable: true,
    configurable: true
});

EventFetcher.prototype.accept = function (comment) {
    return this.executeNativeMethod("accept", null, comment);
};

EventFetcher.prototype.decline = function (comment) {
    return this.executeNativeMethod("decline", null, comment);
};

EventFetcher.prototype.tentativelyAccept = function (comment) {
    return this.executeNativeMethod("tentativelyAccept", null, comment);
};

EventFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getEvent", Event, this._id);
};

utils.extends(EventCollectionFetcher, CollectionFetcher);
function EventCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

EventCollectionFetcher.prototype.fetch = function (count) {
    return this.fetchAll();
};

EventCollectionFetcher.prototype.fetchAll = function () {
    
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });
    
    return this.executeNativeMethod("getEvents", Event, queryParams, true);
};

module.exports.Event = Event;
module.exports.Events = Events;
