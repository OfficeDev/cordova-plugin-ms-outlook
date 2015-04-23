// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');

function ItemBody(data) {
    if (!data) {
        return;
    }

    this.ContentType = BodyType[data.ContentType];
    this.Content = data.Content;
}

ItemBody.prototype.preparePayload = function() {
    var contentType = this.ContentType in BodyType ?
        typeof this.ContentType === "number" ? BodyType[this.ContentType] : this.ContentType :
        undefined;

    return {
        ContentType: contentType,
        Content: this.Content
    };
};

function Recipient(data) {
    if (!data) {
        return;
    }
    
    this.EmailAddress = {
        Name: data.EmailAddress.Name,
        Address: data.EmailAddress.Address
    };
}

Recipient.prototype.preparePayload = function () {
    return {
        ContentType: BodyType[this.ContentType],
        Content: this.Content
    };
};

utils.extends(Attendee, Recipient);
function Attendee (data) {
    Recipient.call(this, data);
    
    if (!data) {
        return;
    }

    this.Status = new ResponseStatus(data.Status);
    this.Type = AttendeeType[data.Type];
}

utils.extends(ResponseStatus, Recipient);
function ResponseStatus(data) {

    if (!data) {
        return;
    }

    this.Response = ResponseType[data.Response];
    this.Time = data.Time ? new Date(data.Time) : null;
}

function Location(data) {
    if (!data) {
        return;
    }

    this.DisplayName = data.DisplayName;
}

function PatternedRecurrence(data) {
    if (!data) {
        return;
    }

    this.Pattern = new RecurrencePattern(data.Pattern);
    this.Range = new RecurrenceRange(data.Range);
}

function RecurrencePattern(data) {
    if (!data) {
        return;
    }

    this.Type = RecurrencePatternType[data.Type];
    this.Interval = data.Interval;
    this.DayOfMonth = data.DayOfMonth;
    this.Month = data.Month;
    this.DaysOfWeek = data.DaysOfWeek;
    this.FirstDayOfWeek = DayOfWeek[data.FirstDayOfWeek];
    this.Index = WeekIndex[data.Index];
}

function RecurrenceRange(data) {
    if (!data) {
        return;
    }

    this.Type = RecurrenceRangeType[data.Type];
    this.StartDate = data.StartDate ? new Date(data.StartDate) : null;
    this.EndDate = data.EndDate ? new Date(data.EndDate) : null;
    this.NumberOfOccurrences = data.NumberOfOccurrences;
}

function PhysicalAddress(data) {
    if (!data) {
        return;
    }

    this.Street = data.Street;
    this.City = data.City;
    this.State = data.State;
    this.CountryOrRegion = data.CountryOrRegion;
    this.PostalCode = data.PostalCode;
}

var AttendeeType = {};
AttendeeType[AttendeeType["Required"] = 0] = "Required";
AttendeeType[AttendeeType["Optional"] = 1] = "Optional";
AttendeeType[AttendeeType["Resource"] = 2] = "Resource";

var BodyType = {};
BodyType[BodyType["Text"] = 0] = "Text";
BodyType[BodyType["HTML"] = 1] = "HTML";

var MeetingMessageType = {};
MeetingMessageType[MeetingMessageType["None"] = 0] = "None";
MeetingMessageType[MeetingMessageType["MeetingRequest"] = 1] = "MeetingRequest";
MeetingMessageType[MeetingMessageType["MeetingCancelled"] = 2] = "MeetingCancelled";
MeetingMessageType[MeetingMessageType["MeetingAccepted"] = 3] = "MeetingAccepted";
MeetingMessageType[MeetingMessageType["MeetingTenativelyAccepted"] = 4] = "MeetingTenativelyAccepted";
MeetingMessageType[MeetingMessageType["MeetingDeclined"] = 5] = "MeetingDeclined";

var Importance = {};
Importance[Importance["Normal"] = 0] = "Normal";
Importance[Importance["Low"] = 1] = "Low";
Importance[Importance["High"] = 2] = "High";

var ResponseType = {};
ResponseType[ResponseType["None"] = 0] = "None";
ResponseType[ResponseType["Organizer"] = 1] = "Organizer";
ResponseType[ResponseType["TentativelyAccepted"] = 2] = "TentativelyAccepted";
ResponseType[ResponseType["Accepted"] = 3] = "Accepted";
ResponseType[ResponseType["Declined"] = 4] = "Declined";
ResponseType[ResponseType["NotResponded"] = 5] = "NotResponded";

var RecurrencePatternType = {};
RecurrencePatternType[RecurrencePatternType["Daily"] = 0] = "Daily";
RecurrencePatternType[RecurrencePatternType["Weekly"] = 1] = "Weekly";
RecurrencePatternType[RecurrencePatternType["AbsoluteMonthly"] = 2] = "AbsoluteMonthly";
RecurrencePatternType[RecurrencePatternType["RelativeMonthly"] = 3] = "RelativeMonthly";
RecurrencePatternType[RecurrencePatternType["AbsoluteYearly"] = 4] = "AbsoluteYearly";
RecurrencePatternType[RecurrencePatternType["RelativeYearly"] = 5] = "RelativeYearly";

var RecurrenceRangeType = {};
RecurrenceRangeType[RecurrenceRangeType["EndDate"] = 0] = "EndDate";
RecurrenceRangeType[RecurrenceRangeType["NoEnd"] = 1] = "NoEnd";
RecurrenceRangeType[RecurrenceRangeType["Numbered"] = 2] = "Numbered";

var DayOfWeek = {};
DayOfWeek[DayOfWeek["Sunday"] = 0] = "Sunday";
DayOfWeek[DayOfWeek["Monday"] = 1] = "Monday";
DayOfWeek[DayOfWeek["Tuesday"] = 2] = "Tuesday";
DayOfWeek[DayOfWeek["Wednesday"] = 3] = "Wednesday";
DayOfWeek[DayOfWeek["Thursday"] = 4] = "Thursday";
DayOfWeek[DayOfWeek["Friday"] = 5] = "Friday";
DayOfWeek[DayOfWeek["Saturday"] = 6] = "Saturday";

var WeekIndex = {};
WeekIndex[WeekIndex["First"] = 0] = "First";
WeekIndex[WeekIndex["Second"] = 1] = "Second";
WeekIndex[WeekIndex["Third"] = 2] = "Third";
WeekIndex[WeekIndex["Fourth"] = 3] = "Fourth";
WeekIndex[WeekIndex["Last"] = 4] = "Last";

var EventType = {};
EventType[EventType["SingleInstance"] = 0] = "SingleInstance";
EventType[EventType["Occurrence"] = 1] = "Occurrence";
EventType[EventType["Exception"] = 2] = "Exception";
EventType[EventType["SeriesMaster"] = 3] = "SeriesMaster";

var FreeBusyStatus = {};
FreeBusyStatus[FreeBusyStatus["Unknown"] = 0] = "Unknown";
FreeBusyStatus[FreeBusyStatus["Free"] = 1] = "Free";
FreeBusyStatus[FreeBusyStatus["Tentative"] = 2] = "Tentative";
FreeBusyStatus[FreeBusyStatus["Busy"] = 3] = "Busy";
FreeBusyStatus[FreeBusyStatus["Oof"] = 4] = "Oof";
FreeBusyStatus[FreeBusyStatus["WorkingElsewhere"] = 5] = "WorkingElsewhere";


module.exports.Attendee = Attendee;
module.exports.AttendeeType = AttendeeType;
module.exports.ResponseType = ResponseType;
module.exports.Location = Location;
module.exports.EventType = EventType;
module.exports.FreeBusyStatus = FreeBusyStatus;
module.exports.PhysicalAddress = PhysicalAddress;
module.exports.PatternedRecurrence = PatternedRecurrence;
module.exports.RecurrencePattern = RecurrencePattern;
module.exports.RecurrencePatternType = RecurrencePatternType;
module.exports.RecurrenceRange = RecurrenceRange;
module.exports.RecurrenceRangeType = RecurrenceRangeType;
module.exports.DayOfWeek = DayOfWeek;
module.exports.WeekIndex = WeekIndex;
module.exports.BodyType = BodyType;
module.exports.ItemBody = ItemBody;
module.exports.Importance = Importance;
module.exports.MeetingMessageType = MeetingMessageType;
module.exports.Recipient = Recipient;


