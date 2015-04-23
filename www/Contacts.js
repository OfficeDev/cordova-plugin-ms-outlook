// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var utils = require('./utility');
var Entity = require('./Entity');
var Item = require('./Items').Item;
var Fetchers = require('./Fetchers');
var Fetcher = Fetchers.Fetcher;
var CollectionFetcher = Fetchers.CollectionFetcher;
var PhysicalAddress = require('./ItemHelpers').PhysicalAddress;

utils.extends(Contact, Item);
function Contact(context, path, data) {
    Item.call(this, context, path, data);

    if (!data) {
        return;
    }

    this._id = this.Id = data.Id;

    this.AssistantName = data.AssistantName;
    this.Birthday = data.Birthday ? new Date(data.Birthday) : null;
    this.BusinessAddress = data.BusinessAddress && new PhysicalAddress(data.BusinessAddress);
    this.BusinessHomePage = data.BusinessHomePage;
    this.BusinessPhones = data.BusinessPhones;
    this.CompanyName = data.CompanyName;
    this.Department = data.Department;
    this.DisplayName = data.DisplayName;
    this.EmailAddresses = data.EmailAddresses;
    this.FileAs = data.FileAs;
    this.Generation = data.Generation;
    this.GivenName = data.GivenName;
    this.HomeAddress = data.HomeAddress && new PhysicalAddress(data.HomeAddress);
    this.HomePhones = data.HomePhones;
    this.ImAddresses = data.ImAddresses;
    this.Initials = data.Initials;
    this.JobTitle = data.JobTitle;
    this.Manager = data.Manager;
    this.MiddleName = data.MiddleName;
    this.MobilePhones = data.MobilePhones;
    this.NickName = data.NickName;
    this.OfficeLocation = data.OfficeLocation;
    this.OtherAddress = data.OtherAddress && new PhysicalAddress(data.OtherAddress);
    this.ParentFolderId = data.ParentFolderId;
    this.Profession = data.Profession;
    this.Surname = data.Surname;
    this.Title = data.Title;
    this.YomiCompanyName = data.YomiCompanyName;
    this.YomiGivenName = data.YomiGivenName;
    this.YomiSurname = data.YomiSurname;
    
    // Remove 'attachments' property, set up by Item constructor
    this.attachments = undefined;
}

Contact.prototype.preparePayload = function() {
    var payload = {
        AssistantName: this.AssistantName,
        Birthday: this.Birthday || undefined,
        BusinessAddress: this.BusinessAddress,
        BusinessHomePage: this.BusinessHomePage,
        BusinessPhones: this.BusinessPhones,
        CompanyName: this.CompanyName,
        Department: this.Department,
        DisplayName: this.DisplayName,
        EmailAddresses: this.EmailAddresses,
        FileAs: this.FileAs,
        Generation: this.Generation,
        GivenName: this.GivenName,
        HomeAddress: this.HomeAddress,
        HomePhones: this.HomePhones,
        ImAddresses: this.ImAddresses,
        Initials: this.Initials,
        JobTitle: this.JobTitle,
        Manager: this.Manager,
        MiddleName: this.MiddleName,
        MobilePhones: this.MobilePhones,
        NickName: this.NickName,
        OfficeLocation: this.OfficeLocation,
        OtherAddress: this.OtherAddress,
        ParentFolderId: this.ParentFolderId,
        Profession: this.Profession,
        Surname: this.Surname,
        Title: this.Title,
        YomiCompanyName: this.YomiCompanyName,
        YomiGivenName: this.YomiGivenName,
        YomiSurname: this.YomiSurname
    };
    return payload;
};

Contact.prototype.update = function () {
    var payload = JSON.stringify(this.preparePayload());
    return this.executeNativeMethod("updateContact", Contact, payload);
};

Contact.prototype.delete = function () {
    return this.executeNativeMethod("deleteContact");
};

utils.extends(Contacts, Entity);
function Contacts(context, path) {
    Entity.call(this, context, path);
}

Contacts.prototype.getContact = function (id) {
    return new ContactFetcher(this.context, this.getPath(id), id);
};

Contacts.prototype.getContacts = function () {
    return new ContactCollectionFetcher(this.context, this.path);
};

Contacts.prototype.addContact = function (item) {
    var payload = JSON.stringify(Contact.prototype.preparePayload.call(item));
    return this.executeNativeMethod("addContact", Contact, payload, true);
};

utils.extends(ContactFetcher, Fetcher);
function ContactFetcher(context, path, id) {
    Fetcher.call(this, context, path);
    this._id = id;
}

ContactFetcher.prototype.fetch = function () {
    return this.executeNativeMethod("getContact", Contact, this._id);
};

utils.extends(ContactCollectionFetcher, CollectionFetcher);
function ContactCollectionFetcher (context, path) {
    CollectionFetcher.call(this, context, path);
}

ContactCollectionFetcher.prototype.fetch = function () {
    return this.fetchAll();
};

ContactCollectionFetcher.prototype.fetchAll = function () {
    
    var queryParams = JSON.stringify({
        top: this._top,
        skip: this._skip,
        selectedId: this._selectedId,
        select: this._select,
        expand: this._expand,
        filter: this._filter
    });

    return this.executeNativeMethod("getContacts", Contact, queryParams, true);
};

module.exports.Contact = Contact;
module.exports.Contacts = Contacts;
