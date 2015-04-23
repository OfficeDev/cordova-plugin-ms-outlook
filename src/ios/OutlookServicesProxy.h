/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 ******************************************************************************/

#import <Foundation/Foundation.h>
#import <Cordova/CDVPlugin.h>

#import <office365_exchange_sdk/office365_exchange_sdk.h>

// Implements Apache Cordova plugin for Office365 Outlook Services
@interface OutlookServicesProxy : CDVPlugin

// Contacts
- (void)addContact:(CDVInvokedUrlCommand *)command;
- (void)getContact:(CDVInvokedUrlCommand *)command;
- (void)getContacts:(CDVInvokedUrlCommand *)command;
- (void)updateContact:(CDVInvokedUrlCommand *)command;
- (void)deleteContact:(CDVInvokedUrlCommand *)command;

// Contact Folders
- (void)addContactFolder:(CDVInvokedUrlCommand *)command;
- (void)getContactFolder:(CDVInvokedUrlCommand *)command;
- (void)getContactFolders:(CDVInvokedUrlCommand *)command;
- (void)updateContactFolder:(CDVInvokedUrlCommand *)command;
- (void)deleteContactFolder:(CDVInvokedUrlCommand *)command;

// Calendars
- (void)addCalendar:(CDVInvokedUrlCommand *)command;
- (void)getCalendar:(CDVInvokedUrlCommand *)command;
- (void)getCalendars:(CDVInvokedUrlCommand *)command;
- (void)updateCalendar:(CDVInvokedUrlCommand *)command;
- (void)deleteCalendar:(CDVInvokedUrlCommand *)command;

// Calendar Groups
- (void)addCalendarGroup:(CDVInvokedUrlCommand *)command;
- (void)getCalendarGroup:(CDVInvokedUrlCommand *)command;
- (void)getCalendarGroups:(CDVInvokedUrlCommand *)command;
- (void)updateCalendarGroup:(CDVInvokedUrlCommand *)command;
- (void)deleteCalendarGroup:(CDVInvokedUrlCommand *)command;

// Events
- (void)addEvent:(CDVInvokedUrlCommand *)command;
- (void)getEvent:(CDVInvokedUrlCommand *)command;
- (void)updateEvent:(CDVInvokedUrlCommand *)command;
- (void)deleteEvent:(CDVInvokedUrlCommand *)command;
- (void)accept:(CDVInvokedUrlCommand *)command;
- (void)tentativelyAccept:(CDVInvokedUrlCommand *)command;
- (void)decline:(CDVInvokedUrlCommand *)command;

// Messages
- (void)getMessages:(CDVInvokedUrlCommand *)command;
- (void)getMessage:(CDVInvokedUrlCommand *)command;
- (void)addMessage:(CDVInvokedUrlCommand *)command;
- (void)copyMessage:(CDVInvokedUrlCommand *)command;
- (void)moveMessage:(CDVInvokedUrlCommand *)command;
- (void)updateMessage:(CDVInvokedUrlCommand *)command;
- (void)deleteMessage:(CDVInvokedUrlCommand *)command;
- (void)createReply:(CDVInvokedUrlCommand *)command;
- (void)createReplyAll:(CDVInvokedUrlCommand *)command;
- (void)createForward:(CDVInvokedUrlCommand *)command;
- (void)reply:(CDVInvokedUrlCommand *)command;
- (void)replyAll:(CDVInvokedUrlCommand *)command;
- (void)forward:(CDVInvokedUrlCommand *)command; // TODO not implemented
- (void)send:(CDVInvokedUrlCommand *)command;

// Folders
- (void)addFolder:(CDVInvokedUrlCommand *)command;
- (void)getFolder:(CDVInvokedUrlCommand *)command;
- (void)getFolders:(CDVInvokedUrlCommand *)command;
- (void)copyFolder:(CDVInvokedUrlCommand *)command;
- (void)moveFolder:(CDVInvokedUrlCommand *)command;
- (void)updateFolder:(CDVInvokedUrlCommand *)command;
- (void)deleteFolder:(CDVInvokedUrlCommand *)command;

// Users
- (void)getUsers:(CDVInvokedUrlCommand *)command;
- (void)getUser:(CDVInvokedUrlCommand *)command;

// Attachments
- (void)getAttachments:(CDVInvokedUrlCommand *)command;
- (void)getAttachment:(CDVInvokedUrlCommand *)command;
- (void)addAttachment:(CDVInvokedUrlCommand *)command;
- (void)updateAttachment:(CDVInvokedUrlCommand *)command;
- (void)deleteAttachment:(CDVInvokedUrlCommand *)command;
@end