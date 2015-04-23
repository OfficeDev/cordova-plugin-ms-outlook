/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 ******************************************************************************/

#import "OutlookServicesProxy.h"
#import "ODataContext.h"

@implementation OutlookServicesProxy

#pragma mark Contacts Api

- (void)addContact:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractParentId];

    NSString *contact = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSOutlookContactCollectionFetcher *fetcher = ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
                                                     ? [[client getMe] getContacts]
                                                     : [[[[client getMe] getContactFolders] getById:parentId] getContacts];

    NSURLSessionTask *task = [fetcher addRaw:contact:^(id result, MSODataException *error) {
        [OutlookServicesProxy passNativeCallResultToJS:result
                                                error:error
                                              delegate:self.commandDelegate
                                           callbackId:command.callbackId];
                                             }];
    [task resume];
}

- (void)getContact:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *contactId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSOutlookContactFetcher *fetcher =
        [[[client getMe] getContacts] getById:contactId];

    NSURLSessionTask *task = [fetcher readRaw:^(id result, MSODataException *error) {
        [OutlookServicesProxy passNativeCallResultToJS:result
                                                error:error
                                             delegate:self.commandDelegate
                                           callbackId:command.callbackId];
    }];
    [task resume];
}

- (void)getContacts:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    NSString *parentId = [ctx extractParentId];

    MSODataCollectionFetcher *fetcher = ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
                                            ? [[client getMe] getContacts]
                                            : [[[[client getMe] getContactFolders] getById:parentId] getContacts];

    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task = [fetcher readRaw:^(id result, MSODataException *error) {
        [OutlookServicesProxy passNativeCallResultToJS:result
                                                error:error
                                             delegate:self.commandDelegate
                                           callbackId:command.callbackId];
    }];
    [task resume];
}

- (void)updateContact:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *contactId = [ctx extractEntityId];

    NSString *contact = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getContacts] getById:contactId];

    NSURLSessionTask *task = [fetcher updateRaw:contact:^(id result, MSODataException *error) {
        [OutlookServicesProxy passNativeCallResultToJS:result
                                                error:error
                                            delegate:self.commandDelegate
                                           callbackId:command.callbackId];
                                                }];
    [task resume];
}

- (void)deleteContact:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *contactId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getContacts] getById:contactId];

    NSURLSessionTask *task = [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                               delegate:self.commandDelegate
                                             callbackId:command.callbackId];
    }];
    [task resume];
}

#pragma mark Contact Folders Api

- (void)addContactFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher = [[client getMe] getContactFolders];

    NSURLSessionTask *task = [fetcher addRaw:
                                        item:^(id result, MSODataException *error) {
            [OutlookServicesProxy passNativeCallResultToJS:result
                                                    error:error
                                                 delegate:self.commandDelegate
                                               callbackId:command.callbackId];
                                        }];
    [task resume];
}

- (void)getContactFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getContactFolders] getById:itemId];

    NSURLSessionTask *task = [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                               delegate:self.commandDelegate
                                             callbackId:command.callbackId];
    }];
    [task resume];
}

- (void)getContactFolders:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher = [[client getMe] getContactFolders];
    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                               delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)updateContactFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        [[[client getMe] getContactFolders] getById:itemId];

    NSURLSessionTask *task = [fetcher updateRaw:
                                           item:^(id result, MSODataException *error) {
                        [OutlookServicesProxy
                            passNativeCallResultToJS:result
                                              error:error
                                           delegate:self.commandDelegate
                                         callbackId:command.callbackId];
                                           }];
    [task resume];
}

- (void)deleteContactFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        [[[client getMe] getContactFolders] getById:itemId];

    NSURLSessionTask *task =
        [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                               delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

#pragma mark Calendard Api

- (void)addCalendar:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractParentId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher =
        ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
            ? [[client getMe] getCalendars]
            : [[[[client getMe] getCalendarGroups]
                  getById:parentId] getCalendars];

    NSURLSessionTask *task = [fetcher
        addRaw:
          item:^(id result, MSODataException *error) {
            [OutlookServicesProxy passNativeCallResultToJS:result
                                                    error:error
                                                 delegate:self.commandDelegate
                                               callbackId:command.callbackId];
          }];
    [task resume];
}

- (void)getCalendar:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        [[[client getMe] getCalendars] getById:itemId];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                               delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)getCalendars:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    NSString *parentId = [ctx extractParentId];

    MSODataCollectionFetcher *fetcher =
        ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
            ? [[client getMe] getCalendars]
            : [[[[client getMe] getCalendarGroups]
                  getById:parentId] getCalendars];

    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                               delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)updateCalendar:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        [[[client getMe] getCalendars] getById:itemId];

    NSURLSessionTask *task =
        [fetcher updateRaw:
                      item:^(id result, MSODataException *error) {
                        [OutlookServicesProxy
                            passNativeCallResultToJS:result
                                              error:error
                                            delegate:self.commandDelegate
                                         callbackId:command.callbackId];
                      }];
    [task resume];
}

- (void)deleteCalendar:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        [[[client getMe] getCalendars] getById:itemId];

    NSURLSessionTask *task =
        [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

#pragma mark Calendar Groups Api

- (void)addCalendarGroup:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher = [[client getMe] getCalendarGroups];

    NSURLSessionTask *task = [fetcher addRaw:item:^(id result, MSODataException *error) {
            [OutlookServicesProxy passNativeCallResultToJS:result
                                                    error:error
                                                  delegate:self.commandDelegate
                                               callbackId:command.callbackId];
                                             }];
    [task resume];
}

- (void)getCalendarGroup:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getCalendarGroups] getById:itemId];

    NSURLSessionTask *task = [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
    }];
    [task resume];
}

- (void)getCalendarGroups:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher = [[client getMe] getCalendarGroups];
    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task = [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
    }];
    [task resume];
}

- (void)updateCalendarGroup:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        [[[client getMe] getCalendarGroups] getById:itemId];

    NSURLSessionTask *task =
        [fetcher updateRaw:
                      item:^(id result, MSODataException *error) {
                        [OutlookServicesProxy
                            passNativeCallResultToJS:result
                                              error:error
                                            delegate:self.commandDelegate
                                         callbackId:command.callbackId];
                      }];
    [task resume];
}

- (void)deleteCalendarGroup:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        [[[client getMe] getCalendarGroups] getById:itemId];

    NSURLSessionTask *task =
        [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

#pragma mark Events Api

- (void)addEvent:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractParentId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher =
        ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
            ? [[client getMe] getEvents]
            : [[[[client getMe] getCalendars] getById:parentId] getEvents];

    NSURLSessionTask *task = [fetcher
        addRaw:
          item:^(id result, MSODataException *error) {
            [OutlookServicesProxy passNativeCallResultToJS:result
                                                    error:error
                                                  delegate:self.commandDelegate
                                               callbackId:command.callbackId];
          }];
    [task resume];
}

- (void)getEvent:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getEvents] getById:itemId];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)getEvents:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    NSString *parentId = [ctx extractParentId];

    MSODataCollectionFetcher *fetcher =
        ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
            ? [[client getMe] getEvents]
            : [[[[client getMe] getCalendars] getById:parentId] getEvents];

    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)updateEvent:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getEvents] getById:itemId];

    NSURLSessionTask *task = [fetcher updateRaw:
                                           item:^(id result, MSODataException *error) {
                        [OutlookServicesProxy
                            passNativeCallResultToJS:result
                                              error:error
                                            delegate:self.commandDelegate
                                         callbackId:command.callbackId];
                                           }];
    [task resume];
}

- (void)deleteEvent:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getEvents] getById:itemId];

    NSURLSessionTask *task =
        [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)accept:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];
    NSString *comment = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSOutlookEventOperations *operations =
        [[[[client getMe] getEvents] getById:itemId] getOperations];

    NSURLSessionTask *task = [operations
         accept:
        comment:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)tentativelyAccept:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];
    NSString *comment = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSOutlookEventOperations *operations =
        [[[[client getMe] getEvents] getById:itemId] getOperations];

    NSURLSessionTask *task = [operations
        tentativelyAccept:
                  comment:^(int result, MSODataException *error) {
                    [OutlookServicesProxy
                        passNativeCallResultToJS:@(result).stringValue
                                          error:error
                                        delegate:self.commandDelegate
                                     callbackId:command.callbackId];
                  }];
    [task resume];
}
- (void)decline:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];
    NSString *comment = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSOutlookEventOperations *operations =
        [[[[client getMe] getEvents] getById:itemId] getOperations];

    NSURLSessionTask *task = [operations
        decline:
        comment:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

#pragma mark Messages Api

- (void)getMessages:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    NSString *parentId = [ctx extractParentId];

    MSODataCollectionFetcher *fetcher =
        ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
            ? [[client getMe] getMessages]
            : [[[[client getMe] getFolders] getById:parentId] getMessages];

    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)getMessage:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];
    MSODataEntityFetcher *fetcher = [[[client getMe] getMessages] getById:itemId];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)addMessage:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractParentId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher =
        ([parentId caseInsensitiveCompare:@"me"] == NSOrderedSame)
            ? [[client getMe] getMessages]
            : [[[[client getMe] getFolders] getById:parentId] getMessages];

    NSURLSessionTask *task = [fetcher
        addRaw:
          item:^(id result, MSODataException *error) {
            [OutlookServicesProxy passNativeCallResultToJS:result
                                                    error:error
                                                  delegate:self.commandDelegate
                                               callbackId:command.callbackId];
          }];
    [task resume];
}
- (void)copyMessage:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];
    NSString *destinationId = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    destinationId = [NSString stringWithFormat:@"'%@'", destinationId];

    NSURLSessionTask *task = [operations
              copyRaw:
        destinationId:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}
- (void)moveMessage:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];
    NSString *destinationId = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    destinationId = [NSString stringWithFormat:@"'%@'", destinationId];

    NSURLSessionTask *task = [operations
              moveRaw:
        destinationId:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}
- (void)updateMessage:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getMessages] getById:itemId];

    NSURLSessionTask *task =
        [fetcher updateRaw:
                      item:^(id result, MSODataException *error) {
                        [OutlookServicesProxy
                            passNativeCallResultToJS:result
                                              error:error
                                            delegate:self.commandDelegate
                                         callbackId:command.callbackId];
                      }];
    [task resume];
}
- (void)deleteMessage:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getMessages] getById:itemId];

    NSURLSessionTask *task =
        [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)createReply:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    NSURLSessionTask *task =
        [operations createReplyRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}
- (void)createReplyAll:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    NSURLSessionTask *task =
        [operations createReplyAllRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}
- (void)createForward:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    NSURLSessionTask *task =
        [operations createForwardRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}
- (void)reply:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];
    NSString *comment = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    NSURLSessionTask *task = [operations
          reply:
        comment:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}
- (void)replyAll:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];
    NSString *comment = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    NSURLSessionTask *task = [operations
        replyAll:
         comment:^(int result, MSODataException *error) {
           [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                   error:error
                                                 delegate:self.commandDelegate
                                              callbackId:command.callbackId];
         }];
    [task resume];
}
- (void)forward:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];
    NSString *comment = [ctx getArg:0];

    comment = [NSString stringWithFormat:@"'%@'", comment];

    NSString *recipients = [ctx getArg:1];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    NSURLSessionTask *task = [operations
        forwardRaw:
           comment:
        recipients:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}
- (void)send:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];
    MSOutlookMessageOperations *operations =
        [[[[client getMe] getMessages] getById:itemId] getOperations];

    NSURLSessionTask *task =
        [operations sendRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

#pragma mark Folders Api

- (void)addFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractParentId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher =
        [[[[client getMe] getFolders] getById:parentId] getChildFolders];

    NSURLSessionTask *task = [fetcher
        addRaw:
          item:^(id result, MSODataException *error) {
            [OutlookServicesProxy passNativeCallResultToJS:result
                                                    error:error
                                                  delegate:self.commandDelegate
                                               callbackId:command.callbackId];
          }];
    [task resume];
}

- (void)getFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getFolders] getById:itemId];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)getFolders:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher = [[client getMe] getFolders];
    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];

        }];
    [task resume];
}

- (void)updateFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getFolders] getById:itemId];

    NSURLSessionTask *task =
        [fetcher updateRaw:
                      item:^(id result, MSODataException *error) {
                        [OutlookServicesProxy
                            passNativeCallResultToJS:result
                                              error:error
                                            delegate:self.commandDelegate
                                         callbackId:command.callbackId];
                      }];
    [task resume];
}

- (void)deleteFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[[client getMe] getFolders] getById:itemId];

    NSURLSessionTask *task =
        [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)copyFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];
    NSString *destinationId = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSOutlookFolderOperations *operations =
        [[[[client getMe] getFolders] getById:itemId] getOperations];

    destinationId = [NSString stringWithFormat:@"'%@'", destinationId];

    NSURLSessionTask *task = [operations
              copyRaw:
        destinationId:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)moveFolder:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];
    NSString *destinationId = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSOutlookFolderOperations *operations =
        [[[[client getMe] getFolders] getById:itemId] getOperations];

    destinationId = [NSString stringWithFormat:@"'%@'", destinationId];

    NSURLSessionTask *task = [operations
              moveRaw:
        destinationId:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

#pragma mark Users Api

- (void)getUser:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher = [[client getUsers] getById:itemId];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)getUsers:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher = [client getUsers];
    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

#pragma mark Attachments Api

- (void)getAttachment:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *itemId = [ctx extractEntityId];
    NSString *parentId = [ctx extractConainerId];
    NSString *parentType = [ctx extractPathSegment:4];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        ([parentType caseInsensitiveCompare:@"messages"] == NSOrderedSame)
            ? [[[[[client getMe] getMessages] getById:parentId] getAttachments]
                  getById:itemId]
            : [[[[[client getMe] getEvents] getById:parentId] getAttachments]
                  getById:itemId];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)getAttachments:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];
    NSString *parentId = [ctx extractParentId];
    NSString *parentType = [ctx extractPathSegment:3];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher =
        ([parentType caseInsensitiveCompare:@"messages"] == NSOrderedSame)
            ? [[[[client getMe] getMessages] getById:parentId] getAttachments]
            : [[[[client getMe] getEvents] getById:parentId] getAttachments];

    [ctx applyQueryParams:fetcher];

    NSURLSessionTask *task =
        [fetcher readRaw:^(id result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:result
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];
    [task resume];
}

- (void)addAttachment:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractParentId];
    NSString *parentType = [ctx extractPathSegment:3];
    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataCollectionFetcher *fetcher =
        ([parentType caseInsensitiveCompare:@"messages"] == NSOrderedSame)
            ? [[[[client getMe] getMessages] getById:parentId] getAttachments]
            : [[[[client getMe] getEvents] getById:parentId] getAttachments];

    NSURLSessionTask *task = [fetcher
        addRaw:
          item:^(id result, MSODataException *error) {
            [OutlookServicesProxy passNativeCallResultToJS:result
                                                    error:error
                                                  delegate:self.commandDelegate
                                               callbackId:command.callbackId];
          }];
    [task resume];
}
- (void)updateAttachment:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractConainerId];
    NSString *itemId = [ctx extractEntityId];
    NSString *parentType = [ctx extractPathSegment:4];
    NSString *item = [ctx getArg:0];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        ([parentType caseInsensitiveCompare:@"messages"] == NSOrderedSame)
            ? [[[[[client getMe] getMessages] getById:parentId] getAttachments]
                  getById:itemId]
            : [[[[[client getMe] getEvents] getById:parentId] getAttachments]
                  getById:itemId];

    NSURLSessionTask *task =
        [fetcher updateRaw:
                      item:^(id result, MSODataException *error) {
                        [OutlookServicesProxy
                            passNativeCallResultToJS:result
                                              error:error
                                            delegate:self.commandDelegate
                                         callbackId:command.callbackId];
                      }];
    [task resume];
}
- (void)deleteAttachment:(CDVInvokedUrlCommand *)command
{
    ODataContext *ctx = [ODataContext parseCordovaArgs:command.arguments];

    NSString *parentId = [ctx extractParentId];
    NSString *itemId = [ctx extractEntityId];
    NSString *parentType = [ctx extractPathSegment:4];

    MSOutlookClient *client = [ctx outlookClient];

    MSODataEntityFetcher *fetcher =
        ([parentType caseInsensitiveCompare:@"messages"] == NSOrderedSame)
            ? [[[[[client getMe] getMessages] getById:parentId] getAttachments]
                  getById:itemId]
            : [[[[[client getMe] getEvents] getById:parentId] getAttachments]
                  getById:itemId];

    NSURLSessionTask *task =
        [fetcher delete:^(int result, MSODataException *error) {
          [OutlookServicesProxy passNativeCallResultToJS:@(result).stringValue
                                                  error:error
                                                delegate:self.commandDelegate
                                             callbackId:command.callbackId];
        }];

    [task resume];
}

#pragma mark Private methods

+ (NSString *)toJsonString:(id)object
{
    JsonParser *parser = [JsonParser new];

    if (![object isKindOfClass:[NSMutableArray class]])
    {
        return [parser toJsonString:object];
    }

    NSMutableString *jsonResult = [[NSMutableString alloc] initWithString:@"["];

    for (int i = 0; i < [object count]; i++)
    {
        if (i != 0)
        {
            [jsonResult appendString:@", "];
        }
        [jsonResult appendString:[parser toJsonString:object[i]]];
    }

    [jsonResult appendString:@"]"];

    return jsonResult;
}

+ (NSString *)serializeError:(MSODataException *)error
{
    NSDictionary *errorDetails = [error.userInfo objectForKey:@"error"];

    // 'error' key stores special error representation which must be returned as a
    // json object
    if (errorDetails)
    {
        NSError *error2;
        NSData *jsonData = [NSJSONSerialization dataWithJSONObject:errorDetails
                                                           options:0
                                                             error:&error2];
        if (!jsonData)
        {
            NSLog(@"Unable to serialize error details: %@", [error2 localizedDescription]);
            return [error localizedDescription];
        }

        return [[NSString alloc] initWithData:jsonData
                                     encoding:NSUTF8StringEncoding];
    }

    // return error description by default
    return [error localizedDescription];
}

+ (void)passNativeCallResultToJS:(id)result
                           error:(MSODataException *)error
                        delegate:(id)delegate
                      callbackId:(NSString *)callbackId
{
    CDVPluginResult *pluginResult;

    if (error != nil)
    {
        pluginResult = [CDVPluginResult
            resultWithStatus:CDVCommandStatus_ERROR
             messageAsString:[OutlookServicesProxy serializeError:error]];
    }
    else
    {
        NSString *payload = ([result isKindOfClass:[NSString class]])
                                ? result
                                : [OutlookServicesProxy toJsonString:result];

        pluginResult = [CDVPluginResult resultWithStatus:CDVCommandStatus_OK
                                         messageAsString:payload];
    }
    [delegate sendPluginResult:pluginResult callbackId:callbackId];
}

@end
