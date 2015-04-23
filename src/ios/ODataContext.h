/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 ******************************************************************************/

#import <Foundation/Foundation.h>
#import <office365_exchange_sdk/office365_exchange_sdk.h>
#import <office365_odata_base/office365_odata_base.h>

// Represents OData method execution context and related helper functionality.
@interface ODataContext : NSObject

// Retuns a new instances of ODataContext based on Apache Cordova arguments
+ (ODataContext *)parseCordovaArgs:(NSArray *)args;

// Returns shared instance of MSOutlookClient
- (MSOutlookClient *)outlookClient;

// Applies query params (top, select, etc) to collection fetcher
- (void)applyQueryParams:(MSODataCollectionFetcher *)fetcher;

// Extracts specific path segment from OData url
- (NSString *)extractPathSegment:(long)indexFromTheEnd;

// Extracts entity id from OData url
- (NSString *)extractEntityId;

// Extracts parent entity id from OData url
- (NSString *)extractParentId;

// Extracts container entity id from OData url
- (NSString *)extractConainerId;

// Returns argument passed to native proxy by index
- (NSString *)getArg:(long)index;

// Returns argument passed to native proxy by index as a specific class instance
- (id)getArg:(long)index forType:(Class)type;

@property (nonatomic, copy) NSString *token;
@property (nonatomic, copy) NSString *path;
@property (nonatomic, copy) NSString *serviceRoot;
@property (nonatomic, retain) NSArray *args;

@end