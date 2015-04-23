/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 *
 * Warning: This code was generated automatically. Edits will be overwritten.
 * To make changes to this code, please make changes to the generation framework itself:
 * https://github.com/MSOpenTech/odata-codegen
 *******************************************************************************/

#import <office365_odata_base/office365_odata_base.h>
#import "MSOutlookFolder.h"
#import "MSOutlookEntityOperations.h"

#import "MSOutlookFolder.h"
#import "MSOutlookMessage.h"

/**
* The header for type MSOutlookFolderOperations.
*/

@interface MSOutlookFolderOperations : MSOutlookEntityOperations

-(id)initWithUrl:(NSString *)urlComponent parent:(id<MSODataExecutable>)parent;

-(NSURLSessionDataTask*)copy : (NSString *) destinationId : (void (^)(MSOutlookFolder *folder, MSODataException *error))callback;			
-(NSURLSessionDataTask*)copyRaw : (NSString*) destinationId : (void (^)(NSString*folder, MSODataException *error))callback;			

-(NSURLSessionDataTask*)move : (NSString *) destinationId : (void (^)(MSOutlookFolder *folder, MSODataException *error))callback;			
-(NSURLSessionDataTask*)moveRaw : (NSString*) destinationId : (void (^)(NSString*folder, MSODataException *error))callback;			

@end