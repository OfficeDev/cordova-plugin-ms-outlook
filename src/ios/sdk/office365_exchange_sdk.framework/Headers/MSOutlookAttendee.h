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

#import "MSOutlookAttendeeType.h"
@class MSOutlookResponseStatus;

#import <Foundation/Foundation.h>
#import "MSOutlookProtocols.h"
#import "MSOutlookRecipient.h"

/**
* The header for type Attendee.
*/

@interface MSOutlookAttendee : MSOutlookRecipient


@property MSOutlookResponseStatus *Status;

@property MSOutlookAttendeeType Type;
-(void)setTypeString:(NSString*)value;

@end