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

@class MSOutlookCalendar;

#import <Foundation/Foundation.h>
#import "MSOutlookProtocols.h"
#import "MSOutlookEntity.h"

/**
* The header for type CalendarGroup.
*/

@interface MSOutlookCalendarGroup : MSOutlookEntity


@property NSString *Name;

@property NSString *ChangeKey;

@property NSString *ClassId;

@property NSMutableArray<MSOutlookCalendar> *Calendars;		
		

@end