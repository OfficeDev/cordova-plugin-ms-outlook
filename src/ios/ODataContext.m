/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 ******************************************************************************/

#import "ODataContext.h"

static MSOutlookClient *sharedClient = nil;
static NSString *sharedClientToken = nil;

@implementation ODataContext

+ (ODataContext *)parseCordovaArgs:(NSArray *)args
{
    ODataContext *ctx = [ODataContext new];
    ctx.token = [args objectAtIndex:0];
    ctx.serviceRoot = [args objectAtIndex:1];
    ctx.path = [args objectAtIndex:2];
    ctx.args = args;

    return ctx;
}

- (NSString *)getArg:(long)index
{
    // first 3 items represent common core properties
    return self.args[index + 3];
}

- (id)getArg:(long)index forType:(Class)type
{
    NSString *arg = [self getArg:index];
    JsonParser *parser = [JsonParser new];
    return [parser parseWithData:[arg dataUsingEncoding:NSUTF8StringEncoding] forType:type selector:nil];
}

- (NSString *)extractPathSegment:(long)indexFromTheEnd
{
    NSArray *segments = [self.path componentsSeparatedByString:@"/"];
    return segments[segments.count - indexFromTheEnd];
}

- (NSString *)extractEntityId
{
    return [self extractPathSegment:1];
}
- (NSString *)extractParentId
{
    return [self extractPathSegment:2];
}
- (NSString *)extractConainerId
{
    return [self extractPathSegment:3];
}

- (void)applyQueryParams:(MSODataCollectionFetcher *)fetcher
{
    NSData *jsonData = [[self getArg:0] dataUsingEncoding:NSUTF8StringEncoding];
    NSError *error;
    NSDictionary *dict = [NSJSONSerialization JSONObjectWithData:jsonData
                                                         options:NSJSONReadingAllowFragments
                                                           error:&error];

    if (error != nil)
    {
        NSLog(@"Unable to deserialize query params: %@", [error localizedDescription]);
        return;
    }
    // TODO: review if we can do somehting like this
    // [MSOutlookEntityFetcherHelper setPathForCollections:path :self.UrlComponent :self.top :self.skip  ...];

    int top = (int)[dict[@"top"] integerValue];
    int skip = [dict[@"skip"] integerValue];
    NSString *select = dict[@"select"];
    NSString *expand = dict[@"expand"];
    NSString *filter = dict[@"filter"];

    if (top > -1)
    {
        [fetcher top:top];
    }

    if (skip > -1)
    {
        [fetcher skip:skip];
    }

    if (select && ![select isEqual:[NSNull null]])
    {
        [fetcher select:select];
    }

    if (expand && ![expand isEqual:[NSNull null]])
    {
        [fetcher expand:expand];
    }

    if (filter && ![filter isEqual:[NSNull null]])
    {
        [fetcher filter:filter];
    }
}

- (MSOutlookClient *)outlookClient
{
    @synchronized(self)
    {
        if (sharedClient == nil || sharedClientToken != self.token)
        {
            sharedClientToken = self.token;
            MSODataDefaultDependencyResolver *resolver = [MSODataDefaultDependencyResolver new];
            MSODataOAuthCredentials *credentials = [MSODataOAuthCredentials new];
            [credentials addToken:self.token];

            MSODataCredentialsImpl *credentialsImpl = [MSODataCredentialsImpl new];

            [credentialsImpl setCredentials:credentials];
            [resolver setCredentialsFactory:credentialsImpl];

            sharedClient = [[MSOutlookClient alloc] initWithUrl:self.serviceRoot
                                             dependencyResolver:resolver];
        }
    }
    return sharedClient;
}

@end
