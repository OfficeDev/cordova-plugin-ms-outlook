/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 ******************************************************************************/

package com.msopentech.o365.outlookServices;

import com.google.common.util.concurrent.FutureCallback;
import com.google.common.util.concurrent.Futures;
import com.google.common.util.concurrent.ListenableFuture;

import com.microsoft.outlookservices.*;
import com.microsoft.outlookservices.odata.*;
import com.microsoft.services.odata.*;
import com.microsoft.services.odata.impl.*;

import com.microsoft.services.odata.interfaces.DependencyResolver;
import org.apache.cordova.CallbackContext;

import org.apache.cordova.LOG;
import org.apache.cordova.PluginResult;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.IOException;

/**
 * Helper class that contains static methods for handling plugin's actions
 */
@SuppressWarnings({"UnusedDeclaration"})
class OutlookServicesMethodsImpl {

    /**
     * Tag constant for logging purposes
     */
    private static final String TAG = "Office 365";

    /**
     * Adds default callback that send future's result back to plugin
     *
     * @param future Future to add callback to
     * @param context Plugin context used to send future result back to plugin
     * @param resolver Dependency resolver, that used to serialize future results
     */
    static <T> void addCordovaCallback(final ListenableFuture<T> future, final CallbackContext context, final DependencyResolver resolver){
        Futures.addCallback(future, new FutureCallback<T>() {
            @Override
            public void onSuccess(T t) {
                if (t != null) {
                    String result = resolver.getJsonSerializer().serialize(t);
                    context.sendPluginResult(new PluginResult(PluginResult.Status.OK, result));
                } else {
                    context.sendPluginResult(new PluginResult(PluginResult.Status.OK));
                }
            }

            @Override
            public void onFailure(Throwable throwable) {
                context.sendPluginResult(new PluginResult(PluginResult.Status.ERROR, throwable.getMessage()));
            }
        });
    }

    /**
     * Adds default callback that send future's result back to plugin
     * This is specially for raw SDK methods which returns a string typed future
     *
     * @param future Future to add callback to
     * @param context Plugin context used to send future result back to plugin
     */
    static void addRawCordovaCallback(final ListenableFuture<String> future, final CallbackContext context) {
        Futures.addCallback(future, new FutureCallback<String>() {
            @Override
            public void onSuccess(String s) {
                PluginResult result = s == null ?
                        new PluginResult(PluginResult.Status.OK) :
                        new PluginResult(PluginResult.Status.OK, s);

                context.sendPluginResult(result);
            }

            @Override
            public void onFailure(Throwable throwable) {
                String error = throwable.getMessage();
                if (throwable instanceof ODataException) {
                    try {
                        String response = new String(((ODataException) throwable).getODataResponse().getPayload());
                        // since error object is encapsulated into response's object
                        // try to get it from response and return instead of raw throwable's message
                        JSONObject errorMessage = new JSONObject(response);
                        error = errorMessage.get("error").toString();
                    } catch (JSONException ignored) {
                    } catch (IOException ignored) { }
                }
                context.sendPluginResult(new PluginResult(PluginResult.Status.ERROR, error));
            }
        });
    }

    /**
     * Updates fetcher object with oData query params, specified in queryObject JSON
     *
     * @param fetcher Fetcher object to update
     * @param queryObject JSONObject that contains query parameters:
     *                    top: int,
     *                    skip: int,
     *                    select: String,
     *                    expand: String,
     *                    filter: String
     */
    static void updateFetcherWithQuery (ODataCollectionFetcher fetcher, JSONObject queryObject) {

        try {
            int top = queryObject.getInt("top");
            int skip = queryObject.getInt("skip");
            String select = queryObject.getString("select");
            select = select.equals("null") ? null : select;
            String expand = queryObject.getString("expand");
            expand = expand.equals("null") ? null : expand;
            String filter = queryObject.getString("filter");
            filter = filter.equals("null") ? null : filter;

            if (top > -1) {
                fetcher.top(top);
            }

            if (skip > -1) {
                fetcher.skip(skip);
            }

            if (select != null) {
                fetcher.select(select);
            }

            if (expand != null) {
                fetcher.expand(expand);
            }

            if (filter != null) {
                fetcher.filter(filter);
            }

        } catch (JSONException ignored) {
            LOG.w(TAG, "Failed to parse query parameters", ignored);
            fetcher.reset();
        }
    }

    //region Calendars

    static void getCalendars(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));
        String parentId = methodArgs.parseIdFromODataPath(2);

        ODataCollectionFetcher<Calendar, CalendarFetcher, CalendarCollectionOperations> fetcher = parentId.equalsIgnoreCase("me") ?
                client.getMe().getCalendars() :
                client.getMe().getCalendarGroups().getById(parentId).getCalendars();

        updateFetcherWithQuery(fetcher, queryObject);
        ListenableFuture<String> future = fetcher.readRaw();
        addRawCordovaCallback(future, context);
    }

    static void getCalendar(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String parentId = methodArgs.parseParentIdFromOdataPath();
        String calendarId = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = (parentId.equalsIgnoreCase("me") || parentId.equalsIgnoreCase("calendars")) ?
                client.getMe().getCalendars().getById(calendarId).readRaw() :
                client.getMe().getCalendarGroups().getById(parentId).getCalendars().getById(calendarId).readRaw();
        addRawCordovaCallback(future, context);
    }

    static void addCalendar(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String calendar = methodArgs.getArgs().get(0);
        String parentId = methodArgs.parseIdFromODataPath(2);

        ListenableFuture<String> future = parentId.equalsIgnoreCase("me") ?
                client.getMe().getCalendars().addRaw(calendar) :
                client.getMe().getCalendarGroups().getById(parentId).getCalendars().addRaw(calendar);

        addRawCordovaCallback(future, context);
    }

    static void updateCalendar(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String calendarId = methodArgs.parseIdFromODataPath();
        String calendar = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getCalendars().getById(calendarId).updateRaw(calendar);
        addRawCordovaCallback(future, context);
    }

    static void deleteCalendar(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs){

        String calendarId = methodArgs.parseIdFromODataPath();

        ListenableFuture future = client.getMe().getCalendars().getById(calendarId).delete();
        addCordovaCallback(future, context, resolver);
    }

    //endregion

    //region Calendar Groups

    static void getCalendarGroups(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));

        ODataCollectionFetcher<CalendarGroup, CalendarGroupFetcher, CalendarGroupCollectionOperations> fetcher = client.getMe().getCalendarGroups();
        updateFetcherWithQuery(fetcher, queryObject);

        ListenableFuture<String> future = fetcher.readRaw();
        addRawCordovaCallback(future, context);
    }

    static void getCalendarGroup(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String calendarGroupId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getCalendarGroups().getById(calendarGroupId).readRaw();
        addRawCordovaCallback(future, context);
    }

    static void addCalendarGroup(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String calendarGroup = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getCalendarGroups().addRaw(calendarGroup);
        addRawCordovaCallback(future, context);
    }

    static void updateCalendarGroup(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String calendarGroupId = methodArgs.parseIdFromODataPath();
        String calendarGroup = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getCalendarGroups().getById(calendarGroupId).updateRaw(calendarGroup);
        addRawCordovaCallback(future, context);
    }

    static void deleteCalendarGroup(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs){

        String calendarGroupId = methodArgs.parseIdFromODataPath();

        ListenableFuture future = client.getMe().getCalendarGroups().getById(calendarGroupId).delete();
        addCordovaCallback(future, context, resolver);
    }
    //endregion

    //region Contacts

    static void deleteContact(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String contactId = methodArgs.parseIdFromODataPath();
        ListenableFuture future = client.getMe().getContacts().getById(contactId).delete();

        addCordovaCallback(future, context, resolver);
    }

    static void updateContact(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String contactId = methodArgs.parseIdFromODataPath();
        String contact = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getContacts().getById(contactId).updateRaw(contact);
        addRawCordovaCallback(future, context);
    }

    static void addContact(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String parentId = methodArgs.parseParentIdFromOdataPath();
        String contact = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = parentId.equalsIgnoreCase("me") ?
                client.getMe().getContacts().addRaw(contact) :
                client.getMe().getContactFolders().getById(parentId).getContacts().addRaw(contact);

        addRawCordovaCallback(future, context);
    }

    static void getContact(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String contactId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getContacts().getById(contactId).readRaw();
        addRawCordovaCallback(future, context);
    }

    static void getContacts(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        String parentId = methodArgs.parseParentIdFromOdataPath();
        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));

        ODataCollectionFetcher<Contact, ContactFetcher, ContactCollectionOperations> fetcher = parentId.equalsIgnoreCase("me") ?
                client.getMe().getContacts() :
                client.getMe().getContactFolders().getById(parentId).getContacts();

        updateFetcherWithQuery(fetcher, queryObject);
        addRawCordovaCallback(fetcher.readRaw(), context);
    }
    //endregion

    //region Events

    static void getEvent(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String eventId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getEvents().getById(eventId).readRaw();
        addRawCordovaCallback(future, context);
    }

    static void getEvents(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        String parentId = methodArgs.parseParentIdFromOdataPath();
        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));

        ODataCollectionFetcher<Event, EventFetcher, EventCollectionOperations> fetcher = parentId.equalsIgnoreCase("me") ?
                client.getMe().getEvents() :
                client.getMe().getCalendars().getById(parentId).getEvents();

        updateFetcherWithQuery(fetcher, queryObject);
        ListenableFuture<String> future = fetcher.readRaw();

        addRawCordovaCallback(future, context);
    }

    static void addEvent(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String parentId = methodArgs.parseParentIdFromOdataPath();
        String event = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = parentId.equalsIgnoreCase("me") ?
                client.getMe().getEvents().addRaw(event) :
                client.getMe().getCalendars().getById(parentId).getEvents().addRaw(event);

        addRawCordovaCallback(future, context);
    }

    static void updateEvent(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String eventId = methodArgs.parseIdFromODataPath();
        String updatedEvent = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getEvents().getById(eventId).updateRaw(updatedEvent);
        addRawCordovaCallback(future, context);
    }

    static void deleteEvent(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String eventToDeleteId = methodArgs.parseIdFromODataPath();

        ListenableFuture future = client.getMe().getEvents().getById(eventToDeleteId).delete();
        addCordovaCallback(future, context, resolver);
    }

    static void accept(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String eventToAcceptId = methodArgs.parseIdFromODataPath();
        String comment = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getEvents().getById(eventToAcceptId).getOperations().acceptRaw(comment);
        addRawCordovaCallback(future, context);
    }

    static void tentativelyAccept(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String eventToAcceptId = methodArgs.parseIdFromODataPath();
        String comment = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getEvents().getById(eventToAcceptId).getOperations().tentativelyAcceptRaw(comment);
        addRawCordovaCallback(future, context);
    }

    static void decline(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String eventId = methodArgs.parseIdFromODataPath();
        String comment = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getEvents().getById(eventId).getOperations().declineRaw(comment);
        addRawCordovaCallback(future, context);
    }
    //endregion

    //region Folders

    static void getFolders(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        String parentFolderId = methodArgs.parseParentIdFromOdataPath();
        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));

        ODataCollectionFetcher<Folder, FolderFetcher, FolderCollectionOperations> fetcher = client.getMe().getFolders().getById(parentFolderId).getChildFolders();
        updateFetcherWithQuery(fetcher, queryObject);
        ListenableFuture<String> future = fetcher.readRaw();

        addRawCordovaCallback(future, context);
    }

    static void getFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String folderId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getFolders().getById(folderId).readRaw();
        addRawCordovaCallback(future, context);
    }

    static void addFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String parentFolderId = methodArgs.parseParentIdFromOdataPath();
        String folder = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getFolders().getById(parentFolderId).getChildFolders().addRaw(folder);
        addRawCordovaCallback(future, context);
    }

    static void copyFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String folderId = methodArgs.parseIdFromODataPath();
        String destinationId = methodArgs.getArgs().get(0);

        ListenableFuture<String> future =
                client.getMe().getFolders().getById(folderId).getOperations().copyRaw("\"" + destinationId + "\"");

        addRawCordovaCallback(future, context);
    }

    static void moveFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String folderId = methodArgs.parseIdFromODataPath();
        String destinationId = methodArgs.getArgs().get(0);

        ListenableFuture<String> future =
                client.getMe().getFolders().getById(folderId).getOperations().moveRaw("\"" + destinationId + "\"");

        addRawCordovaCallback(future, context);
    }

    static void updateFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String folderId = methodArgs.parseIdFromODataPath();
        String folder = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getFolders().getById(folderId).updateRaw(folder);
        addRawCordovaCallback(future, context);
    }

    static void deleteFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String folderId = methodArgs.parseIdFromODataPath();

        ListenableFuture future = client.getMe().getFolders().getById(folderId).delete();
        addCordovaCallback(future, context, resolver);
    }
    //endregion

    //region Messages

    static void getMessages(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        String parentId = methodArgs.parseParentIdFromOdataPath();
        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));

        ODataCollectionFetcher<Message, MessageFetcher, MessageCollectionOperations> fetcher = parentId.equalsIgnoreCase("me") ?
                client.getMe().getMessages() :
                client.getMe().getFolders().getById(parentId).getMessages();

        updateFetcherWithQuery(fetcher, queryObject);
        ListenableFuture<String> messages = fetcher.readRaw();

        addRawCordovaCallback(messages, context);
    }

    static void getMessage(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> message = client.getMe().getMessages().getById(messageId).readRaw();
        addRawCordovaCallback(message, context);
    }

    static void addMessage(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String parentId = methodArgs.parseParentIdFromOdataPath();
        String message = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = parentId.equalsIgnoreCase("me") ?
                client.getMe().getMessages().addRaw(message) :
                client.getMe().getFolders().getById(parentId).getMessages().addRaw(message);

        addRawCordovaCallback(future, context);
    }

    static void copyMessage(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageId = methodArgs.parseIdFromODataPath();
        String destinationId = methodArgs.getArgs().get(0);

        ListenableFuture<String> future =
                client.getMe().getMessages().getById(messageId).getOperations().copyRaw("\"" + destinationId + "\"");

         addRawCordovaCallback(future, context);
    }

    static void moveMessage(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageToMoveId = methodArgs.parseIdFromODataPath();
        String destinationId = methodArgs.getArgs().get(0);

        ListenableFuture<String> future =
                client.getMe().getMessages().getById(messageToMoveId).getOperations().moveRaw("\"" + destinationId + "\"");

        addRawCordovaCallback(future, context);
    }

    static void updateMessage(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String messageId = methodArgs.parseIdFromODataPath();
        String message = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getMessages().getById(messageId).updateRaw(message);
        addRawCordovaCallback(future, context);
    }

    static void deleteMessage(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageToDeleteId = methodArgs.parseIdFromODataPath();

        ListenableFuture future = client.getMe().getMessages().getById(messageToDeleteId).delete();
        addCordovaCallback(future, context, resolver);
    }

    static void createReply(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getMessages().getById(messageId).getOperations().createReplyRaw();
        addRawCordovaCallback(future, context);
    }

    static void createReplyAll(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getMessages().getById(messageId).getOperations().createReplyAllRaw();
        addRawCordovaCallback(future, context);
    }

    static void createForward(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getMessages().getById(messageId).getOperations().createForwardRaw();
        addRawCordovaCallback(future, context);
    }

    static void reply(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageToReplyId = methodArgs.parseIdFromODataPath();
        String comment = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getMessages().getById(messageToReplyId).getOperations().replyRaw(comment);
        addRawCordovaCallback(future, context);
    }

    static void replyAll(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageToReplyId = methodArgs.parseIdFromODataPath();
        String comment = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getMessages().getById(messageToReplyId).getOperations().replyAllRaw(comment);
        addRawCordovaCallback(future, context);
    }

    static void forward(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String messageId = methodArgs.parseIdFromODataPath();
        String comment = methodArgs.getArgs().get(0);
        String recipients = methodArgs.getArgs().get(1);

        ListenableFuture<String> future =
                client.getMe().getMessages().getById(messageId).getOperations().forwardRaw(comment, recipients);

        addRawCordovaCallback(future, context);
    }

    static void send(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String messageToSendId = methodArgs.parseIdFromODataPath();

        ListenableFuture<String> future = client.getMe().getMessages().getById(messageToSendId).getOperations().sendRaw();
        addRawCordovaCallback(future, context);
    }
    //endregion

    //region Users

    static void getUsers(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));

        ODataCollectionFetcher<User, UserFetcher, UserCollectionOperations> fetcher = client.getUsers();
        updateFetcherWithQuery(fetcher, queryObject);
        ListenableFuture<String> future = fetcher.readRaw();

        addRawCordovaCallback(future, context);
    }

    static void getUser(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String userId = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getUsers().getById(userId).readRaw();
        addRawCordovaCallback(future, context);
    }

    //endregion

    //region Attachments

    static void getAttachments(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));
        String parentId = methodArgs.parseParentIdFromOdataPath();
        String parentType = methodArgs.parseParentTypeFromOdataPath(3);

        ODataCollectionFetcher<Attachment, AttachmentFetcher, AttachmentCollectionOperations> fetcher;
        if (parentType.equals("messages")){
            fetcher = client.getMe().getMessages().getById(parentId).getAttachments();
        } else {
            fetcher = client.getMe().getEvents().getById(parentId).getAttachments();
        }

        updateFetcherWithQuery(fetcher, queryObject);

        ListenableFuture<String> future = fetcher.readRaw();
        addRawCordovaCallback(future, context);
    }

    static void getAttachment(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String attachmentId = methodArgs.parseIdFromODataPath();
        String parentId = methodArgs.parseContainerIdFromODataPath();
        String parentType = methodArgs.parseParentTypeFromOdataPath(4);

        ODataCollectionFetcher<Attachment, AttachmentFetcher, AttachmentCollectionOperations> fetcher;
        if (parentType.equals("messages")){
            fetcher = client.getMe().getMessages().getById(parentId).getAttachments();
        } else {
            fetcher = client.getMe().getEvents().getById(parentId).getAttachments();
        }

        ListenableFuture<String> future = fetcher.getById(attachmentId).readRaw();
        addRawCordovaCallback(future, context);
    }

    static void getAttachmentItem(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String attachmentId = methodArgs.parseIdFromODataPath(2);
        String parentId = methodArgs.parseIdFromODataPath(4);
        String parentType = methodArgs.parseParentTypeFromOdataPath(5);

        ODataCollectionFetcher<Attachment, AttachmentFetcher, AttachmentCollectionOperations> fetcher = parentType.equals("messages") ?
                client.getMe().getMessages().getById(parentId).getAttachments() :
                client.getMe().getEvents().getById(parentId).getAttachments();

        ListenableFuture<String> future = fetcher.getById(attachmentId).asItemAttachment().getItem().readRaw();
        addRawCordovaCallback(future, context);
    }

    static void addAttachment(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String attachment = methodArgs.getArgs().get(0);
        String parentId = methodArgs.parseParentIdFromOdataPath();
        String parentType = methodArgs.parseParentTypeFromOdataPath(3);

        ODataCollectionFetcher<Attachment, AttachmentFetcher, AttachmentCollectionOperations> fetcher = parentType.equals("messages") ?
                client.getMe().getMessages().getById(parentId).getAttachments() :
                client.getMe().getEvents().getById(parentId).getAttachments();

        ListenableFuture<String> future = fetcher.addRaw(attachment);
        addRawCordovaCallback(future, context);
    }

    static void updateAttachment(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String attachment = methodArgs.getArgs().get(0);
        String attachmentId = methodArgs.parseIdFromODataPath();
        String parentId = methodArgs.parseParentIdFromOdataPath();
        String parentType = methodArgs.parseParentTypeFromOdataPath(4);

        ODataCollectionFetcher<Attachment, AttachmentFetcher, AttachmentCollectionOperations> fetcher = parentType.equals("messages") ?
                client.getMe().getMessages().getById(parentId).getAttachments() :
                client.getMe().getEvents().getById(parentId).getAttachments();

        ListenableFuture<String> future = fetcher.getById(attachmentId).updateRaw(attachment);
        addRawCordovaCallback(future, context);
    }

    static void deleteAttachment(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String attachmentId = methodArgs.parseIdFromODataPath();
        String parentId = methodArgs.parseContainerIdFromODataPath();
        String parentType = methodArgs.parseParentTypeFromOdataPath(4);

        ODataCollectionFetcher<Attachment, AttachmentFetcher, AttachmentCollectionOperations> fetcher = parentType.equals("messages") ?
                client.getMe().getMessages().getById(parentId).getAttachments() :
                client.getMe().getEvents().getById(parentId).getAttachments();

        ListenableFuture future = fetcher.getById(attachmentId).delete();
        addCordovaCallback(future, context, resolver);
    }

    //endregion

    //region ContactFolders

    static void getContactFolders(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws JSONException {

        JSONObject queryObject = new JSONObject(methodArgs.getArgs().get(0));
        String parentFolderId = methodArgs.parseParentIdFromOdataPath();

        ODataCollectionFetcher<ContactFolder, ContactFolderFetcher, ContactFolderCollectionOperations> fetcher = parentFolderId.equalsIgnoreCase("me") ?
                client.getMe().getContactFolders() :
                client.getMe().getContactFolders().getById(parentFolderId).getChildFolders();

        updateFetcherWithQuery(fetcher, queryObject);

        ListenableFuture<String> future = fetcher.readRaw();
        addRawCordovaCallback(future, context);
    }

    static void getContactFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String folderId = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getContactFolders().getById(folderId).readRaw();
        addRawCordovaCallback(future, context);
    }

    static void addContactFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String parentFolderId = methodArgs.parseParentIdFromOdataPath();
        String folder = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getContactFolders().getById(parentFolderId).getChildFolders().addRaw(folder);
        addRawCordovaCallback(future, context);
    }

    static void updateContactFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) throws Throwable {

        String folderId = methodArgs.parseIdFromODataPath();
        String folder = methodArgs.getArgs().get(0);

        ListenableFuture<String> future = client.getMe().getContactFolders().getById(folderId).updateRaw(folder);
        addRawCordovaCallback(future, context);
    }

    static void deleteContactFolder(final CallbackContext context, final OutlookClient client, final DefaultDependencyResolver resolver, final ODataMethodArgs methodArgs) {

        String folderId = methodArgs.parseIdFromODataPath();

        ListenableFuture future = client.getMe().getContactFolders().getById(folderId).delete();
        addCordovaCallback(future, context, resolver);
    }

    //endregion
}
