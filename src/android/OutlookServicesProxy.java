/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 ******************************************************************************/

package com.msopentech.o365.outlookServices;

import com.microsoft.outlookservices.odata.*;
import com.microsoft.services.odata.impl.*;

import org.apache.cordova.CallbackContext;
import org.apache.cordova.CordovaPlugin;
import org.apache.cordova.PluginResult;

import org.json.JSONArray;
import org.json.JSONException;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * Class that handles calls to Office365 plugin's native layer
 */
public class OutlookServicesProxy extends CordovaPlugin {

    private DefaultDependencyResolver resolver;
    private OutlookClient client;
    private String clientServiceRoot;
    private String clientToken;

    /**
     * Creates a new DefaultDependencyResolver object or return an existing if token is unchanged from previous call
     *
     * @param token access token
     * @return new DefaultDependencyResolver instance or existing one if token is unchanged from previous call
     */
    private DefaultDependencyResolver getResolver (final String token){
        if (this.resolver == null || !token.equals(this.clientToken)){
            this.resolver = new DefaultDependencyResolver(token);
        }
        return this.resolver;
    }

    /**
     * Creates a new OutlookClient object or return an existing if token and serviceRoot are unchanged from previous call
     *
     * @param serviceRoot service root URI
     * @param token access token
     * @return new OutlookClient instance or existing one if token and serviceRoot are unchanged from previous call
     */
    private OutlookClient getClient (final String serviceRoot, final String token) {
        // Check that token and url are the same as in the previous request
        if (this.client == null || !serviceRoot.equals(this.clientServiceRoot) || !token.equals(this.clientToken)) {
            this.client = new OutlookClient(serviceRoot, getResolver(token));
        }
        this.clientServiceRoot = serviceRoot;
        this.clientToken = token;

        return this.client;
    }

    @Override
    public boolean execute(String action, JSONArray args, CallbackContext callbackContext) throws JSONException {

        Method handler;
        try {
            // Get appropriate method for handling provided action
            handler = OutlookServicesMethodsImpl.class.getDeclaredMethod(action, CallbackContext.class, OutlookClient.class,
                    DefaultDependencyResolver.class, ODataMethodArgs.class);
        } catch (NoSuchMethodException e) {
            // If no appropriate method found, send return false to indicate this
            return false;
        }

        // parse arguments passed from JS layer to ArrayList objects
        ODataMethodArgs methodArgs = ODataMethodArgs.parseInvocationArgs(args);

        //Get common parameters necessary for creating OutlookClient object
        //and create a new one or use an existing if token and serviceRoot are unchanged from previous call
        final String token = methodArgs.getToken();
        final String serviceRoot = methodArgs.getServiceRoot();
        OutlookClient client = getClient(serviceRoot, token);

        try {
            // If appropriate method found, invoke it with arguments parsed from action args
            handler.invoke(null, callbackContext, client, this.resolver, methodArgs);
        } catch (InvocationTargetException e) {
            // Catch inner method's exception and send back an error result
            String message = "Method " + action + " failed with error: " + e.getMessage();
            callbackContext.sendPluginResult(new PluginResult(PluginResult.Status.ERROR, message));
        } catch (IllegalAccessException e) {
            // Catch method invocation exception and send back an error result
            String message = "Failed to call native method " + action + ": " + e.getMessage();
            callbackContext.sendPluginResult(new PluginResult(PluginResult.Status.ILLEGAL_ACCESS_EXCEPTION, message));
        }

        // Return true here to indicate that action is handled by plugin
        return true;
    }
}
