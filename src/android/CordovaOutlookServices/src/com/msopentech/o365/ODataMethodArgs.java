/*******************************************************************************
 * Copyright (c) Microsoft Open Technologies, Inc.
 * All Rights Reserved
 * Licensed under the Apache License, Version 2.0.
 * See License.txt in the project root for license information.
 ******************************************************************************/

package com.msopentech.o365.outlookServices;

import org.json.JSONArray;
import org.json.JSONException;

import java.util.ArrayList;
import java.util.List;

/**
 * Class that represents arguments passed from common JS layer
 */
class ODataMethodArgs {

    /**
     * OAuth access token
     */
    private final String token;

    /**
     * Office365 service root
     */
    private final String serviceRoot;

    /**
     * Entity OData path (includes service root URL)
     */
    private final String oDataPath;

    /**
     * Method arguments, specific for each method
     */
    private final List<String> args;

    private ODataMethodArgs(String token, String serviceRoot, String oDataPath, List<String> args) {
        this.token = token;
        this.serviceRoot = serviceRoot;
        this.oDataPath = oDataPath;
        this.args = args;
    }

    /**
     * Parses Arguments JSON passed to execute method from JS layer to ODataMethodArgs object
     * that contains necessary arguments for all methods: token, service root and oData path
     *
     * @param actionArgs JSONArray of arguments, passed from common JS layer
     * @return ArrayList that contains token, service root and oData path
     * @throws JSONException
     */
    static ODataMethodArgs parseInvocationArgs(JSONArray actionArgs) throws JSONException {
        String token = actionArgs.getString(0);
        String serviceRoot = actionArgs.getString(1);
        String oDataPath = actionArgs.getString(2);

        ArrayList<String> args = new ArrayList<String>();

        for (int i = 3; i < actionArgs.length(); i++) {
            args.add(actionArgs.getString(i));
        }

        return new ODataMethodArgs(token, serviceRoot, oDataPath, args);
    }

    /**
     * @return token
     */
    public String getToken() {
        return this.token;
    }

    /**
     * @return service root
     */
    public String getServiceRoot() {
        return this.serviceRoot;
    }

    /**
     * @return OData path
     */
    public String getODataPath() {
        return this.oDataPath;
    }

    /**
     * @return Method's specific arguments
     */
    public List<String> getArgs() {
        return this.args;
    }

    /**
     * @param indexFromTheEnd index of Id parameter in OData path (from the end of path)
     * @return Id
     * @throws IndexOutOfBoundsException
     */
    public String parseIdFromODataPath(int indexFromTheEnd) throws IndexOutOfBoundsException {
        String[] oDataPathParts = this.oDataPath.split("/");
        if (indexFromTheEnd > oDataPathParts.length){
            throw new IndexOutOfBoundsException();
        }
        return oDataPathParts[oDataPathParts.length - indexFromTheEnd];
    }

    /**
     * @return Id (first from the end)
     * @throws IndexOutOfBoundsException
     */
    public String parseIdFromODataPath(){
        return this.parseIdFromODataPath(1);
    }

    /**
     * Used to get parent entity's Id from OData path (for nested entities)
     * @return Entity parent's Id (second from the end)
     * @throws IndexOutOfBoundsException
     */
    public String parseParentIdFromOdataPath() {
        return this.parseIdFromODataPath(2);
    }

    /**
     * Used to get parent container's Id from OData path (for Attachments)
     * @return Entity parent's Id (third from the end)
     * @throws IndexOutOfBoundsException
     */
    public String parseContainerIdFromODataPath() {
        return this.parseIdFromODataPath(3);
    }

    /**
     * Used to get parent container's type from OData path (for Attachments)
     * @param indexFromTheEnd index of container's type parameter in OData path (from the end of path)
     * @return Entity container's type ("message" or "event")
     * @throws Throwable
     */
    public String parseParentTypeFromOdataPath(int indexFromTheEnd) throws Throwable {
        String type = this.parseIdFromODataPath(indexFromTheEnd).toLowerCase();
        if (type.equals("events") || type.equals("messages")){
            return type;
        }
        throw new Throwable("Can't parse parent container type");
    }
}
