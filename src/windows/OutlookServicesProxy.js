// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved. Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var Windows = window.Windows,
    WinJS = window.WinJS;

//Helper methods

/**
 * Method that generates a string of OData query
 * parameters, such as $filter of $top, from JSON object
 * @param  {String}  jsonParams String representation of query parameters object
 *                              Following fields are common and used oftenly:
 *                                  top: Number
 *                                  skip: Number
 *                                  filter: String
 *                                  select: String
 *                                  expand: String
 * @return {String}             String with prepared query or empty string
 *                                     if no query params provided
 */
function createQueryParams(jsonParams) {

    if (typeof jsonParams === "string") {
        try {
            jsonParams = JSON.parse(jsonParams);
        } catch (e) {
            return "";
        }
    }

    var params = [];
    if (typeof jsonParams === "object") {

        var escape = Windows.Foundation.Uri.escapeComponent;

        // Iterate through all properties provided
        for (var property in jsonParams) {
            if (jsonParams.hasOwnProperty(property)) {
                var propertyValue = jsonParams[property];

                // if property's value is defined and not null or "null"
                // we accumulate its' string representation into array
                if (propertyValue && propertyValue !== "null" && propertyValue !== -1) {
                    var propertyString = "$" + property + "=" + escape(propertyValue);
                    params.push(propertyString);
                }
            }
        }
    }

    // if jsonParams is not a object then
    // we'll just silently continue and return an empty string

    var result = params.length === 0 ? "" : "?" + params.join("&");

    return result;
}

/**
 * Method that creates a base options for XmlHttpRequest object
 * @param  {String} path     Url for accessing to Office REST API
 * @param  {String} token    Access token for Office API, will be added
 *                           to 'Authorization' header of the request
 * @param  {String} httpVerb HTTP verb that vill be used to send a request.
 *                           Possible verbs are "GET", "POST", "PATCH" and "DELETE"
 * @return {Object}          JSON object that can be passed to XHR send() method
 */
function getBaseXhrOptions(path, token, httpVerb) {

    var xhrOptions = {
        url: path,
        type: (httpVerb && typeof httpVerb === "string") ? httpVerb.toUpperCase() : "GET",
        responseType: "json",
        headers: {
            "Accept": "application/json",
            "Authorization": "Bearer " + token,
            "Content-type": "application/json"
        }
    };

    return xhrOptions;
}

/**
 * Method that creates a base options for XmlHttpRequest GET request
 * @param  {String} path     Url for accessing to Office REST API
 * @param  {String} token    Access token for Office API, will be added
 *                           to 'Authorization' header of the request
 * @return {Object}          JSON object that can be passed to XHR send() method
 */
function getXhrOptionsForGet(path, token) {
    var xhrOptions = getBaseXhrOptions(path, token, "GET");
    xhrOptions.headers["If-Modified-Since"] = (new Date(0)).toDateString();

    return xhrOptions;
}

/**
 * Method that creates a base options for XmlHttpRequest POST request
 * @param  {String} path     Url for accessing to Office REST API
 * @param  {String} token    Access token for Office API, will be added
 *                           to 'Authorization' header of the request
 * @return {Object}          JSON object that can be passed to XHR send() method
 */
function getXhrOptionsForPost(path, token, payload) {
    var xhrOptions = getBaseXhrOptions(path, token, "POST");

    if (payload) {
        xhrOptions.data = payload;
    }

    return xhrOptions;
}

/**
 * method that performs an XHR request with provided options
 * @param  {Function}  win               Success callback
 * @param  {Function}  fail              Error callback
 * @param  {Object}  xhrOptions          XHR options object that will be passed to XHR send() method
 */
function doRequest(win, fail, xhrOptions) {
    var xhrUrl = xhrOptions.url;
    WinJS.xhr(xhrOptions).done(
        function completed(result) {
            if (result.readyState === 4) {
                if (result.status >= 200 && result.status <= 299) {
                    var callbackResult = (result.status === 200 || result.status === 201) ? result.response : "";
                    win(callbackResult);
                } else {
                    fail(result.statusText);
                }
            }
        },
        function error(request) {
            try {
                var err = JSON.parse(request.response).error;
                err.details = request.getAllResponseHeaders().split(/\n/g);
                err.url = xhrUrl;
                fail(err);
            } catch (e) {
                fail(request.responseText);
            }
        });
}

/**
 * Base method that performs a POST request to URL specified in args
 * @param  {Function} win                   Success callback
 * @param  {Function} fail                  Error callback
 * @param  {String[]} args                  Array of strings that are passed from
 *                                          common JS layer to JS proxy
 * @param  {String} additionalUriComponent  A string that will be added to OData path
 *                                          before request is performed
 */
function basePostMethod(win, fail, args, additionalUriComponent) {
        try {
            var token = args[0],
                path = additionalUriComponent ? args[2] + '/' + additionalUriComponent : args[2],
                xhrOptions = getXhrOptionsForPost(path, token);

            if (!("data" in xhrOptions)) {
                xhrOptions.data = "";
            }

            doRequest(win, fail, xhrOptions);

        } catch (ex) {
            fail(ex);
        }
}

// API manipulation methods

/**
 * Method that gets a single item by sending a GET request to path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function getItemMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2];

        doRequest(win, fail, getXhrOptionsForGet(path, token));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that gets a collection of items by sending a GET request to path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function getCollectionMethod (win, fail, args) {
    try {
        var token = args[0],
            jsonParams = args[3],
            path = args[2] + createQueryParams(jsonParams);

        doRequest(win, fail, getXhrOptionsForGet(path, token));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that sends a POST request to path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function postMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2],
            newObject = args[3];

        doRequest(win, fail, getXhrOptionsForPost(path, token, newObject));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that accepts an event by sending a POST request to path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function acceptMethod(win, fail, args) {

    try {
        var token = args[0],
            path = args[2] + '/accept',
            data = JSON.stringify({ Comment: args[3] });

        doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that tentatively accepts an event by sending a POST request to path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function tentativelyAcceptMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2] + '/tentativelyaccept',
            data = JSON.stringify({ Comment: args[3] });

        doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that declines an event by sending a POST request to path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function declineMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2] + '/decline',
            data = JSON.stringify({ Comment: args[3] });

        doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that sends a message with path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy * @return {[type]}      [description]
 */
function sendMethod (win, fail, args) {
    basePostMethod(win, fail, args, 'send');
}

/**
 * Method that replies to a message with path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function replyMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2] + '/reply',
            data = JSON.stringify({ Comment: args[3] });

        doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that replies to a message with path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function replyAllMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2] + '/replyall',
            data = JSON.stringify({ Comment: args[3] });

        doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that forwards a message with path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function forwardMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2] + '/forward',
            data = JSON.stringify({
                Comment: args[3],
                ToRecipients: args[4]
            });

        doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that updates an item with path, specified in args by some data
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function patchMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2],
            updatedData = args[3];

        var xhrOptions = getBaseXhrOptions(path, token, "PATCH");
        xhrOptions.data = updatedData;

        doRequest(win, fail, xhrOptions);

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that deletes an item with path, specified in args
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function deleteMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2];

        var xhrOptions = getBaseXhrOptions(path, token, "DELETE");

        doRequest(win, fail, xhrOptions);

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that copies an item with path, specified in args to some destination
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function copyMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2] + '/copy',
            data = JSON.stringify({ DestinationId: args[3] });

        doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that copies an item with path, specified in args to some destination
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function moveMethod (win, fail, args) {
    try {
        var token = args[0],
            path = args[2] + '/move',
            data = JSON.stringify({ DestinationId: args[3] });

            doRequest(win, fail, getXhrOptionsForPost(path, token, data));

    } catch (ex) {
        fail(ex);
    }
}

/**
 * Method that creates a reply to message with path, specified in args to some destination
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function createReplyMethod (win, fail, args) {
    basePostMethod(win, fail, args, 'createreply');
}

/**
 * Method that creates a reply to all to message with path, specified in args to some destination
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function createReplyAllMethod (win, fail, args) {
    basePostMethod(win, fail, args, 'createreplyall');
}

/**
 * Method that creates a forward to message with path, specified in args to some destination
 * @param  {Function} win   Success callback
 * @param  {Function} fail  Error callback
 * @param  {String[]} args  Array of strings that are passed from
 *                          common JS layer to JS proxy
 */
function createForwardMethod (win, fail, args) {
    basePostMethod(win, fail, args, 'createforward');
}

module.exports = {
    addAttachment: postMethod,
    addCalendar: postMethod,
    addCalendarGroup: postMethod,
    addContact: postMethod,
    addContactFolder: postMethod,
    addEvent: postMethod,
    addFolder: postMethod,
    addMessage: postMethod,
    copyFolder: copyMethod,
    copyMessage: copyMethod,
    createForward: createForwardMethod,
    createReply: createReplyMethod,
    createReplyAll: createReplyAllMethod,
    deleteAttachment: deleteMethod,
    deleteCalendar: deleteMethod,
    deleteCalendarGroup: deleteMethod,
    deleteContact: deleteMethod,
    deleteContactFolder: deleteMethod,
    deleteEvent: deleteMethod,
    deleteFolder: deleteMethod,
    deleteMessage: deleteMethod,
    getAttachment: getItemMethod,
    getAttachmentAsFile: getItemMethod,
    getAttachmentAsItem: getItemMethod,
    getAttachmentItem: getItemMethod,
    getAttachments: getCollectionMethod,
    getCalendar: getItemMethod,
    getCalendarGroup: getItemMethod,
    getCalendarGroups: getCollectionMethod,
    getCalendars: getCollectionMethod,
    getContact: getItemMethod,
    getContactFolder: getItemMethod,
    getContactFolders: getCollectionMethod,
    getContacts: getCollectionMethod,
    getEvent: getItemMethod,
    getEvents: getCollectionMethod,
    getFolder: getItemMethod,
    getFolders: getCollectionMethod,
    getMessage: getItemMethod,
    getMessages: getCollectionMethod,
    getUser: getItemMethod,
    getUsers: getCollectionMethod,
    moveFolder: moveMethod,
    moveMessage: moveMethod,
    updateAttachment: patchMethod,
    updateCalendar: patchMethod,
    updateCalendarGroup: patchMethod,
    updateContact: patchMethod,
    updateContactFolder: patchMethod,
    updateEvent: patchMethod,
    updateFolder: patchMethod,
    updateMessage: patchMethod,
    accept: acceptMethod,
    tentativelyAccept: tentativelyAcceptMethod,
    decline: declineMethod,
    forward: forwardMethod,
    reply: replyMethod,
    replyAll: replyAllMethod,
    send: sendMethod
};

require("cordova/exec/proxy").add("OutlookServices", module.exports);
