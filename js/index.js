// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var AuthenticationContext;

var authority = 'https://login.windows.net/keckmedicine.onmicrosoft.com';
var resourceUrl = 'https://graph.windows.net/';
var appId = '2b12cfe7-b4d8-4256-9072-ca27dade4e55';
var redirectUrl = 'http://localhost:4400/services/aad/redirectTarget.html';

var tenantName = 'keckmedicine.onmicrosoft.com';
var endpointUrl = resourceUrl + tenantName;

function pre(json) {
    return '<pre>' + JSON.stringify(json, null, 4) + '</pre>';
}

var app = {
    // Application Constructor
    initialize: function () {
        this.bindEvents();
    },
    // Bind Event Listeners
    //
    // Bind any events that are required on startup. Common events are:
    // 'load', 'deviceready', 'offline', and 'online'.
    bindEvents: function () {
        alert("bindEvents");
        document.addEventListener('deviceready', app.onDeviceReady, false);
        
    },
    // deviceready Event Handler
    //
    // The scope of 'this' is the event. In order to call the 'receivedEvent'
    // function, we must explicitly call 'app.receivedEvent(...);'
    onDeviceReady: function () {
        alert("onDeviceReady");
        // app.receivedEvent('deviceready');
        document.getElementById('search').addEventListener('click', app.search);
        app.logArea = document.getElementById("log-area");
       // app.log("Cordova initialized, 'deviceready' event was fired");
        AuthenticationContext = Microsoft.ADAL.AuthenticationContext;
    },
    // Implements search operations.
    search: function () {
        document.getElementById('userlist').innerHTML = "";
        alert("Search");
        app.authenticate(function (authresult) {               
            var searchText = document.getElementById('searchfield').value;
            alert("searchText = " + searchText);
            app.requestData(authresult, searchText);
        });
    },
    // Update DOM on a Received Event
    receivedEvent: function (id) {
        var parentElement = document.getElementById(id);
        var listeningElement = parentElement.querySelector('.listening');
        var receivedElement = parentElement.querySelector('.received');

        listeningElement.setAttribute('style', 'display:none;');
        receivedElement.setAttribute('style', 'display:block;');

        console.log('Received Event: ' + id);
    },

    log: function (message, isError) {
        isError ? console.error(message) : console.log(message);
        var logItem = document.createElement('li');
        logItem.classList.add("topcoat-list__item");
        isError && logItem.classList.add("error-item");
        var timestamp = '<span class="timestamp">' + new Date().toLocaleTimeString() + ': </span>';
        logItem.innerHTML = (timestamp + message);
        app.logArea.insertBefore(logItem, app.logArea.firstChild);
    },
    error: function (message) {
        alert(message);
        app.log(message, true);
    },
    createContext: function () {
        alert("createContext");  
        AuthenticationContext.createAsync(authority)
        .then(function (context) {
            app.authContext = context;
            alert("Created authentication context for authority URL: " + context.authority);
      //      app.log("Created authentication context for authority URL: " + context.authority);
        }, app.error);
    },
    acquireToken: function () {
        if (app.authContext == null) {
   //         app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }

        app.authContext.acquireTokenAsync(resourceUrl, appId, redirectUrl)
            .then(function(authResult) {
     //           app.log('Acquired token successfully: ' + pre(authResult));
            }, function(err) {
    //            app.error("Failed to acquire token: " + pre(err));
            });
    },
    acquireTokenSilent: function () {
        alert("acquireTokenSilent");
        if (app.authContext == null) {
            alert("Authentication context isn\'t created yet. Create context first");
     //       app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }

        // testUserId parameter is needed if you have > 1 token cache items to avoid "multiple_matching_tokens_detected" error
        // Note: This is for the test purposes only
        var testUserId;
        app.authContext.tokenCache.readItems().then(function (cacheItems) {
            if (cacheItems.length > 0) {
                alert("cacheItems.length");
                testUserId = cacheItems[0].userInfo.userId;
            }

            app.authContext.acquireTokenSilentAsync(resourceUrl, appId, testUserId).then(function (authResult) {
                alert("Acquired token successfull");
      //          app.log('Acquired token successfully: ' + pre(authResult));
            }, function (err) {
                alert("Failed to acquire token silently");
        //        app.error("Failed to acquire token silently: " + pre(err));
            });
        }, function (err) {
            app.acquireToken();
       //     app.error("Unable to get User ID from token cache. Have you acquired token already? " + pre(err));
        });
    },
    readTokenCache: function () {
        if (app.authContext == null) {
     //       app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }

        app.authContext.tokenCache.readItems()
        .then(function (res) {
            var text = "Read token cache successfully. There is " + res.length + " items stored.";
            if (res.length > 0) {
                text += "The first one is: " + pre(res[0]);
            }
      //      app.log(text);

        }, function (err) {
         //   app.error("Failed to read token cache: " + pre(err));
        });
    },
    clearTokenCache: function () {
        if (app.authContext == null) {
          //  app.error('Authentication context isn\'t created yet. Create context first');
            return;
        }

        app.authContext.tokenCache.clear().then(function () {
            app.log("Cache cleaned up successfully.");
        }, function (err) {
           // app.error("Failed to clear token cache: " + pre(err));
        });
    },
    // Shows user authentication dialog if required.
    authenticate: function (authCompletedCallback) {
        alert("authenticate");
        app.createContext();
        app.acquireTokenSilent();
       
        app.context = new Microsoft.ADAL.AuthenticationContext(authority);
        app.context.tokenCache.readItems().then(function (items) {
            if (items.length > 0) {
                authority = items[0].authority;
                app.context = new Microsoft.ADAL.AuthenticationContext(authority);
            }
            // Attempt to authorize user silently
            app.context.acquireTokenSilentAsync(resourceUri, clientId)
            .then(authCompletedCallback, function () {
                // We require user cridentials so triggers authentication dialog
                app.context.acquireTokenAsync(resourceUri, clientId, redirectUri)
                .then(authCompletedCallback, function (err) {
                    app.error("Failed to authenticate: " + err);
                });
            });
        });
        
    },
    // Makes Api call to receive user list.
    requestData: function (authResult, searchText) {
        var req = new XMLHttpRequest();
        var url = resourceUri + "/" + authResult.tenantId + "/users?api-version=" + graphApiVersion;
        url = searchText ? url + "&$filter=mailNickname eq '" + searchText + "'" : url + "&$top=10";

        req.open("GET", url, true);
        req.setRequestHeader('Authorization', 'Bearer ' + authResult.accessToken);

        req.onload = function(e) {
            if (e.target.status >= 200 && e.target.status < 300) {
                app.renderData(JSON.parse(e.target.response));
                return;
            }
            app.error('Data request failed: ' + e.target.response);
        };
        req.onerror = function(e) {
            app.error('Data request failed: ' + e.error);
        }

        req.send();
    },
    // Renders user list.
    renderData: function(data) {
        var users = data && data.value;
        if (users.length === 0) {
            app.error("No users found");
            return;
        }

        var userlist = document.getElementById('userlist');
        userlist.innerHTML = "";

        // Helper function for generating HTML
        function $new(eltName, classlist, innerText, children, attributes) {
            var elt = document.createElement(eltName);
            classlist.forEach(function (className) {
                elt.classList.add(className);
            });

            if (innerText) {
                elt.innerText = innerText;
            }

            if (children && children.constructor === Array) {
                children.forEach(function (child) {
                    elt.appendChild(child);
                });
            } else if (children instanceof HTMLElement) {
                elt.appendChild(children);
            }

            if(attributes && attributes.constructor === Object) {
                for(var attrName in attributes) {
                    elt.setAttribute(attrName, attributes[attrName]);
                }
            }

            return elt;
        }

        users.map(function(userInfo) {
            return $new('li', ['topcoat-list__item'], null, [
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'First name: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.givenName || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'Last name: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.surname || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'UPN: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.userPrincipalName || ''
                    })
                ]),
                $new('div', [], null, [
                    $new('p', ['userinfo-label'], 'Phone: '),
                    $new('input', ['topcoat-text-input', 'userinfo-data-field'], null, null, {
                        type: 'text',
                        readonly: '',
                        placeholder: '',
                        value: userInfo.telephoneNumber || ''
                    })
                ])
            ]);
        }).forEach(function(userListItem) {
            userlist.appendChild(userListItem);
        });
    },
    // Renders application error.
    error: function(err) {
        var userlist = document.getElementById('userlist');
        userlist.innerHTML = "";

        var errorItem = document.createElement('li');
        errorItem.classList.add('topcoat-list__item');
        errorItem.classList.add('error-item');
        errorItem.innerText = err;

        userlist.appendChild(errorItem);
    }
};

document.addEventListener('deviceready', app.onDeviceReady, false);
