//! Copyright (c) Microsoft Corporation. All rights reserved.
var __extends = this && this.__extends ||
function(e, t) {
    function r() {
        this.constructor = e
    }
    for (var i in t) t.hasOwnProperty(i) && (e[i] = t[i]);
    r.prototype = t.prototype;
    e.prototype = new r
}; !
function(e) {
    if ("object" == typeof exports && "undefined" != typeof module) module.exports = e();
    else if ("function" == typeof define && define.amd) define([], e);
    else {
        var t;
        t = "undefined" != typeof window ? window: "undefined" != typeof global ? global: "undefined" != typeof self ? self: this;
        t.OneDrive = e()
    }
} (function() {
    return function e(t, r, i) {
        function n(a, s) {
            if (!r[a]) {
                if (!t[a]) {
                    var l = "function" == typeof require && require;
                    if (!s && l) return l(a, !0);
                    if (o) return o(a, !0);
                    var c = new Error("Cannot find module '" + a + "'");
                    throw c.code = "MODULE_NOT_FOUND",
                    c
                }
                var d = r[a] = {
                    exports: {}
                };
                t[a][0].call(d.exports,
                function(e) {
                    var r = t[a][1][e];
                    return n(r ? r: e)
                },
                d, d.exports, e, t, r, i)
            }
            return r[a].exports
        }
        for (var o = "function" == typeof require && require,
        a = 0; a < i.length; a++) n(i[a]);
        return n
    } ({
        1 : [function(e, t, r) {
            var i = e("./models/ErrorType"),
            n = function() {
                function e() {}
                e.ERROR_ACCESS_DENIED = "access_denied";
                e.ERROR_AUTH_REQUIRED = "user_authentication_required";
                e.ERROR_LOGIN_REQUIRED = "login_required";
                e.ERROR_POPUP_OPEN = {
                    errorCode: i.popupOpen,
                    message: "popup window is already open"
                };
                e.ERROR_WEB_REQUEST = {
                    errorCode: i.webRequestFailure,
                    message: "web request failed, see console logs for details"
                };
                e.PROTOCOL_HTTP = "HTTP";
                e.PROTOCOL_HTTPS = "HTTPS";
                e.HTTP_GET = "GET";
                e.HTTP_POST = "POST";
                e.HTTP_PUT = "PUT";
                e.LINKTYPE_WEB = "webLink";
                e.LINKTYPE_DOWNLOAD = "downloadLink";
                e.PARAM_ACCESS_TOKEN = "access_token";
                e.PARAM_ERROR = "error";
                e.PARAM_STATE = "state";
                e.PARAM_SDK_STATE = "sdk_state";
                e.PARAM_ID_TOKEN = "id_token";
                e.SDK_VERSION = "js-v2.1";
                e.STATE_AAD_LOGIN = "aad_login";
                e.STATE_AAD_PICKER = "aad_picker";
                e.STATE_MSA_PICKER = "msa_picker";
                e.STATE_OPEN_POPUP = "open_popup";
                e.STATE_GRAPH = "graph";
                e.TYPE_BOOLEAN = "boolean";
                e.TYPE_FUNCTION = "function";
                e.TYPE_OBJECT = "object";
                e.TYPE_STRING = "string";
                e.TYPE_NUMBER = "number";
                e.VROOM_URL = "https://api.onedrive.com/v1.0/";
                e.VROOM_INT_URL = "https://newapi.storage.live-int.com/v1.0/";
                e.GRAPH_URL = "https://graph.microsoft.com/v1.0/";
                e.NONCE_LENGTH = 5;
                e.CUSTOMER_TID = "9188040d-6c67-4c5b-b112-36a304b66dad";
                e.DEFAULT_QUERY_ITEM_PARAMETER = "expand=thumbnails&select=id,name,size,webUrl,folder";
                e.GLOBAL_FUNCTION_PREFIX = "oneDriveFilePicker";
                e.LOGIN_HINT_KEY = "odsdkLoginHint";
                e.LOGIN_HINT_LIFE_SPAN = 864e5;
                e.DOMAIN_HINT_COMMON = "common";
                e.DOMAIN_HINT_AAD = "organizations";
                e.DOMAIN_HINT_MSA = "consumers";
                return e
            } ();
            t.exports = n
        },
        {
            "./models/ErrorType": 6
        }],
        2 : [function(e, t, r) {
            var i = e("./Constants"),
            n = e("./OneDriveApp"),
            o = function() {
                function e() {}
                e.open = function(e) {
                    n.open(e)
                };
                e.save = function(e) {
                    n.save(e)
                };
                e.webLink = i.LINKTYPE_WEB;
                e.downloadLink = i.LINKTYPE_DOWNLOAD;
                return e
            } ();
            n.onloadInit();
            t.exports = o
        },
        {
            "./Constants": 1,
            "./OneDriveApp": 3
        }],
        3 : [function(e, t, r) {
            var i = e("./utilities/DomUtilities"),
            n = e("./utilities/ErrorHandler"),
            o = e("./utilities/Logging"),
            a = e("./OneDriveState"),
            s = e("./utilities/Picker"),
            l = e("./models/PickerMode"),
            c = e("./utilities/RedirectUtilities"),
            d = e("./utilities/ResponseParser"),
            u = e("./utilities/Saver"),
            p = function() {
                function e() {}
                e.onloadInit = function() {
                    n.registerErrorObserver(a.clearState);
                    i.getScriptInput();
                    o.logMessage("initialized");
                    var e = c.handleRedirect();
                    if (e) {
                        var t = d.parsePickerResponse(e),
                        r = e.windowState.options,
                        p = e.windowState.optionsMode;
                        r.clientId ? a.clientId = r.clientId: n.throwError("client id is missing in options");
                        switch (p) {
                        case l[l.open]:
                            var f = new s(r);
                            t.error ? f.handlePickerError(t) : f.handlePickerSuccess(t);
                            break;
                        case l[l.save]:
                            var h = new u(r);
                            t.error ? h.handleSaverError(t) : h.handleSaverSuccess(t);
                            break;
                        default:
                            n.throwError("invalid value for options.mode: " + p)
                        }
                    }
                };
                e.open = function(e) {
                    if (a.readyCheck()) {
                        e || n.throwError("missing picker options");
                        o.logMessage("open started");
                        var t = new s(e);
                        t.launchPicker()
                    }
                };
                e.save = function(e) {
                    if (a.readyCheck()) {
                        e || n.throwError("missing saver options");
                        o.logMessage("save started");
                        var t = new u(e);
                        t.launchSaver()
                    }
                };
                return e
            } ();
            t.exports = p
        },
        {
            "./OneDriveState": 4,
            "./models/PickerMode": 9,
            "./utilities/DomUtilities": 15,
            "./utilities/ErrorHandler": 16,
            "./utilities/Logging": 18,
            "./utilities/Picker": 20,
            "./utilities/RedirectUtilities": 22,
            "./utilities/ResponseParser": 23,
            "./utilities/Saver": 24
        }],
        4 : [function(e, t, r) {
            var i = e("./Constants"),
            n = e("./utilities/Logging"),
            o = function() {
                function e() {}
                e.clearState = function() {
                    window.name = "";
                    e._isSdkReady = !0
                };
                e.readyCheck = function() {
                    if (!e._isSdkReady) return ! 1;
                    e._isSdkReady = !1;
                    return ! 0
                };
                e.getODCHost = function() {
                    return (e.debug ? "live-int": "live") + ".com"
                };
                e.getLoginHint = function(t) {
                    if (localStorage) {
                        var r = JSON.parse(localStorage.getItem(i.LOGIN_HINT_KEY));
                        null != r && r[t] && (e.loginHint = r[t])
                    } else n.logError("the browser does not support local storage")
                };
                e.updateLoginHint = function(t) {
                    if (localStorage) {
                        var r = JSON.parse(localStorage.getItem(i.LOGIN_HINT_KEY));
                        null == r && (r = {});
                        r[t] = e.loginHint;
                        localStorage.setItem(i.LOGIN_HINT_KEY, JSON.stringify(r))
                    } else n.logError("the browser does not support local storage")
                };
                e.deleteLoghinHint = function(e) {
                    if (localStorage) {
                        var t = JSON.parse(localStorage.getItem(i.LOGIN_HINT_KEY));
                        if (null != t && t[e]) {
                            delete t[e];
                            localStorage.setItem(i.LOGIN_HINT_KEY, JSON.stringify(t))
                        }
                    } else n.logError("the browser does not support local storage")
                };
                e.debug = !1;
                e._isSdkReady = !0;
                return e
            } ();
            t.exports = o
        },
        {
            "./Constants": 1,
            "./utilities/Logging": 18
        }],
        5 : [function(e, t, r) {
            "use strict";
            var i; !
            function(e) {
                e[e.filesV2 = 0] = "filesV2";
                e[e.graph_odc = 1] = "graph_odc";
                e[e.graph_odb = 2] = "graph_odb";
                e[e.other = 3] = "other"
            } (i || (i = {}));
            t.exports = i
        },
        {}],
        6 : [function(e, t, r) {
            "use strict";
            var i; !
            function(e) {
                e[e.badResponse = 0] = "badResponse";
                e[e.fileReaderFailure = 1] = "fileReaderFailure";
                e[e.popupOpen = 2] = "popupOpen";
                e[e.unknown = 3] = "unknown";
                e[e.unsupportedFeature = 4] = "unsupportedFeature";
                e[e.webRequestFailure = 5] = "webRequestFailure"
            } (i || (i = {}));
            t.exports = i
        },
        {}],
        7 : [function(e, t, r) {
            var i = e("../utilities/CallbackInvoker"),
            n = e("../Constants"),
            o = e("../utilities/ErrorHandler"),
            a = e("../utilities/Logging"),
            s = e("../OneDriveState"),
            l = e("../utilities/StringUtilities"),
            c = e("../utilities/TypeValidators"),
            d = e("../utilities/UrlUtilities"),
            u = new RegExp("^[a-fA-F\\d]{8}-([a-fA-F\\d]{4}-){3}[a-fA-F\\d]{12}$"),
            p = function() {
                function e(t) {
                    this.openInNewWindow = c.validateType(t.openInNewWindow, n.TYPE_BOOLEAN, !0, !0);
                    this.expectGlobalFunction = !this.openInNewWindow;
                    if (this.expectGlobalFunction) {
                        this.cancelName = t.cancel;
                        this.errorName = t.error
                    }
                    var r = c.validateCallback(t.cancel, !0, this.expectGlobalFunction);
                    this.cancel = function() {
                        a.logMessage("user cancelled operation");
                        i.invokeAppCallback(r, !0)
                    };
                    var o = c.validateCallback(t.error, !0, this.expectGlobalFunction);
                    this.error = function(e) {
                        a.logError(l.format("error occured - code: '{0}', message: '{1}'", e.errorCode, e.message));
                        i.invokeAppCallback(o, !0, e)
                    };
                    this.advanced = c.validateType(t.advanced, n.TYPE_OBJECT, !0, {
                        redirectUri: d.trimUrlQuery(window.location.href)
                    });
                    this.advanced.redirectUri || (this.advanced.redirectUri = d.trimUrlQuery(window.location.href));
                    this.advanced.sharePointDoclibUrl && (this.advanced.sharePointDoclibUrl = e.setProtectedInfo(this.advanced.sharePointDoclibUrl));
                    this.clientId = c.validateType(t.clientId, n.TYPE_STRING);
                    this.isSharePointRedirect = !!this.advanced.sharePointDoclibUrl;
                    e.checkClientId(this.clientId)
                }
                e.getProtectedInfo = function(e) {
                    return e ? window.localStorage.getItem(e) : void 0
                };
                e.setProtectedInfo = function(e) {
                    var t = d.generateNonce();
                    window.localStorage.setItem(t, e);
                    return t
                };
                e.checkClientId = function(e) {
                    if (e) {
                        u.test(e) ? a.logMessage("parsed AAD client id: " + e) : o.throwError(l.format("invalid format for client id '{0}' - AAD: 32 characters (HEX) GUID", e));
                        s.clientId = e
                    } else o.throwError("client id is missing in options")
                };
                return e
            } ();
            t.exports = p
        },
        {
            "../Constants": 1,
            "../OneDriveState": 4,
            "../utilities/CallbackInvoker": 14,
            "../utilities/ErrorHandler": 16,
            "../utilities/Logging": 18,
            "../utilities/StringUtilities": 25,
            "../utilities/TypeValidators": 26,
            "../utilities/UrlUtilities": 27
        }],
        8 : [function(e, t, r) {
            "use strict";
            var i; !
            function(e) {
                e[e.download = 0] = "download";
                e[e.query = 1] = "query";
                e[e.share = 2] = "share"
            } (i || (i = {}));
            t.exports = i
        },
        {}],
        9 : [function(e, t, r) {
            "use strict";
            var i; !
            function(e) {
                e[e.open = 0] = "open";
                e[e.save = 1] = "save"
            } (i || (i = {}));
            t.exports = i
        },
        {}],
        10 : [function(e, t, r) {
            var i = e("./PickerActionType"),
            n = e("../utilities/CallbackInvoker"),
            o = e("../Constants"),
            a = e("./InvokerOptions"),
            s = e("../utilities/Logging"),
            l = e("./PickerMode"),
            c = e("../utilities/TypeValidators"),
            d = function(e) {
                function t(t) {
                    e.call(this, t);
                    this.expectGlobalFunction && (this.successName = t.success);
                    var r = c.validateCallback(t.success, !1, this.expectGlobalFunction);
                    this.success = function(e) {
                        s.logMessage("picker operation succeeded");
                        n.invokeAppCallback(r, !0, e)
                    };
                    this.multiSelect = c.validateType(t.multiSelect, o.TYPE_BOOLEAN, !0, !1);
                    var a = c.validateType(t.action, o.TYPE_STRING);
                    this.action = i[a]
                }
                __extends(t, e);
                t.prototype.isSharing = function() {
                    return this.action === i.share
                };
                t.prototype.serializeState = function() {
                    return {
                        optionsMode: l[l.open],
                        options: {
                            action: i[this.action],
                            advanced: this.advanced,
                            clientId: this.clientId,
                            isSharePointRedirect: this.isSharePointRedirect,
                            success: this.successName,
                            cancel: this.cancelName,
                            error: this.errorName,
                            multiSelect: this.multiSelect,
                            openInNewWindow: this.openInNewWindow
                        }
                    }
                };
                return t
            } (a);
            t.exports = d
        },
        {
            "../Constants": 1,
            "../utilities/CallbackInvoker": 14,
            "../utilities/Logging": 18,
            "../utilities/TypeValidators": 26,
            "./InvokerOptions": 7,
            "./PickerActionType": 8,
            "./PickerMode": 9
        }],
        11 : [function(e, t, r) {
            "use strict";
            var i; !
            function(e) {
                e[e.save = 0] = "save";
                e[e.query = 1] = "query"
            } (i || (i = {}));
            t.exports = i
        },
        {}],
        12 : [function(e, t, r) {
            var i = e("../utilities/CallbackInvoker"),
            n = e("../Constants"),
            o = e("../utilities/DomUtilities"),
            a = e("../utilities/ErrorHandler"),
            s = e("./InvokerOptions"),
            l = e("../utilities/Logging"),
            c = e("./PickerMode"),
            d = e("./SaverActionType"),
            u = e("../utilities/StringUtilities"),
            p = e("../utilities/TypeValidators"),
            f = e("./UploadType"),
            h = e("../utilities/UrlUtilities"),
            g = 104857600,
            _ = "100 MB",
            v = function(e) {
                function t(t) {
                    e.call(this, t);
                    this.invalidFile = !1;
                    if (this.expectGlobalFunction) {
                        this.successName = t.success;
                        this.progressName = t.progress
                    }
                    var r = p.validateCallback(t.success, !1, this.expectGlobalFunction);
                    this.success = function(e) {
                        l.logMessage("saver operation succeeded");
                        i.invokeAppCallback(r, !0, e)
                    };
                    var o = p.validateCallback(t.progress, !0, this.expectGlobalFunction);
                    this.progress = function(e) {
                        l.logMessage(u.format("upload progress: {0}%", e));
                        i.invokeAppCallback(o, !1, e)
                    };
                    var a = p.validateType(t.action, n.TYPE_STRING, !0, "query");
                    this.action = d[a];
                    this.action === d.save && this._setFileInfo(t)
                }
                __extends(t, e);
                t.prototype.serializeState = function() {
                    return {
                        optionsMode: c[c.save],
                        options: {
                            action: d[this.action],
                            advanced: this.advanced,
                            cancel: this.cancelName,
                            clientId: this.clientId,
                            error: this.errorName,
                            fileName: this.fileName,
                            isSharePointRedirect: this.isSharePointRedirect,
                            openInNewWindow: this.openInNewWindow,
                            progress: this.progressName,
                            sourceInputElementId: this.sourceInputElementId,
                            sourceUri: this.sourceUri,
                            success: this.successName
                        }
                    }
                };
                t.prototype._setFileInfo = function(e) {
                    e.sourceInputElementId && e.sourceUri && a.throwError("Only one type of file to save.");
                    this.sourceInputElementId = e.sourceInputElementId;
                    this.sourceUri = e.sourceUri;
                    var t = p.validateType(e.fileName, n.TYPE_STRING, !0, null);
                    if (this.sourceUri) {
                        if (h.isPathFullUrl(this.sourceUri)) {
                            this.uploadType = f.url;
                            this.fileName = t || h.getFileNameFromUrl(this.sourceUri);
                            this.fileName || a.throwError("must supply a file name or a URL that ends with a file name")
                        } else if (h.isPathDataUrl(this.sourceUri)) {
                            this.uploadType = f.dataUrl;
                            this.fileName = t;
                            this.fileName || a.throwError("must supply a file name for data URL uploads")
                        }
                    } else if (this.sourceInputElementId) {
                        this.uploadType = f.form;
                        var r = o.getElementById(this.sourceInputElementId);
                        if (r instanceof HTMLInputElement) {
                            "file" !== r.type && a.throwError("input elemenet must be of type 'file'");
                            if (!r.value) {
                                this.error({
                                    errorCode: 0,
                                    message: "user has not supplied a file to upload"
                                });
                                this.invalidFile = !0;
                                return
                            }
                            var i = r.files;
                            i && window.FileReader || a.throwError("browser does not support Files API");
                            1 !== i.length && a.throwError("can not upload more than one file at a time");
                            var s = i[0];
                            s || a.throwError("missing file input");
                            if (s.size > g) {
                                this.error({
                                    errorCode: 1,
                                    message: "the user has selected a file larger than " + _
                                });
                                this.invalidFile = !0;
                                return
                            }
                            this.fileName = t || s.name;
                            this.fileInput = s
                        } else a.throwError("element was not an instance of an HTMLInputElement")
                    } else a.throwError("please specified one type of resource to save")
                };
                return t
            } (s);
            t.exports = v
        },
        {
            "../Constants": 1,
            "../utilities/CallbackInvoker": 14,
            "../utilities/DomUtilities": 15,
            "../utilities/ErrorHandler": 16,
            "../utilities/Logging": 18,
            "../utilities/StringUtilities": 25,
            "../utilities/TypeValidators": 26,
            "../utilities/UrlUtilities": 27,
            "./InvokerOptions": 7,
            "./PickerMode": 9,
            "./SaverActionType": 11,
            "./UploadType": 13
        }],
        13 : [function(e, t, r) {
            "use strict";
            var i; !
            function(e) {
                e[e.dataUrl = 0] = "dataUrl";
                e[e.form = 1] = "form";
                e[e.url = 2] = "url"
            } (i || (i = {}));
            t.exports = i
        },
        {}],
        14 : [function(e, t, r) {
            var i = e("../Constants"),
            n = e("../OneDriveState"),
            o = function() {
                function e() {}
                e.invokeAppCallback = function(e, t) {
                    for (var r = [], o = 2; o < arguments.length; o++) r[o - 2] = arguments[o];
                    t && n.clearState();
                    typeof e === i.TYPE_FUNCTION && e.apply(null, r)
                };
                e.invokeCallbackAsynchronous = function(e) {
                    for (var t = [], r = 1; r < arguments.length; r++) t[r - 1] = arguments[r];
                    window.setTimeout(function() {
                        e.apply(null, t)
                    },
                    0)
                };
                return e
            } ();
            t.exports = o
        },
        {
            "../Constants": 1,
            "../OneDriveState": 4
        }],
        15 : [function(e, t, r) {
            var i = e("./Logging"),
            n = e("../OneDriveState"),
            o = e("../../onedrive-sdk/utilities/ErrorHandler"),
            a = "debug",
            s = "enable-logging",
            l = "onedrive-js",
            c = function() {
                function e() {}
                e.getScriptInput = function() {
                    var t = e.getElementById(l);
                    if (t) {
                        var r = t.getAttribute(s);
                        "true" === r && (i.loggingEnabled = !0);
                        var c = t.getAttribute(a);
                        "true" === c && (n.debug = !0)
                    }
                    window.localStorage || o.throwError("current browser is not compatible with OneDriveSDK.")
                };
                e.getElementById = function(e) {
                    return document.getElementById(e)
                };
                e.onDocumentReady = function(e) {
                    "interactive" === document.readyState || "complete" === document.readyState ? e() : document.addEventListener("DOMContentLoaded", e, !1)
                };
                return e
            } ();
            t.exports = c
        },
        {
            "../../onedrive-sdk/utilities/ErrorHandler": 16,
            "../OneDriveState": 4,
            "./Logging": 18
        }],
        16 : [function(e, t, r) {
            var i = e("./Logging"),
            n = "[OneDriveSDK Error] ",
            o = function() {
                function e() {}
                e.registerErrorObserver = function(t) {
                    e._errorObservers.push(t)
                };
                e.throwError = function(t) {
                    var r = e._errorObservers;
                    for (var o in r) try {
                        r[o]()
                    } catch(a) {
                        i.logError("exception thrown invoking error observer", a)
                    }
                    throw new Error(n + t)
                };
                e._errorObservers = [];
                return e
            } ();
            t.exports = o
        },
        {
            "./Logging": 18
        }],
        17 : [function(e, t, r) {
            var i = e("../models/ApiEndpoint"),
            n = e("../Constants"),
            o = e("./Logging"),
            a = e("./ObjectUtilities"),
            s = e("../OneDriveState"),
            l = e("./UrlUtilities"),
            c = e("./XHR"),
            d = e("./StringUtilities"),
            u = 10,
            p = function() {
                function e() {}
                e.callGraphShareBatch = function(t, r, i, n) {
                    var a, s = [],
                    l = 0,
                    c = [],
                    p = function(n, o) {
                        for (var l = n; o > l; l++) {
                            var d = function(e, t) {
                                e.permissions = [t];
                                c.push(e);
                                a()
                            },
                            u = function(e) {
                                s.push(e.id);
                                a()
                            };
                            e._callGraphShare(t, r[l], i, d, u)
                        }
                    };
                    a = function() {
                        if (++l === r.length) {
                            s.length && o.logMessage(d.format("Create sharing link failed for {0} items", s.length));
                            n(c, s)
                        } else l % u === 0 && p(l, Math.min(r.length, l + u))
                    };
                    p(0, Math.min(r.length, u))
                };
                e.callGraphGetODC = function(t, r, u, p, f) {
                    var h = t.apiEndpoint,
                    g = t.accessToken,
                    _ = t.itemId,
                    v = {
                        Authorization: "bearer " + g
                    };
                    switch (h) {
                    case i.graph_odc:
                        break;
                    case i.graph_odb:
                    }
                    var m = l.appendToPath(t.apiEndpointUrl, "drive/items/" + _),
                    E = new c({
                        url: l.appendQueryStrings(m, r),
                        clientId: s.clientId,
                        method: n.HTTP_GET,
                        apiEndpoint: h,
                        headers: v
                    });
                    o.logMessage("performing GET on item with id: " + _);
                    E.start(function(r, i) {
                        var n = a.deserializeJSON(r.responseText);
                        "share" === t.action ? e._callGraphShare(t, n, f,
                        function(e, t) {
                            e.webUrl = t.link.webUrl;
                            u(e)
                        },
                        function() {
                            o.logError(d.format("Create link failed for bundle with id {0}:", n.id));
                            p()
                        }) : u(n)
                    },
                    function(e, t, r) {
                        p()
                    })
                };
                e.callGraphGetODB = function(e, t, r, i) {
                    var p = e.apiEndpointUrl;
                    p += e.sharePointDriveId ? "drives/" + e.sharePointDriveId + "/items": "drive/items/";
                    var f, h = e.apiEndpoint,
                    g = e.accessToken,
                    _ = e.itemIds,
                    v = {
                        Authorization: "bearer " + g,
                        "Cache-Control": "no-cache, no-store, must-revalidate"
                    },
                    m = [],
                    E = 0,
                    T = 0,
                    w = _.length,
                    S = function(e, r) {
                        o.logMessage(d.format("running batch for items {0} - {1}", e + 1, r + 1));
                        for (var i = e; r > i; i++) {
                            var u = _[i],
                            g = l.appendToPath(p, u + "/?" + t),
                            T = new c({
                                url: g,
                                clientId: s.clientId,
                                method: n.HTTP_GET,
                                apiEndpoint: h,
                                headers: v
                            });
                            o.logMessage("performing GET on item with id: " + u);
                            T.start(function(e, t, r) {
                                var i = a.deserializeJSON(e.responseText);
                                m.push(i);
                                f()
                            },
                            function(e, t, r) {
                                E++;
                                f()
                            })
                        }
                    };
                    f = function() {
                        if (++T === w) {
                            if (m.length) {
                                o.logMessage(d.format("GET metadata succeeded for '{0}' items", m.length));
                                r(m)
                            }
                            if (E) {
                                o.logMessage(d.format("GET metadata failed for '{0}' items", E));
                                i(E)
                            }
                        } else T % u === 0 && S(T, Math.min(w, T + u))
                    };
                    S(0, Math.min(w, u))
                };
                e._callGraphShare = function(e, t, r, i, l) {
                    var u = e.apiEndpointUrl + "drive/items/" + t.id + "/" + e.apiActionNamingSpace + ".createLink",
                    p = new c({
                        url: u,
                        clientId: s.clientId,
                        method: n.HTTP_POST,
                        apiEndpoint: e.apiEndpoint,
                        headers: {
                            Authorization: "bearer " + e.accessToken
                        },
                        json: JSON.stringify(r)
                    });
                    p.start(function(e, r, n) {
                        o.logMessage(d.format("POST createLink succeeded via path {0}", u));
                        i(t, a.deserializeJSON(e.responseText))
                    },
                    function(e, r, i) {
                        o.logMessage(d.format("POST createLink failed via path {0}", u));
                        l(t)
                    })
                };
                return e
            } ();
            t.exports = p
        },
        {
            "../Constants": 1,
            "../OneDriveState": 4,
            "../models/ApiEndpoint": 5,
            "./Logging": 18,
            "./ObjectUtilities": 19,
            "./StringUtilities": 25,
            "./UrlUtilities": 27,
            "./XHR": 29
        }],
        18 : [function(e, t, r) {
            "use strict";
            var i = "onedrive_enable_logging",
            n = "[OneDriveSDK] ",
            o = function() {
                function e() {}
                e.logError = function(t) {
                    for (var r = [], i = 1; i < arguments.length; i++) r[i - 1] = arguments[i];
                    e._log(t, !0, r)
                };
                e.logMessage = function(t) {
                    e._log(t, !1)
                };
                e._log = function(t, r) {
                    for (var o = [], a = 2; a < arguments.length; a++) o[a - 2] = arguments[a]; (r || e.loggingEnabled || window[i]) && console.log(n + t, o)
                };
                e.loggingEnabled = !1;
                return e
            } ();
            t.exports = o
        },
        {}],
        19 : [function(e, t, r) {
            var i = e("../Constants"),
            n = e("./Logging"),
            o = function() {
                function e() {}
                e.shallowClone = function(e) {
                    if (typeof e !== i.TYPE_OBJECT || !e) return null;
                    var t = {};
                    for (var r in e) e.hasOwnProperty(r) && (t[r] = e[r]);
                    return t
                };
                e.deserializeJSON = function(e) {
                    var t = null;
                    try {
                        t = JSON.parse(e)
                    } catch(r) {
                        n.logError("deserialization error" + r)
                    }
                    typeof t === i.TYPE_OBJECT && null !== t || (t = {});
                    return t
                };
                e.serializeJSON = function(e) {
                    return JSON.stringify(e)
                };
                return e
            } ();
            t.exports = o
        },
        {
            "../Constants": 1,
            "./Logging": 18
        }],
        20 : [function(e, t, r) {
            var i = e("../Constants"),
            n = e("./ErrorHandler"),
            o = e("../models/ErrorType"),
            a = e("./GraphWrapper"),
            s = e("./Logging"),
            l = e("./ObjectUtilities"),
            c = e("./Popup"),
            d = e("../models/PickerActionType"),
            u = e("../models/PickerOptions"),
            p = e("./RedirectUtilities"),
            f = e("./StringUtilities"),
            h = e("./UrlUtilities"),
            g = function() {
                function e(e) {
                    var t = l.shallowClone(e);
                    this._pickerOptions = new u(t)
                }
                e.prototype.launchPicker = function() {
                    var e = this,
                    t = this._pickerOptions,
                    r = t.serializeState();
                    if (t.openInNewWindow) {
                        var n = h.appendQueryStrings(t.advanced.redirectUri, {
                            sdk_state: JSON.stringify(r),
                            state: i.STATE_OPEN_POPUP
                        }),
                        o = new c(n,
                        function(t) {
                            e.handlePickerSuccess(t)
                        },
                        function(t) {
                            e.handlePickerError(t)
                        });
                        o.openPopup() || t.error(i.ERROR_POPUP_OPEN)
                    } else p.redirectToAADLogin(t, r)
                };
                e.prototype.handlePickerSuccess = function(e) {
                    var t = e.pickerType,
                    r = this._pickerOptions,
                    o = i.DEFAULT_QUERY_ITEM_PARAMETER;
                    r.action === d.query && r.advanced.queryParameters ? o = r.advanced.queryParameters: r.action === d.download && (o += ",@content.downloadUrl");
                    switch (t) {
                    case i.STATE_MSA_PICKER:
                        this._handleMSAOpenResponse(e, o);
                        break;
                    case i.STATE_AAD_PICKER:
                        this._handleAADOpenResponse(e, o);
                        break;
                    default:
                        n.throwError("invalid value for picker type: " + t)
                    }
                };
                e.prototype.handlePickerError = function(e) {
                    e.error === i.ERROR_ACCESS_DENIED ? this._pickerOptions.cancel() : this._pickerOptions.error({
                        errorCode: o.unknown,
                        message: "something went wrong: " + e.error
                    })
                };
                e.prototype._handleMSAOpenResponse = function(e, t) {
                    var r, n, o = this,
                    s = this._pickerOptions;
                    if (s.action === d.share) {
                        r = s.advanced.createLinkParameters || {
                            type: "view"
                        };
                        n = function(t) {
                            a.callGraphShareBatch(e, t.children, r,
                            function(e, r) {
                                var i = null;
                                s.action === d.share && (s.multiSelect ? i = t.webUrl: e && e[0] && e[0].permissions && e[0].permissions[0] && (i = e[0].permissions[0].link.webUrl));
                                o._handleSuccessResponse({
                                    webUrl: i,
                                    files: e
                                })
                            })
                        }
                    } else n = function(e) {
                        o._handleSuccessResponse({
                            webUrl: e.webUrl,
                            files: e.children
                        },
                        !0)
                    };
                    var l = {
                        select: "id,webUrl",
                        expand: "children(" + t.replace("&", ";") + ")"
                    };
                    a.callGraphGetODC(e, l, n,
                    function() {
                        s.error(i.ERROR_WEB_REQUEST)
                    },
                    r)
                };
                e.prototype._handleAADOpenResponse = function(e, t) {
                    var r, n, o = this,
                    s = this._pickerOptions;
                    if (s.action === d.share) {
                        r = s.advanced.createLinkParameters || {
                            type: "view",
                            scope: "organization"
                        };
                        n = function(t) {
                            a.callGraphShareBatch(e, t, r,
                            function(e, t) {
                                o._handleSuccessResponse({
                                    webUrl: null,
                                    files: e
                                })
                            })
                        }
                    } else n = function(e) {
                        o._handleSuccessResponse({
                            webUrl: null,
                            files: e
                        },
                        !0)
                    };
                    a.callGraphGetODB(e, t, n,
                    function() {
                        s.error(i.ERROR_WEB_REQUEST)
                    })
                };
                e.prototype._handleSuccessResponse = function(e, t) {
                    var r = this._pickerOptions,
                    i = {
                        webUrl: r.action === d.share ? e.webUrl: null,
                        value: []
                    },
                    n = e.files;
                    n && n.length || r.error({
                        errorCode: o.badResponse,
                        message: "no files returned"
                    });
                    s.logMessage(f.format("returning '{0}' files picked", n.length));
                    for (var a = 0; a < n.length; a++) {
                        var l = n[a];
                        i.value.push(l)
                    }
                    r.success(i)
                };
                return e
            } ();
            t.exports = g
        },
        {
            "../Constants": 1,
            "../models/ErrorType": 6,
            "../models/PickerActionType": 8,
            "../models/PickerOptions": 10,
            "./ErrorHandler": 16,
            "./GraphWrapper": 17,
            "./Logging": 18,
            "./ObjectUtilities": 19,
            "./Popup": 21,
            "./RedirectUtilities": 22,
            "./StringUtilities": 25,
            "./UrlUtilities": 27
        }],
        21 : [function(e, t, r) {
            var i = e("./CallbackInvoker"),
            n = e("../Constants"),
            o = e("./Logging"),
            a = e("./ResponseParser"),
            s = 800,
            l = 650,
            c = 500,
            d = function() {
                function e(e, t, r) {
                    this._messageCallbackInvoked = !1;
                    this._url = e;
                    this._successCallback = t;
                    this._failureCallback = r
                }
                e.canReceiveMessage = function(e) {
                    return e.origin === window.location.origin
                };
                e._createPopupFeatures = function() {
                    var e = window.screenX + Math.max(window.outerWidth - s, 0) / 2,
                    t = window.screenY + Math.max(window.outerHeight - l, 0) / 2,
                    r = ["width=" + s, "height=" + l, "top=" + t, "left=" + e, "status=no", "resizable=yes", "toolbar=no", "menubar=no", "scrollbars=yes"];
                    return r.join(",")
                };
                e.prototype.openPopup = function() {
                    if (e._currentPopup && e._currentPopup._isPopupOpen()) return ! 1;
                    window.onedriveReceiveMessage || (window.onedriveReceiveMessage = function(t) {
                        var r = e._currentPopup;
                        if (r && r._isPopupOpen()) {
                            e._currentPopup = null;
                            var n = a.parsePickerResponse(t);
                            r._messageCallbackInvoked = !0;
                            void 0 === n.error ? i.invokeCallbackAsynchronous(r._successCallback, n) : i.invokeCallbackAsynchronous(r._failureCallback, n)
                        }
                    });
                    this._popup = window.open(this._url, "_blank", e._createPopupFeatures());
                    this._popup.focus();
                    this._createPopupPinger();
                    e._currentPopup = this;
                    return ! 0
                };
                e.prototype._createPopupPinger = function() {
                    var t = this,
                    r = window.setInterval(function() {
                        if (t._isPopupOpen()) t._popup.postMessage("ping", "*");
                        else {
                            window.clearInterval(r);
                            e._currentPopup = null;
                            if (!t._messageCallbackInvoked) {
                                o.logMessage("closed callback");
                                t._failureCallback({
                                    error: n.ERROR_ACCESS_DENIED
                                })
                            }
                        }
                    },
                    c)
                };
                e.prototype._isPopupOpen = function() {
                    return null !== this._popup && !this._popup.closed
                };
                return e
            } ();
            t.exports = d
        },
        {
            "../Constants": 1,
            "./CallbackInvoker": 14,
            "./Logging": 18,
            "./ResponseParser": 23
        }],
        22 : [function(e, t, r) {
            var i = e("../Constants"),
            n = e("./DomUtilities"),
            o = e("./ErrorHandler"),
            a = e("./Logging"),
            s = e("./ObjectUtilities"),
            l = e("../OneDriveState"),
            c = e("../models/PickerMode"),
            d = e("./StringUtilities"),
            u = e("./TypeValidators"),
            p = e("./UrlUtilities"),
            f = e("./WindowState"),
            h = e("./XHR"),
            g = e("../models/InvokerOptions"),
            _ = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
            v = function() {
                function e() {}
                e.redirect = function(e, t, r) {
                    void 0 === t && (t = null);
                    void 0 === r && (r = null);
                    t && f.setWindowState(t, r);
                    p.validateUrlProtocol(e);
                    window.location.replace(e)
                };
                e.handleRedirect = function() {
                    var t = p.readCurrentUrlParameters(),
                    r = f.getWindowState(),
                    n = t[i.PARAM_STATE];
                    if (!n) return null;
                    a.logMessage("current state: " + n);
                    n !== i.STATE_OPEN_POPUP || r.options || (r = JSON.parse(t[i.PARAM_SDK_STATE]));
                    var s = r.options,
                    d = r.optionsMode,
                    h = c[d];
                    s || o.throwError("missing options from serialized state");
                    if (t[i.PARAM_ERROR] === i.ERROR_LOGIN_REQUIRED || t[i.PARAM_ERROR] === i.ERROR_AUTH_REQUIRED) {
                        l.deleteLoghinHint(s.clientId);
                        l.loginHint = null;
                        e.redirectToAADLogin(s, r, !0);
                        return null
                    }
                    var g = u.validateType(s.openInNewWindow, i.TYPE_BOOLEAN);
                    g && e._displayOverlay();
                    s.advanced && s.advanced.redirectUri || o.throwError("advanced options is missing");
                    p.validateRedirectUrlHost(s.advanced.redirectUri);
                    switch (n) {
                    case i.STATE_OPEN_POPUP:
                        e.redirectToAADLogin(s, r);
                        break;
                    case i.STATE_AAD_LOGIN:
                        e._handleAADLogin(t, s, h);
                        break;
                    case i.STATE_MSA_PICKER:
                    case i.STATE_AAD_PICKER:
                        var _ = {
                            windowState: r,
                            queryParameters: t
                        };
                        a.logMessage("sending invoker response");
                        if (!g) return _;
                        e._sendResponse(_);
                        break;
                    default:
                        o.throwError("invalid value for redirect state: " + n)
                    }
                    return null
                };
                e.redirectToAADLogin = function(t, r, n) {
                    t.openInNewWindow || e._displayOverlay();
                    var o;
                    if (t.isSharePointRedirect) {
                        o = p.appendQueryStrings("https://login.microsoftonline.com/common/oauth2/authorize", {
                            redirect_uri: t.advanced.redirectUri,
                            client_id: t.clientId,
                            response_type: "token",
                            state: i.STATE_AAD_LOGIN,
                            resource: p.getOrigin(g.getProtectedInfo(t.advanced.sharePointDoclibUrl))
                        });
                        t.advanced.loginHint && (o = p.appendQueryString(o, "login_hint", t.advanced.loginHint));
                        e.redirect(o, r)
                    } else {
                        t.advanced && t.advanced.loginHint ? l.loginHint = {
                            loginHint: t.advanced.loginHint,
                            timeStamp: (new Date).getTime()
                        }: l.getLoginHint(t.clientId);
                        o = _; ! n && l.loginHint && (new Date).getTime() - i.LOGIN_HINT_LIFE_SPAN < l.loginHint.timeStamp && (o = p.appendQueryStrings(o, {
                            login_hint: l.loginHint.loginHint
                        }));
                        o = p.appendQueryStrings(o, {
                            redirect_uri: t.advanced.redirectUri,
                            client_id: t.clientId,
                            scope: "profile openid https://graph.microsoft.com/Files.ReadWrite https://graph.microsoft.com/User.Read",
                            response_mode: "fragment",
                            state: i.STATE_AAD_LOGIN,
                            nonce: p.generateNonce()
                        });
                        o += "&response_type=id_token+token";
                        e.redirect(o, r)
                    }
                };
                e._handleAADLogin = function(t, r, n) {
                    r.openInNewWindow || e._displayOverlay();
                    r.advanced.accessToken && window.localStorage.removeItem(r.advanced.accessToken);
                    r.advanced.accessToken = g.setProtectedInfo(t[i.PARAM_ACCESS_TOKEN]);
                    if (r.isSharePointRedirect) e._redirectToTenant(r, n, r.openInNewWindow);
                    else {
                        var a = t[i.PARAM_ID_TOKEN];
                        a || o.throwError("id_toekn is missing in returned parameters");
                        var s = e._parseTenantId(a);
                        if (s.tid === i.CUSTOMER_TID) {
                            if (s.preferred_username) {
                                l.loginHint = {
                                    loginHint: s.preferred_username,
                                    timeStamp: (new Date).getTime()
                                };
                                l.updateLoginHint(r.clientId)
                            }
                            e._redirectToODCPicker(r, n)
                        } else {
                            if (s.preferred_username) {
                                l.loginHint = {
                                    loginHint: s.preferred_username,
                                    timeStamp: (new Date).getTime()
                                };
                                l.updateLoginHint(r.clientId)
                            }
                            e._redirectToTenant(r, n, r.openInNewWindow)
                        }
                    }
                };
                e._redirectToODCPicker = function(t, r) {
                    var n, o, a;
                    switch (r) {
                    case c.open:
                        n = "read";
                        o = "files";
                        a = t.multiSelect ? "multiple": "single";
                        break;
                    case c.save:
                        n = "readwrite";
                        o = "folders";
                        a = "single"
                    }
                    var s = "https://login." + l.getODCHost() + "/oauth20_authorize.srf",
                    d = p.appendQueryStrings(s, {
                        client_id: t.clientId,
                        redirect_uri: t.advanced.redirectUri,
                        response_type: "token",
                        scope: "onedrive_onetime.access:" + n + o + "|" + a + "|downloadLink",
                        state: i.STATE_MSA_PICKER
                    });
                    e.redirect(d, {
                        options: t,
                        picker: {
                            accessLevel: n,
                            selectionMode: a,
                            viewType: o,
                            filter: t.advanced.filter,
                            linkType: "download"
                        }
                    })
                };
                e._redirectToTenant = function(t, r, n) {
                    var a = function(i) {
                        var n, o, a;
                        switch (r) {
                        case c.open:
                            n = "read";
                            o = "files";
                            a = t.multiSelect ? "multiple": "single";
                            break;
                        case c.save:
                            n = "readwrite";
                            o = "folders";
                            a = "single"
                        }
                        e.redirect(i, {
                            picker: {
                                accessLevel: n,
                                selectionMode: a,
                                viewType: o,
                                filter: t.advanced.filter
                            },
                            options: t,
                            ODBParams: {
                                p: "2"
                            }
                        })
                    };
                    if (t.isSharePointRedirect) {
                        var l = g.getProtectedInfo(t.advanced.sharePointDoclibUrl);
                        l.indexOf("-my") > -1 && (l += "_layouts/onedrive.aspx");
                        p.validateUrlProtocol(l, [i.PROTOCOL_HTTPS]);
                        a(l)
                    } else {
                        var u = new h({
                            url: p.appendQueryString(i.GRAPH_URL + "me", "$select", "mySite"),
                            method: i.HTTP_GET,
                            headers: {
                                Authorization: "bearer " + g.getProtectedInfo(t.advanced.accessToken),
                                Accept: "application/json"
                            }
                        });
                        u.start(function(e, t) {
                            var r = s.deserializeJSON(e.responseText);
                            r.mySite ? a(r.mySite + "_layouts/onedrive.aspx") : o.throwError(d.format("Cannot find the personal tenant url, response text: {0}", e.responseText))
                        },
                        function(t, r, i) {
                            e._handleError(d.format("graph/me request failed, status code: '{0}', response text: '{1}'", h.statusCodeToString(r), t.responseText), n)
                        })
                    }
                };
                e._sendResponse = function(t) {
                    var r = window.opener;
                    if (r.onedriveReceiveMessage) r.onedriveReceiveMessage(t);
                    else {
                        a.logError("error in window's opener, pop up will close.");
                        e._handleError("SDK message receiver is undefined.", !0)
                    }
                    window.close()
                };
                e._handleError = function(t, r) {
                    var n = {};
                    n[i.PARAM_ERROR] = t;
                    if (r) e._sendResponse({
                        queryParameters: n
                    });
                    else {
                        a.logMessage("error in picker flow, redirecting back to app");
                        n[i.PARAM_STATE] = i.STATE_AAD_PICKER;
                        var o = p.trimUrlQuery(window.location.href);
                        e.redirect(p.appendQueryStrings(o, n))
                    }
                };
                e._parseTenantId = function(e) {
                    var t = e.split(".")[1];
                    if (t) {
                        var r = t.replace("-", "+").replace("_", "/");
                        try {
                            var i = JSON.parse(atob(r));
                            i.tid || o.throwError("tid is missing in parsed open id");
                            i.preferred_username || o.throwError("preferred_username is missing in parsed open id");
                            return i
                        } catch(n) {
                            o.throwError("Base64URL decode and JSON parse of JWT segment failed")
                        }
                    }
                    o.throwError("Received invalid JWT format token")
                };
                e._displayOverlay = function() {
                    var e = document.createElement("div"),
                    t = ["position: fixed", "width: 100%", "height: 100%", "top: 0px", "left: 0px", "background-color: white", "opacity: 1", "z-index: 10000"];
                    e.id = "od-overlay";
                    e.style.cssText = t.join(";");
                    var r = document.createElement("img"),
                    i = ["position: absolute", "top: calc(50% - 40px)", "left: calc(50% - 40px)"];
                    r.id = "od-spinner";
                    r.src = "https://p.sfx.ms/common/spinner_grey_40_transparent.gif";
                    r.style.cssText = i.join(";");
                    e.appendChild(r);
                    var o = document.createElement("style");
                    o.type = "text/css";
                    o.innerHTML = "body { visibility: hidden !important; }";
                    document.head.appendChild(o);
                    n.onDocumentReady(function() {
                        var t = document.body;
                        null !== t ? t.insertBefore(e, t.firstChild) : document.createElement("body").appendChild(e);
                        document.head.removeChild(o)
                    })
                };
                return e
            } ();
            t.exports = v
        },
        {
            "../Constants": 1,
            "../OneDriveState": 4,
            "../models/InvokerOptions": 7,
            "../models/PickerMode": 9,
            "./DomUtilities": 15,
            "./ErrorHandler": 16,
            "./Logging": 18,
            "./ObjectUtilities": 19,
            "./StringUtilities": 25,
            "./TypeValidators": 26,
            "./UrlUtilities": 27,
            "./WindowState": 28,
            "./XHR": 29
        }],
        23 : [function(e, t, r) {
            var i = e("../models/ApiEndpoint"),
            n = e("../Constants"),
            o = e("./ErrorHandler"),
            a = e("./Logging"),
            s = e("./UrlUtilities"),
            l = e("../models/InvokerOptions"),
            c = "0000000000000000",
            d = c.length,
            u = new RegExp("^\\w+\\.\\w+:\\w+[\\|\\w+]+:([\\w]+\\!\\d+)(?:\\!(.+))*$"),
            p = function() {
                function e() {}
                e.parsePickerResponse = function(t) {
                    a.logMessage("parsing picker response");
                    var r = t.windowState;
                    r || o.throwError("missing windowState from picker response");
                    var i = t.queryParameters;
                    i || o.throwError("missing queryParameters from picker response");
                    var s = i[n.PARAM_ERROR];
                    if (s) return {
                        error: s
                    };
                    var c, d = i[n.PARAM_STATE];
                    if (r.options.advanced.accessToken) {
                        c = l.getProtectedInfo(r.options.advanced.accessToken);
                        window.localStorage.removeItem(r.options.advanced.accessToken)
                    }
                    var u = {
                        pickerType: d,
                        accessToken: c
                    };
                    u.action = r.options.action;
                    switch (d) {
                    case n.STATE_MSA_PICKER:
                        e._parseMSAResponse(u, i);
                        break;
                    case n.STATE_AAD_PICKER:
                        e._parseAADResponse(u, i, r);
                        break;
                    default:
                        o.throwError("invalid value for picker type: " + d)
                    }
                    u.accessToken || o.throwError("missing access token");
                    u.apiEndpointUrl || o.throwError("missing API endpoint URL");
                    return u
                };
                e._parseMSAResponse = function(t, r) {
                    t.apiEndpoint = i.graph_odc;
                    t.apiEndpointUrl = n.GRAPH_URL;
                    t.apiActionNamingSpace = "microsoft.graph";
                    var a = r.scope;
                    a || o.throwError("missing 'scope' paramter from MSA picker response");
                    for (var s, l = a.split(" "), c = 0; c < l.length && !s; c++) s = u.exec(l[c]);
                    s || o.throwError("scope was not formatted correctly");
                    var d = s[1].split("_"),
                    p = d[1],
                    f = p.indexOf("!"),
                    h = p.substring(0, f),
                    g = p.substring(f),
                    _ = e._leftPadCid(h),
                    v = _ + g;
                    t.ownerCid = _;
                    t.itemId = v;
                    t.authKey = s[2]
                };
                e._parseAADResponse = function(e, t, r) {
                    if (r.options.isSharePointRedirect) {
                        e.apiEndpointUrl = s.getOrigin(l.getProtectedInfo(r.options.advanced.sharePointDoclibUrl)) + "_api/v2.0/";
                        e.sharePointDriveId = r.options.advanced.sharePointDriveId;
                        e.apiEndpoint = i.filesV2;
                        e.apiActionNamingSpace = "action";
                        window.localStorage.removeItem(r.options.advanced.sharePointDoclibUrl)
                    } else {
                        e.apiEndpoint = i.graph_odb;
                        e.apiEndpointUrl = n.GRAPH_URL + "me/";
                        e.apiActionNamingSpace = "microsoft.graph"
                    }
                    var a = t["item-id"].split(",");
                    a.length || o.throwError("missing item id(s)");
                    e.itemIds = a
                };
                e._leftPadCid = function(e) {
                    return e.length === d ? e: c.substring(0, d - e.length) + e
                };
                return e
            } ();
            t.exports = p
        },
        {
            "../Constants": 1,
            "../models/ApiEndpoint": 5,
            "../models/InvokerOptions": 7,
            "./ErrorHandler": 16,
            "./Logging": 18,
            "./UrlUtilities": 27
        }],
        24 : [function(e, t, r) {
            var i = e("../models/ApiEndpoint"),
            n = e("./CallbackInvoker"),
            o = e("../Constants"),
            a = e("./ErrorHandler"),
            s = e("../models/ErrorType"),
            l = e("./GraphWrapper"),
            c = e("./Logging"),
            d = e("./ObjectUtilities"),
            u = e("../OneDriveState"),
            p = e("./Popup"),
            f = e("./RedirectUtilities"),
            h = e("../models/SaverActionType"),
            g = e("../models/SaverOptions"),
            _ = e("./StringUtilities"),
            v = e("../models/UploadType"),
            m = e("./UrlUtilities"),
            E = e("./XHR"),
            T = 1e3,
            w = 5,
            S = function() {
                function e(e) {
                    var t = d.shallowClone(e);
                    this._saverOptions = new g(t)
                }
                e.prototype.launchSaver = function() {
                    var e = this,
                    t = this._saverOptions;
                    if (!t.invalidFile) {
                        var r = t.serializeState();
                        if (t.openInNewWindow) {
                            var i = m.appendQueryStrings(t.advanced.redirectUri, {
                                sdk_state: JSON.stringify(r),
                                state: o.STATE_OPEN_POPUP
                            }),
                            n = new p(i,
                            function(t) {
                                e.handleSaverSuccess(t)
                            },
                            function(t) {
                                e.handleSaverError(t)
                            });
                            n.openPopup() || t.error(o.ERROR_POPUP_OPEN)
                        } else f.redirectToAADLogin(t, r)
                    }
                };
                e.prototype.handleSaverSuccess = function(e) {
                    var t = this,
                    r = e.pickerType,
                    i = o.DEFAULT_QUERY_ITEM_PARAMETER,
                    n = this._saverOptions;
                    n.action === h.query && n.advanced.queryParameters && (i = n.advanced.queryParameters);
                    switch (r) {
                    case o.STATE_MSA_PICKER:
                        l.callGraphGetODC(e, {
                            select: "id, webUrl",
                            expand: "children(" + i.replace("&", ";") + ")"
                        },
                        function(r) {
                            var i = r.children;
                            i || a.throwError("empty API response");
                            var o = i[0];
                            o && 1 === i.length || a.throwError("incorrect number of folders returned");
                            n.action === h.query ? n.success({
                                webUrl: null,
                                value: [o]
                            }) : n.action === h.save && t._executeUpload(e, o)
                        },
                        function() {
                            n.error(o.ERROR_WEB_REQUEST)
                        });
                        break;
                    case o.STATE_AAD_PICKER:
                        var s = e.itemIds;
                        1 !== s.length && a.throwError("incorrect number of folders returned");
                        var c = s[0];
                        c || (c = "root");
                        u.debug && OneDriveDebug && (OneDriveDebug.accessToken = e.accessToken);
                        n.action === h.query ? l.callGraphGetODB(e, i,
                        function(e) {
                            n.success({
                                webUrl: null,
                                value: [e]
                            })
                        },
                        function() {
                            n.error(o.ERROR_WEB_REQUEST)
                        }) : this._executeUpload(e, {
                            id: c
                        });
                        break;
                    default:
                        a.throwError("invalid value for picker type: " + r)
                    }
                };
                e.prototype.handleSaverError = function(e) {
                    e.error === o.ERROR_ACCESS_DENIED ? this._saverOptions.cancel() : this._saverOptions.error({
                        errorCode: s.unknown,
                        message: "something went wrong: " + e.error
                    })
                };
                e.prototype._executeUpload = function(e, t) {
                    var r = this._saverOptions.uploadType;
                    c.logMessage(_.format("beginning '{0}' upload", v[r]));
                    var i = e.accessToken;
                    switch (r) {
                    case v.dataUrl:
                    case v.url:
                        this._executeUrlUpload(e, t, i, r);
                        break;
                    case v.form:
                        this._executeFormUpload(e, t, i);
                        break;
                    default:
                        a.throwError("invalid value for upload type: " + r)
                    }
                };
                e.prototype._executeUrlUpload = function(e, t, r, i) {
                    var n = this,
                    a = this._saverOptions;
                    if (i !== v.url || e.pickerType !== o.STATE_AAD_PICKER) {
                        var l = m.appendToPath(e.apiEndpointUrl, "drive/items/" + t.id + "/children"),
                        c = {};
                        c.Prefer = "respond-async";
                        c.Authorization = "bearer " + r;
                        var p = {
                            "@microsoft.graph.sourceUrl": a.sourceUri,
                            name: a.fileName,
                            file: {}
                        };
                        p[this._getContentSourceUrl(e.apiEndpoint)] = a.sourceUri;
                        var f = new E({
                            url: l,
                            clientId: u.clientId,
                            method: o.HTTP_POST,
                            headers: c,
                            json: d.serializeJSON(p),
                            apiEndpoint: e.apiEndpoint
                        });
                        f.start(function(t, r) {
                            if (i !== v.dataUrl || 200 !== r && 201 !== r) if (i === v.url && 202 === r) {
                                var l = t.getResponseHeader("Location");
                                l || a.error({
                                    errorCode: s.badResponse,
                                    message: "missing 'Location' header on response"
                                });
                                n._beginPolling(l,
                                function(t) {
                                    n._getCreatedItem(e, t)
                                })
                            } else a.error(o.ERROR_WEB_REQUEST);
                            else {
                                u.debug && OneDriveDebug && (OneDriveDebug.itemUrl = d.deserializeJSON(t.responseText)["@odata.id"]);
                                a.success(d.deserializeJSON(t.responseText))
                            }
                        },
                        function(e, t, r) {
                            a.error(o.ERROR_WEB_REQUEST)
                        })
                    } else a.error({
                        errorCode: s.unsupportedFeature,
                        message: "URL upload not supported for AAD"
                    })
                };
                e.prototype._executeFormUpload = function(e, t, r) {
                    var i = this._saverOptions,
                    n = i.fileInput,
                    l = null;
                    window.File && n instanceof window.File ? l = new FileReader: a.throwError("file reader not supported");
                    l.onerror = function(e) {
                        c.logError("failed to read or upload the file", e);
                        i.error({
                            errorCode: s.fileReaderFailure,
                            message: "failed to read or upload the file, see console log for details"
                        })
                    };
                    l.onload = function(n) {
                        var a = m.appendToPath(e.apiEndpointUrl, "drive/items/" + t.id + "/children/" + i.fileName + "/content"),
                        s = {};
                        s["@name.conflictBehavior"] = e.pickerType === o.STATE_AAD_PICKER ? "fail": "rename";
                        var l = {};
                        l.Authorization = "bearer " + r;
                        l["Content-Type"] = "multipart/form-data";
                        var c = new E({
                            url: m.appendQueryStrings(a, s),
                            clientId: u.clientId,
                            headers: l,
                            apiEndpoint: e.apiEndpoint
                        }),
                        d = n.target.result;
                        c.upload(d,
                        function(e, t) {
                            i.success({
                                webUrl: null,
                                value: [JSON.parse(e.responseText)]
                            })
                        },
                        function(e, t, r) {
                            i.error(o.ERROR_WEB_REQUEST)
                        },
                        function(e, t) {
                            i.progress(t.progressPercentage)
                        })
                    };
                    l.readAsArrayBuffer(n)
                };
                e.prototype._beginPolling = function(e, t) {
                    c.logMessage("polling for URL upload completion");
                    var r = T,
                    i = w,
                    a = {
                        url: e,
                        method: o.HTTP_GET
                    },
                    l = this._saverOptions,
                    u = function() {
                        var e = new E(a);
                        e.start(function(e, a) {
                            switch (a) {
                            case 202:
                                var c = d.deserializeJSON(e.responseText);
                                l.progress(c.percentageComplete);
                                if (!i--) {
                                    r *= 2;
                                    i = w
                                }
                                n.invokeCallbackAsynchronous(u, r);
                                break;
                            case 200:
                                var p = d.deserializeJSON(e.responseText);
                                l.progress(p.percentageComplete);
                                p.resourceId ? t(p.resourceId) : l.error({
                                    errorCode: s.badResponse,
                                    message: "missing resourceId in return value: " + p
                                });
                                break;
                            default:
                                l.error(o.ERROR_WEB_REQUEST)
                            }
                        },
                        function(e, t, r) {
                            l.error(o.ERROR_WEB_REQUEST)
                        })
                    };
                    n.invokeCallbackAsynchronous(u, r)
                };
                e.prototype._getContentSourceUrl = function(e) {
                    return e === i.graph_odb || i.graph_odc ? "@microsoft.graph.sourceUrl": "@content.sourceUrl"
                };
                e.prototype._getCreatedItem = function(e, t) {
                    var r = this._saverOptions;
                    e.itemId = t;
                    l.callGraphGetODC(e, {},
                    function(e) {
                        r.success({
                            webUrl: null,
                            value: [e]
                        })
                    },
                    function() {
                        r.error(o.ERROR_WEB_REQUEST)
                    })
                };
                return e
            } ();
            t.exports = S
        },
        {
            "../Constants": 1,
            "../OneDriveState": 4,
            "../models/ApiEndpoint": 5,
            "../models/ErrorType": 6,
            "../models/SaverActionType": 11,
            "../models/SaverOptions": 12,
            "../models/UploadType": 13,
            "./CallbackInvoker": 14,
            "./ErrorHandler": 16,
            "./GraphWrapper": 17,
            "./Logging": 18,
            "./ObjectUtilities": 19,
            "./Popup": 21,
            "./RedirectUtilities": 22,
            "./StringUtilities": 25,
            "./UrlUtilities": 27,
            "./XHR": 29
        }],
        25 : [function(e, t, r) {
            "use strict";
            var i = /[\{\}]/g,
            n = /\{\d+\}/g,
            o = function() {
                function e() {}
                e.format = function(e) {
                    for (var t = [], r = 1; r < arguments.length; r++) t[r - 1] = arguments[r];
                    var o = function(e) {
                        var r = t[e.replace(i, "")];
                        null === r && (r = "");
                        return r
                    };
                    return e.replace(n, o)
                };
                return e
            } ();
            t.exports = o
        },
        {}],
        26 : [function(e, t, r) {
            var i = e("../Constants"),
            n = e("./ErrorHandler"),
            o = e("./Logging"),
            a = e("./ObjectUtilities"),
            s = e("./StringUtilities"),
            l = function() {
                function e() {}
                e.validateType = function(t, r, i, l, c) {
                    void 0 === i && (i = !1);
                    if (void 0 === t) {
                        if (i) {
                            void 0 === l && n.throwError("default value missing");
                            o.logMessage("applying default value: " + l);
                            return l
                        }
                        n.throwError("object was missing and not optional")
                    }
                    var d = typeof t;
                    d !== r && n.throwError(s.format("expected object type: '{0}', actual object type: '{1}'", r, d));
                    e._isValidValue(t, c) || n.throwError(s.format("object does not match any valid values: '{0}'", a.serializeJSON(c)));
                    return t
                };
                e.validateCallback = function(e, t, r) {
                    void 0 === t && (t = !1);
                    void 0 === r && (r = !1);
                    if (void 0 === e) {
                        if (t) return null;
                        n.throwError("function was missing and not optional")
                    }
                    var o = typeof e;
                    o !== i.TYPE_STRING && o !== i.TYPE_FUNCTION && n.throwError(s.format("expected function type: 'function | string', actual type: '{0}'", o));
                    var a = null;
                    if (o === i.TYPE_STRING) {
                        var l = window[e];
                        typeof l === i.TYPE_FUNCTION ? a = l: n.throwError(s.format("could not find a function with name '{0}' on the window object", e))
                    } else r ? n.throwError("expected a global function") : a = e;
                    return a
                };
                e._isValidValue = function(e, t) {
                    if (!Array.isArray(t)) return ! 0;
                    for (var r in t) if (e === t[r]) return ! 0;
                    return ! 1
                };
                return e
            } ();
            t.exports = l
        },
        {
            "../Constants": 1,
            "./ErrorHandler": 16,
            "./Logging": 18,
            "./ObjectUtilities": 19,
            "./StringUtilities": 25
        }],
        27 : [function(e, t, r) {
            var i = e("../Constants"),
            n = e("./ErrorHandler"),
            o = e("./StringUtilities"),
            a = function() {
                function e() {}
                e.appendToPath = function(e, t) {
                    return e + ("/" !== e.charAt(e.length - 1) ? "/": "") + t
                };
                e.appendQueryString = function(t, r, i) {
                    return e.appendQueryStrings(t, (n = {},
                    n[r] = i, n));
                    var n
                };
                e.appendQueryStrings = function(e, t, r) {
                    if (!t || 0 === Object.keys(t).length) return e;
                    r ? e += "#": -1 === e.indexOf("?") ? e += "?": "&" !== e.charAt(e.length - 1) && (e += "&");
                    var i = "";
                    for (var n in t) i += (i.length ? "&": "") + o.format("{0}={1}", encodeURIComponent(n), encodeURIComponent(t[n]));
                    return e + i
                };
                e.readCurrentUrlParameters = function() {
                    return e.readUrlParameters(window.location.href)
                };
                e.readUrlParameters = function(t) {
                    var r = {},
                    i = t.indexOf("?") + 1,
                    n = t.indexOf("#") + 1;
                    if (i > 0) {
                        var o = n > i ? n - 1 : t.length;
                        e._deserializeParameters(t.substring(i, o), r)
                    }
                    n > 0 && e._deserializeParameters(t.substring(n), r);
                    return r
                };
                e.trimUrlQuery = function(e) {
                    var t = ["?", "#"];
                    for (var r in t) {
                        var i = e.indexOf(t[r]);
                        i > 0 && (e = e.substring(0, i))
                    }
                    return e
                };
                e.getFileNameFromUrl = function(t) {
                    var r = e.trimUrlQuery(t);
                    return r.substr(r.lastIndexOf("/") + 1)
                };
                e.getOrigin = function(t) {
                    return e.appendToPath(t.replace(/^((\w+:)?\/\/[^\/]+\/?).*$/, "$1"), "")
                };
                e.isPathFullUrl = function(e) {
                    return 0 === e.indexOf("https://") || 0 === e.indexOf("http://")
                };
                e.isPathDataUrl = function(e) {
                    return 0 === e.indexOf("data:")
                };
                e.generateNonce = function() {
                    for (var e = "",
                    t = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",
                    r = 0; r < i.NONCE_LENGTH; r++) e += t.charAt(Math.floor(Math.random() * t.length));
                    return e
                };
                e.validateUrlProtocol = function(e, t) {
                    t = t ? t: [i.PROTOCOL_HTTP, i.PROTOCOL_HTTPS];
                    for (var r = 0,
                    o = t; r < o.length; r++) {
                        var a = o[r];
                        if (0 === e.toUpperCase().indexOf(a)) return
                    }
                    n.throwError("redirect uri does not match any protocol in " + t)
                };
                e.validateRedirectUrlHost = function(t) {
                    e.validateUrlProtocol(t);
                    if (t.indexOf("://") > -1) {
                        var r = t.split("/")[2];
                        r !== window.location.host && n.throwError("redirect uri is not in the same domain as picker sdk")
                    } else n.throwError("redirect uri is not an absolute url")
                };
                e._deserializeParameters = function(e, t) {
                    for (var r = e.split("&"), i = 0; i < r.length; i++) {
                        var n = r[i].split("=");
                        2 === n.length && (t[decodeURIComponent(n[0])] = decodeURIComponent(n[1]))
                    }
                };
                return e
            } ();
            t.exports = a
        },
        {
            "../Constants": 1,
            "./ErrorHandler": 16,
            "./StringUtilities": 25
        }],
        28 : [function(e, t, r) {
            var i = e("./Logging"),
            n = e("./ObjectUtilities"),
            o = function() {
                function e() {}
                e.getWindowState = function() {
                    return n.deserializeJSON(window.name || "{}")
                };
                e.setWindowState = function(t, r) {
                    void 0 === r && (r = null);
                    null === r && (r = e.getWindowState());
                    for (var o in t) r[o] = t[o];
                    var a = n.serializeJSON(r);
                    i.logMessage("window.name = " + a);
                    window.name = a
                };
                return e
            } ();
            t.exports = o
        },
        {
            "./Logging": 18,
            "./ObjectUtilities": 19
        }],
        29 : [function(e, t, r) {
            var i = e("../models/ApiEndpoint"),
            n = e("../Constants"),
            o = e("./ErrorHandler"),
            a = e("./Logging"),
            s = e("./StringUtilities"),
            l = 3e4,
            c = -1,
            d = -2,
            u = -3,
            p = function() {
                function e(e) {
                    this._url = e.url;
                    this._json = e.json;
                    this._headers = e.headers || {};
                    this._method = e.method;
                    this._clientId = e.clientId;
                    this._apiEndpoint = e.apiEndpoint || i.other;
                    o.registerErrorObserver(this._abortRequest)
                }
                e.statusCodeToString = function(e) {
                    switch (e) {
                    case - 1 : return "EXCEPTION";
                    case - 2 : return "TIMEOUT";
                    case - 3 : return "REQUEST ABORTED";
                    default:
                        return e.toString()
                    }
                };
                e.prototype.start = function(e, t) {
                    var r = this;
                    try {
                        this._successCallback = e;
                        this._failureCallback = t;
                        this._request = new XMLHttpRequest;
                        this._request.ontimeout = this._onTimeout;
                        this._request.onreadystatechange = function() {
                            if (!r._completed && 4 === r._request.readyState) {
                                r._completed = !0;
                                var e = r._request.status;
                                400 > e && e > 0 ? r._callSuccessCallback(e) : r._callFailureCallback(e)
                            }
                        };
                        this._method || (this._method = this._json ? n.HTTP_POST: n.HTTP_GET);
                        this._request.open(this._method, this._url, !0);
                        this._request.timeout = l;
                        this._setHeaders();
                        a.logMessage("starting request to: " + this._url);
                        this._request.send(this._json)
                    } catch(i) {
                        this._callFailureCallback(c, i)
                    }
                };
                e.prototype.upload = function(e, t, r, i) {
                    var o = this;
                    try {
                        this._successCallback = t;
                        this._progressCallback = i;
                        this._failureCallback = r;
                        this._request = new XMLHttpRequest;
                        this._request.ontimeout = this._onTimeout;
                        this._method = n.HTTP_PUT;
                        this._request.open(this._method, this._url, !0);
                        this._setHeaders();
                        this._request.onload = function(e) {
                            o._completed = !0;
                            var t = o._request.status;
                            200 === t || 201 === t ? o._callSuccessCallback(t) : o._callFailureCallback(t, e)
                        };
                        this._request.onerror = function(e) {
                            o._completed = !0;
                            o._callFailureCallback(o._request.status, e)
                        };
                        this._request.upload.onprogress = function(e) {
                            if (e.lengthComputable) {
                                var t = {
                                    bytesTransferred: e.loaded,
                                    totalBytes: e.total,
                                    progressPercentage: 0 === e.total ? 0 : e.loaded / e.total * 100
                                };
                                o._callProgressCallback(t)
                            }
                        };
                        a.logMessage("starting upload to: " + this._url);
                        this._request.send(e)
                    } catch(s) {
                        this._callFailureCallback(c, s)
                    }
                };
                e.prototype._callSuccessCallback = function(t) {
                    a.logMessage("calling xhr success callback, status: " + e.statusCodeToString(t));
                    this._successCallback(this._request, t, this._url)
                };
                e.prototype._callFailureCallback = function(t, r) {
                    a.logError("calling xhr failure callback, status: " + e.statusCodeToString(t), this._request, r);
                    this._failureCallback(this._request, t, t === d)
                };
                e.prototype._callProgressCallback = function(e) {
                    a.logMessage("calling xhr upload progress callback");
                    this._progressCallback(this._request, e)
                };
                e.prototype._abortRequest = function() {
                    if (!this._completed) {
                        this._completed = !0;
                        if (this._request) try {
                            this._request.abort()
                        } catch(e) {}
                        this._callFailureCallback(u)
                    }
                };
                e.prototype._onTimeout = function() {
                    if (!this._completed) {
                        this._completed = !0;
                        this._callFailureCallback(d)
                    }
                };
                e.prototype._setHeaders = function() {
                    for (var e in this._headers) this._request.setRequestHeader(e, this._headers[e]);
                    this._clientId && this._apiEndpoint !== i.other && this._request.setRequestHeader("Application", "0x" + this._clientId);
                    var t = s.format("{0}={1}", "SDK-Version", n.SDK_VERSION);
                    switch (this._apiEndpoint) {
                    case i.graph_odb:
                        this._request.setRequestHeader("X-ClientService-ClientTag", t);
                        break;
                    case i.graph_odc:
                        this._request.setRequestHeader("X-RequestStats", t);
                        break;
                    case i.other:
                        break;
                    default:
                        o.throwError("invalid API endpoint: " + this._apiEndpoint)
                    }
                    this._method === n.HTTP_POST && this._request.setRequestHeader("Content-Type", this._json ? "application/json": "text/plain")
                };
                return e
            } ();
            t.exports = p
        },
        {
            "../Constants": 1,
            "../models/ApiEndpoint": 5,
            "./ErrorHandler": 16,
            "./Logging": 18,
            "./StringUtilities": 25
        }]
    },
    {},
    [2])(2)
});