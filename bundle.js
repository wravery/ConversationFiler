/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId])
/******/ 			return installedModules[moduleId].exports;
/******/
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// identity function for calling harmony imports with the correct context
/******/ 	__webpack_require__.i = function(value) { return value; };
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 9);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

module.exports = React;

/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var Data;
(function (Data) {
    var Progress;
    (function (Progress) {
        Progress[Progress["GetCallbackToken"] = 0] = "GetCallbackToken";
        Progress[Progress["GetConversation"] = 1] = "GetConversation";
        Progress[Progress["GetExcludedFolders"] = 2] = "GetExcludedFolders";
        Progress[Progress["GetFolderNames"] = 3] = "GetFolderNames";
        Progress[Progress["Success"] = 4] = "Success";
        Progress[Progress["NotFound"] = 5] = "NotFound";
        Progress[Progress["Error"] = 6] = "Error";
    })(Progress = Data.Progress || (Data.Progress = {}));
})(Data = exports.Data || (exports.Data = {}));


/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/// <reference path="../_references.ts" />

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var React = __webpack_require__(0);
var Model_1 = __webpack_require__(1);
var RESTData_1 = __webpack_require__(5);
var EWSData_1 = __webpack_require__(4);
var StatusMessage_1 = __webpack_require__(8);
var SearchResults_1 = __webpack_require__(7);
var Feedback_1 = __webpack_require__(6);
var ConversationFiler = (function (_super) {
    __extends(ConversationFiler, _super);
    function ConversationFiler(props) {
        var _this = _super.call(this, props) || this;
        _this.state = { progress: Model_1.Data.Progress.GetCallbackToken };
        return _this;
    }
    // Start the chain of requests by getting a callback token.
    ConversationFiler.prototype.componentDidMount = function () {
        var _this = this;
        if (this.props.mockResults) {
            if (this.props.mockResults.length > 0) {
                this.setState({ progress: Model_1.Data.Progress.Success, matches: this.props.mockResults });
            }
            else {
                this.setState({ progress: Model_1.Data.Progress.NotFound });
            }
            return;
        }
        else if (!this.props.mailbox) {
            return;
        }
        var data = this.props.mailbox.restUrl
            ? new RESTData_1.RESTData.Model(this.props.mailbox)
            : new EWSData_1.EWSData.Model(this.props.mailbox);
        this.setState({ data: data });
        data.getItemsAsync(function (results) {
            if (results.length > 0) {
                _this.setState({ progress: Model_1.Data.Progress.Success, matches: results });
            }
            else {
                _this.setState({ progress: Model_1.Data.Progress.NotFound });
            }
        }, function (progress) {
            _this.setState({ progress: progress });
        }, function (message) {
            _this.setState({ progress: Model_1.Data.Progress.Error, error: message });
        });
    };
    ConversationFiler.prototype.onSelection = function (folderId) {
        var _this = this;
        console.log("Selected a folder: " + folderId);
        this.state.data.moveItemsAsync(folderId, function (count) {
            if (_this.props.onComplete) {
                _this.props.onComplete();
            }
        }, function (message) {
            _this.setState({ progress: Model_1.Data.Progress.Error, error: message });
        });
    };
    ConversationFiler.prototype.render = function () {
        return (React.createElement("div", null,
            React.createElement(StatusMessage_1.StatusMessage, { progress: this.state.progress, message: this.state.error }),
            React.createElement(SearchResults_1.SearchResults, { matches: this.state.matches, onSelection: this.onSelection.bind(this) }),
            React.createElement(Feedback_1.Feedback, null)));
    };
    return ConversationFiler;
}(React.Component));
exports.ConversationFiler = ConversationFiler;


/***/ }),
/* 3 */
/***/ (function(module, exports) {

module.exports = ReactDOM;

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var Model_1 = __webpack_require__(1);
var EWSData;
(function (EWSData) {
    var RequestBuilder = (function () {
        function RequestBuilder() {
        }
        RequestBuilder.getItemsRequest = function (messages) {
            var builder = [
                RequestBuilder.beginRequest,
                '  <m:GetItem>',
                '    <m:ItemShape>',
                '      <t:BaseShape>IdOnly</t:BaseShape>',
                '      <t:BodyType>Text</t:BodyType>',
                '      <t:AdditionalProperties>',
                '        <t:FieldURI FieldURI="item:ParentFolderId" />',
                '        <t:FieldURI FieldURI="message:Sender" />',
                '        <t:FieldURI FieldURI="message:ToRecipients" />',
                '        <t:FieldURI FieldURI="item:Body" />',
                '      </t:AdditionalProperties>',
                '    </m:ItemShape>',
                '    <m:ItemIds>',
            ];
            for (var i = 0; i < messages.length; i++) {
                builder.push("      <t:ItemId Id=\"" + messages[i].itemId.id + "\" ChangeKey=\"" + messages[i].itemId.changeKey + "\" />");
            }
            builder.push('    </m:ItemIds>', '  </m:GetItem>', RequestBuilder.endRequest);
            return builder.join('\n');
        };
        RequestBuilder.getFolderNamesRequest = function (folders) {
            var builder = [
                RequestBuilder.beginRequest,
                '  <m:GetFolder>',
                '    <m:FolderShape>',
                '      <t:BaseShape>IdOnly</t:BaseShape>',
                '      <t:AdditionalProperties>',
                '        <t:FieldURI FieldURI="folder:DisplayName" />',
                '      </t:AdditionalProperties>',
                '    </m:FolderShape>',
                '    <m:FolderIds>'
            ];
            for (var i = 0; i < folders.length; i++) {
                builder.push("      <t:FolderId Id=\"" + folders[i].folderId.id + "\" ChangeKey=\"" + folders[i].folderId.changeKey + "\" />");
            }
            builder.push('    </m:FolderIds>', '  </m:GetFolder >', RequestBuilder.endRequest);
            return builder.join('\n');
        };
        RequestBuilder.moveItemsRequest = function (messages, folderId) {
            var builder = [
                RequestBuilder.beginRequest,
                '  <m:MoveItem>',
                '    <m:ToFolderId>',
                '      <t:FolderId Id="' + folderId + '"/>',
                '    </m:ToFolderId>',
                '    <m:ItemIds>',
            ];
            for (var i = 0; i < messages.length; i++) {
                builder.push("      <t:ItemId Id=\"" + messages[i].itemId.id + "\" ChangeKey=\"" + messages[i].itemId.changeKey + "\" />");
            }
            builder.push('    </m:ItemIds>', '  </m:MoveItem>', RequestBuilder.endRequest);
            return builder.join('\n');
        };
        return RequestBuilder;
    }());
    RequestBuilder.beginRequest = [
        '<?xml version="1.0" encoding="utf-8" ?>',
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"',
        '    xmlns:xsd="http://www.w3.org/2001/XMLSchema"',
        '    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"',
        '    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"',
        '    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">',
        '<soap:Header>',
        '  <t:RequestServerVersion Version="Exchange2010_SP1" />',
        '</soap:Header>',
        '<soap:Body>',
    ].join('\n');
    RequestBuilder.endRequest = [
        '</soap:Body>',
        '</soap:Envelope>'
    ].join('\n');
    RequestBuilder.findConversationRequest = [
        RequestBuilder.beginRequest,
        '  <m:FindConversation>',
        '    <m:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="20" />',
        '    <m:ParentFolderId>',
        '      <t:DistinguishedFolderId Id="inbox"/>',
        '    </m:ParentFolderId>',
        '  </m:FindConversation>',
        RequestBuilder.endRequest
    ].join('\n');
    RequestBuilder.excludedFolderIdsRequest = [
        RequestBuilder.beginRequest,
        '  <m:GetFolder>',
        '    <m:FolderShape>',
        '      <t:BaseShape>IdOnly</t:BaseShape>',
        '    </m:FolderShape>',
        '    <m:FolderIds>',
        '      <t:DistinguishedFolderId Id="inbox"/>',
        '      <t:DistinguishedFolderId Id="drafts"/>',
        '      <t:DistinguishedFolderId Id="sentitems"/>',
        '      <t:DistinguishedFolderId Id="deleteditems"/>',
        '    </m:FolderIds>',
        '  </m:GetFolder >',
        RequestBuilder.endRequest
    ].join('\n');
    var Context = (function () {
        function Context(mailbox) {
            this.mailbox = mailbox;
        }
        Context.prototype.loadItems = function (onLoadComplete, onProgress, onError) {
            var _this = this;
            this.onLoadComplete = onLoadComplete;
            this.onProgress = onProgress;
            this.onError = onError;
            console.log('Finding the conversation with EWS...');
            this.onProgress(Model_1.Data.Progress.GetConversation);
            this.mailbox.makeEwsRequestAsync(RequestBuilder.findConversationRequest, function (result) {
                if (!result.value) {
                    _this.onError(result.error.message);
                    return;
                }
                _this.conversationXml = $.parseXML(result.value);
                _this.getConversation();
            });
        };
        Context.prototype.getConversation = function () {
            var $conversation = $(this.conversationXml.querySelectorAll('Conversation > GlobalItemIds > ItemId'))
                .filter("[Id=\"" + this.mailbox.item.itemId + "\"]")
                .parents('Conversation');
            if (!$conversation.length) {
                this.onError("This conversation isn't in your inbox's top 20.");
                return;
            }
            var messageCount = parseInt($conversation.find('MessageCount').text());
            var globalCount = parseInt($conversation.find('GlobalMessageCount').text());
            if (messageCount >= globalCount) {
                this.onLoadComplete([]);
                return;
            }
            var sameFolderItemIds = [];
            $conversation.find('ItemIds > ItemId').each(function () {
                var $this = $(this);
                sameFolderItemIds.push({
                    id: $this.attr('Id'),
                    changeKey: $this.attr('ChangeKey')
                });
            });
            var $otherFolderItemIds = $conversation.find('GlobalItemIds > ItemId');
            for (var i = 0; i < sameFolderItemIds.length; i++) {
                $otherFolderItemIds = $otherFolderItemIds.filter("[Id!=\"" + sameFolderItemIds[i].id + "\"]");
            }
            var otherFolderItemIds = [];
            $otherFolderItemIds.each(function () {
                var $this = $(this);
                otherFolderItemIds.push({
                    id: $this.attr('Id'),
                    changeKey: $this.attr('ChangeKey')
                });
            });
            if (!sameFolderItemIds.length || !otherFolderItemIds.length) {
                this.onLoadComplete([]);
                return;
            }
            this.conversation = {
                id: this.mailbox.item.conversationId,
                items: [],
                global: []
            };
            for (var i_1 = 0; i_1 < sameFolderItemIds.length; i_1++) {
                this.conversation.items.push({
                    itemId: sameFolderItemIds[i_1],
                    conversation: this.conversation
                });
            }
            for (var i = 0; i < otherFolderItemIds.length; i++) {
                this.conversation.global.push({
                    itemId: otherFolderItemIds[i],
                    conversation: this.conversation
                });
            }
            this.loadExcludedFolders();
        };
        Context.prototype.loadExcludedFolders = function () {
            var _this = this;
            if (this.excludedFolders) {
                this.loadMessages();
            }
            else {
                console.log('Getting the list of excluded folders');
                this.onProgress(Model_1.Data.Progress.GetExcludedFolders);
                this.mailbox.makeEwsRequestAsync(RequestBuilder.excludedFolderIdsRequest, function (result) {
                    if (!result.value) {
                        _this.onError(result.error.message);
                        return;
                    }
                    var foldersXml = $.parseXML(result.value);
                    var excludedFolders = [];
                    $(foldersXml.querySelectorAll('GetFolderResponseMessage > Folders > Folder > FolderId')).each(function () {
                        var $this = $(this);
                        excludedFolders.push({
                            folderId: {
                                id: $this.attr('Id'),
                                changeKey: $this.attr('ChangeKey')
                            }
                        });
                    });
                    _this.excludedFolders = excludedFolders;
                    _this.loadMessages();
                });
            }
        };
        Context.prototype.loadMessages = function () {
            var _this = this;
            if (this.itemsXml) {
                this.getMessages();
            }
            else {
                console.log("Getting the messages in other folders: " + this.conversation.global.length);
                this.onProgress(Model_1.Data.Progress.GetConversation);
                this.mailbox.makeEwsRequestAsync(RequestBuilder.getItemsRequest(this.conversation.global), function (result) {
                    if (!result.value) {
                        _this.onError(result.error.message);
                        return;
                    }
                    _this.itemsXml = $.parseXML(result.value);
                    _this.getMessages();
                });
            }
        };
        Context.prototype.getMessages = function () {
            var $messages = $(this.itemsXml.querySelectorAll('GetItemResponseMessage > Items > Message > ParentFolderId'));
            for (var i = 0; i < this.excludedFolders.length; i++) {
                $messages = $messages.filter("[Id!=\"" + this.excludedFolders[i].folderId.id + "\"]");
            }
            $messages = $messages.parent();
            for (var i = 0; i < this.conversation.global.length; i++) {
                var item = this.conversation.global[i];
                for (var j = 0; j < $messages.length; j++) {
                    var msg = $messages[j];
                    if (msg.querySelector("ItemId[Id=\"" + item.itemId.id + "\"]")) {
                        var folderId = msg.querySelector('ParentFolderId');
                        item.folder = {
                            folderId: {
                                id: folderId.getAttribute('Id'),
                                changeKey: folderId.getAttribute('ChangeKey')
                            }
                        };
                        item.from = msg.querySelector('Sender > Mailbox > Name').textContent;
                        item.to = msg.querySelector('ToRecipients > Mailbox > Name').textContent;
                        item.body = msg.querySelector('Body').textContent.slice(0, 200);
                        break;
                    }
                }
            }
            this.loadFolderDisplayNames();
        };
        Context.prototype.loadFolderDisplayNames = function () {
            var _this = this;
            if (this.folderNamesXml) {
                this.getFolderDisplayNames();
            }
            else {
                var destinations = [];
                for (var i = 0; i < this.conversation.global.length; i++) {
                    var item = this.conversation.global[i];
                    if (item.folder) {
                        destinations.push(item.folder);
                    }
                }
                if (!destinations.length) {
                    this.onLoadComplete([]);
                    return;
                }
                console.log("Getting the display names of the other folders: " + destinations.length);
                this.onProgress(Model_1.Data.Progress.GetFolderNames);
                this.mailbox.makeEwsRequestAsync(RequestBuilder.getFolderNamesRequest(destinations), function (result) {
                    if (!result.value) {
                        _this.onError(result.error.message);
                        return;
                    }
                    _this.folderNamesXml = $.parseXML(result.value);
                    _this.getFolderDisplayNames();
                });
            }
        };
        Context.prototype.getFolderDisplayNames = function () {
            var _this = this;
            var matches = [];
            this.conversation.global.map(function (item) {
                if (!item.folder) {
                    return;
                }
                var folder = _this.folderNamesXml.querySelector("GetFolderResponseMessage > Folders > Folder > FolderId[Id=\"" + item.folder.folderId.id + "\"]").parentNode;
                item.folder.displayName = folder.querySelector('DisplayName').textContent;
                matches.push({
                    message: {
                        Id: item.itemId.id,
                        BodyPreview: item.body,
                        Sender: item.from,
                        ToRecipients: item.to,
                        ParentFolderId: item.folder.folderId.id
                    },
                    folder: {
                        Id: item.folder.folderId.id,
                        DisplayName: item.folder.displayName
                    }
                });
            });
            console.log("Finished loading items in other folders: " + matches.length);
            this.onLoadComplete(matches);
        };
        Context.prototype.moveItems = function (folderId, onMoveComplete, onError) {
            var _this = this;
            this.onMoveComplete = onMoveComplete;
            this.onError = onError;
            console.log("Moving items to folder: " + folderId);
            this.mailbox.makeEwsRequestAsync(RequestBuilder.moveItemsRequest(this.conversation.items, folderId), function (result) {
                if (!result.value) {
                    _this.onError(result.error.message);
                    return;
                }
                console.log("Finished moving items to other folder: " + _this.conversation.items.length);
                _this.onMoveComplete(_this.conversation.items.length);
            });
        };
        return Context;
    }());
    var Model = (function () {
        function Model(mailbox) {
            this.context = new Context(mailbox);
        }
        Model.prototype.getItemsAsync = function (onLoadComplete, onProgress, onError) {
            this.context.loadItems(onLoadComplete, onProgress, onError);
        };
        Model.prototype.moveItemsAsync = function (folderId, onMoveComplete, onError) {
            this.context.moveItems(folderId, onMoveComplete, onError);
        };
        return Model;
    }());
    EWSData.Model = Model;
})(EWSData = exports.EWSData || (exports.EWSData = {}));


/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/// <reference path="../_references.ts" />

Object.defineProperty(exports, "__esModule", { value: true });
var Model_1 = __webpack_require__(1);
var RESTData;
(function (RESTData) {
    var Endpoint = '/v2.0/me';
    var ExcludedFolders;
    (function (ExcludedFolders) {
        ExcludedFolders[ExcludedFolders["Inbox"] = 0] = "Inbox";
        ExcludedFolders[ExcludedFolders["Drafts"] = 1] = "Drafts";
        ExcludedFolders[ExcludedFolders["SentItems"] = 2] = "SentItems";
        ExcludedFolders[ExcludedFolders["DeletedItems"] = 3] = "DeletedItems";
        // Sentinel value for enumerating the folder names
        ExcludedFolders[ExcludedFolders["Count"] = 4] = "Count";
    })(ExcludedFolders || (ExcludedFolders = {}));
    var Context = (function () {
        function Context(mailbox) {
            this.mailbox = mailbox;
        }
        Context.prototype.loadItems = function (onLoadComplete, onProgress, onError) {
            var _this = this;
            this.onLoadComplete = onLoadComplete;
            this.onProgress = onProgress;
            this.onError = onError;
            console.log('Requesting the REST callback token...');
            this.onProgress(Model_1.Data.Progress.GetCallbackToken);
            // Start the chain of requests by getting a callback token.
            this.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    _this.getConversation(result);
                }
                else {
                    _this.onError(result.error.message);
                }
            });
        };
        // Sometimes we need to make separate REST requests for multiple items. Wait until they all complete and then
        // invoke the callbacks all at once with an array of typed results.
        Context.prototype.collateRequests = function (requests, onDone, onFail) {
            var _this = this;
            if (requests.length > 1) {
                $.when.apply($, requests)
                    .done(function () {
                    var results = [];
                    for (var _i = 0; _i < arguments.length; _i++) {
                        results[_i] = arguments[_i];
                    }
                    var values = [];
                    results.map(function (result) {
                        values.push(result[0]);
                    });
                    onDone(values);
                }).fail(function (message) {
                    _this.onError(message);
                });
            }
            else {
                requests[0]
                    .done(function (result) {
                    onDone([result]);
                }).fail(function (message) {
                    _this.onError(message);
                });
            }
        };
        // Send a REST request to retrieve a list of messages in this conversation.
        Context.prototype.getConversation = function (result) {
            var _this = this;
            this.token = result.value;
            var conversationId = this.mailbox.item.conversationId;
            var restConversationId = this.mailbox.diagnostics.hostName === 'OutlookIOS'
                ? conversationId
                : this.mailbox.convertToRestId(conversationId, Office.MailboxEnums.RestVersion.v2_0);
            var restUrl = "" + this.mailbox.restUrl + Endpoint + "/messages?$filter=ConversationId eq '" + restConversationId + "'&$select=Id,Subject,BodyPreview,Sender,ToRecipients,ParentFolderId";
            console.log("Getting the list of items in the conversation: " + restUrl);
            this.onProgress(Model_1.Data.Progress.GetConversation);
            $.ajax({
                url: restUrl,
                async: true,
                dataType: 'json',
                headers: { 'Authorization': "Bearer " + this.token }
            }).done(function (result) {
                _this.getExcludedFolders(result);
            }).fail(function (message) {
                _this.onError(message);
            });
        };
        // Send a REST request to identify each of the folders we want to exclude in our results.
        Context.prototype.getExcludedFolders = function (result) {
            var _this = this;
            if (!result || !result.value || 0 === result.value.length) {
                this.onLoadComplete([]);
                return;
            }
            this.conversationMessages = result.value;
            var currentFolderId;
            var excludedFolderIds = [];
            // We should ignore any messages in the same folder.
            var itemId = this.mailbox.item.itemId;
            var restItemId = this.mailbox.diagnostics.hostName === 'OutlookIOS'
                ? itemId
                : this.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
            for (var i = 0; i < this.conversationMessages.length; ++i) {
                if (this.conversationMessages[i].Id === restItemId) {
                    currentFolderId = this.conversationMessages[i].ParentFolderId;
                    excludedFolderIds.push(currentFolderId);
                    break;
                }
            }
            // We should also exclude some special folders, but we need to get their folderIds.
            var requests = [];
            for (var i = 0; i < ExcludedFolders.Count; ++i) {
                var folderId = ExcludedFolders[i];
                var restUrl = "" + this.mailbox.restUrl + Endpoint + "/mailfolders/" + folderId + "?$select=Id";
                console.log("Getting excluded folder ID: " + restUrl);
                requests.push($.ajax({
                    url: restUrl,
                    async: true,
                    dataType: 'json',
                    headers: { 'Authorization': "Bearer " + this.token }
                }));
            }
            this.onProgress(Model_1.Data.Progress.GetExcludedFolders);
            this.collateRequests(requests, function (results) {
                results.map(function (value) {
                    excludedFolderIds.push(value.Id);
                });
                _this.getFolderNames(currentFolderId, excludedFolderIds);
            }, function (message) {
                _this.onError(message);
            });
        };
        // Send REST requests to fill in the display names of all the folders we are not excluding.
        Context.prototype.getFolderNames = function (currentFolderId, excludedFolderIds) {
            var _this = this;
            var folderMap = [];
            this.conversationMessages.map(function (message) {
                for (var i = 0; i < excludedFolderIds.length; ++i) {
                    if (excludedFolderIds[i] === message.ParentFolderId) {
                        // Skip this message.
                        return;
                    }
                }
                for (var i = 0; i < folderMap.length; ++i) {
                    if (folderMap[i].folder.Id === message.ParentFolderId) {
                        // Add this message to the existing entry.
                        folderMap[i].messages.push(message);
                        return;
                    }
                }
                // Create a new entry for this folder.
                folderMap.push({ folder: { Id: message.ParentFolderId }, messages: [message] });
            });
            if (folderMap.length === 0) {
                this.onLoadComplete([]);
                return;
            }
            this.currentFolderId = currentFolderId;
            this.excludedFolderIds = excludedFolderIds;
            var requests = [];
            folderMap.map(function (entry) {
                var restUrl = "" + _this.mailbox.restUrl + Endpoint + "/mailfolders/" + entry.folder.Id + "?$select=Id,DisplayName";
                console.log("Getting included folder name: " + restUrl);
                requests.push($.ajax({
                    url: restUrl,
                    async: true,
                    dataType: 'json',
                    headers: { 'Authorization': "Bearer " + _this.token }
                }));
            });
            this.onProgress(Model_1.Data.Progress.GetFolderNames);
            this.collateRequests(requests, function (results) {
                results.map(function (value) {
                    for (var i = 0; i < folderMap.length; ++i) {
                        if (folderMap[i].folder.Id === value.Id) {
                            folderMap[i].folder.DisplayName = value.DisplayName;
                            break;
                        }
                    }
                });
                var matches = [];
                folderMap.map(function (entry) {
                    entry.messages.map(function (message) {
                        var recipients = [];
                        message.ToRecipients.map(function (address) {
                            recipients.push(address.EmailAddress.Name);
                        });
                        var value = {
                            Id: message.Id,
                            BodyPreview: message.BodyPreview,
                            Sender: message.Sender.EmailAddress.Name,
                            ToRecipients: recipients.join('; '),
                            ParentFolderId: message.ParentFolderId
                        };
                        matches.push({
                            message: {
                                Id: message.Id,
                                BodyPreview: message.BodyPreview,
                                Sender: message.Sender.EmailAddress.Name,
                                ToRecipients: recipients.join('; '),
                                ParentFolderId: message.ParentFolderId
                            },
                            folder: {
                                Id: entry.folder.Id,
                                DisplayName: entry.folder.DisplayName
                            }
                        });
                    });
                });
                console.log("Finished loading items in other folders: " + matches.length);
                _this.onLoadComplete(matches);
            }, function (message) {
                _this.onError(message);
            });
        };
        Context.prototype.moveItems = function (folderId, onMoveComplete, onError) {
            var _this = this;
            this.onMoveComplete = onMoveComplete;
            this.onError = onError;
            console.log("Moving items to folder: " + folderId);
            var requests = [];
            this.conversationMessages.map(function (message) {
                if (message.ParentFolderId !== _this.currentFolderId) {
                    // Skip any messages that are not in the current folder.
                    return;
                }
                var restUrl = "" + _this.mailbox.restUrl + Endpoint + "/messages/" + message.Id + "/move";
                console.log("Moving item: " + restUrl);
                requests.push($.ajax({
                    url: restUrl,
                    async: true,
                    method: 'POST',
                    contentType: 'application/json',
                    dataType: 'json',
                    data: JSON.stringify({ DestinationId: folderId }),
                    headers: { 'Authorization': "Bearer " + _this.token }
                }));
            });
            this.collateRequests(requests, function (results) {
                console.log("Finished moving items to other folder: " + results.length);
                _this.onMoveComplete(results.length);
            }, function (message) {
                _this.onError(message);
            });
        };
        return Context;
    }());
    var Model = (function () {
        function Model(mailbox) {
            this.context = new Context(mailbox);
        }
        Model.prototype.getItemsAsync = function (onLoadComplete, onProgress, onError) {
            this.context.loadItems(onLoadComplete, onProgress, onError);
        };
        Model.prototype.moveItemsAsync = function (folderId, onMoveComplete, onError) {
            this.context.moveItems(folderId, onMoveComplete, onError);
        };
        return Model;
    }());
    RESTData.Model = Model;
})(RESTData = exports.RESTData || (exports.RESTData = {}));


/***/ }),
/* 6 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var React = __webpack_require__(0);
var Feedback = (function (_super) {
    __extends(Feedback, _super);
    function Feedback() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Feedback.prototype.render = function () {
        return (React.createElement("div", { className: "feedback" },
            React.createElement("a", { href: "https://beandotnet.azurewebsites.net/" }, "about this app"),
            "\u00A0",
            React.createElement("a", { href: "mailto:wravery@hotmail.com?Subject=Auto%20Filer%20App%20for%20Outlook" }, "send feedback")));
    };
    return Feedback;
}(React.Component));
exports.Feedback = Feedback;


/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var React = __webpack_require__(0);
var SearchResults = (function (_super) {
    __extends(SearchResults, _super);
    function SearchResults(props) {
        var _this = _super.call(this, props) || this;
        _this.onClickFolder = _this.handleClick.bind(_this);
        return _this;
    }
    SearchResults.prototype.render = function () {
        var _this = this;
        if (!this.props.matches || this.props.matches.length === 0) {
            return null;
        }
        var rows = [];
        this.props.matches.map(function (value, index) {
            rows.push(React.createElement("tr", { key: index },
                React.createElement("td", null,
                    React.createElement("a", { name: value.folder.Id, onClick: _this.onClickFolder }, value.folder.DisplayName)),
                React.createElement("td", null, value.message.Sender),
                React.createElement("td", null, value.message.ToRecipients),
                React.createElement("td", null, value.message.BodyPreview)));
        });
        return (React.createElement("table", null,
            React.createElement("thead", null,
                React.createElement("tr", null,
                    React.createElement("th", null, "Folder"),
                    React.createElement("th", null, "From"),
                    React.createElement("th", null, "To"),
                    React.createElement("th", null, "Preview"))),
            React.createElement("tbody", null, rows)));
    };
    SearchResults.prototype.handleClick = function (evt) {
        this.props.onSelection(evt.currentTarget.name);
        evt.preventDefault();
    };
    return SearchResults;
}(React.Component));
exports.SearchResults = SearchResults;


/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var React = __webpack_require__(0);
var Model_1 = __webpack_require__(1);
var StatusMessage = (function (_super) {
    __extends(StatusMessage, _super);
    function StatusMessage() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    StatusMessage.prototype.render = function () {
        var className;
        var status;
        switch (this.props.progress) {
            case Model_1.Data.Progress.GetCallbackToken:
            case Model_1.Data.Progress.GetConversation:
            case Model_1.Data.Progress.GetExcludedFolders:
            case Model_1.Data.Progress.GetFolderNames:
                return React.createElement("h3", null, "Looking for other messages in this conversation...");
            case Model_1.Data.Progress.Success:
                return null;
            case Model_1.Data.Progress.NotFound:
                return React.createElement("h3", null, "It looks like you haven't filed this conversation anywhere before.");
            default:
                return (React.createElement("div", null,
                    React.createElement("h3", null, "Sorry, I couldn't figure out where this message should go. :("),
                    React.createElement("span", null, this.props.message)));
        }
    };
    return StatusMessage;
}(React.Component));
exports.StatusMessage = StatusMessage;


/***/ }),
/* 9 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";
/// <reference path="_references.ts" />
/// <reference path="./components/ConversationFiler.tsx" />

Object.defineProperty(exports, "__esModule", { value: true });
var React = __webpack_require__(0);
var ReactDOM = __webpack_require__(3);
var ConversationFiler_1 = __webpack_require__(2);
Office.initialize = function () {
    var functionsRegex = /functions\.html(\?.*)?$/i;
    var noUI = functionsRegex.test(window.location.pathname);
    if (noUI) {
        // Add the UI-less function callback if we're loaded from functions.html instead of index.html
        window.fileDialog = function (event) {
            Office.context.ui.displayDialogAsync(window.location.href.replace(functionsRegex, "dialog.html"), { height: 25, width: 80, displayInIframe: true }, function (result) {
                var dialog = result.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function () {
                    dialog.close();
                    event.completed();
                });
            });
        };
        return;
    }
    // Show the UI...
    var mailbox = (Office.context || {}).mailbox;
    var onComplete;
    if (mailbox && /dialog\.html(\?.*)?$/i.test(window.location.pathname)) {
        // When we finish moving the items, we want to dismiss the dialog and complete the callback
        onComplete = function () {
            Office.context.ui.messageParent(true);
        };
    }
    ReactDOM.render(React.createElement(ConversationFiler_1.ConversationFiler, { mailbox: mailbox, onComplete: onComplete }), document.getElementById("conversationFilerRoot"));
    // ...and if we're running outside of an Outlook client, run through the tests
    if (!mailbox) {
        var testEmpty = function () {
            console.log("Testing the behavior with an empty set of matches...");
            // Need to clear out the DOM so it will mount a new ConversationFiler
            ReactDOM.render(React.createElement("div", null, "Testing..."), document.getElementById("conversationFilerRoot"));
            ReactDOM.render(React.createElement(ConversationFiler_1.ConversationFiler, { mailbox: null, mockResults: [] }), document.getElementById("conversationFilerRoot"));
            window.setTimeout(testDummy_1, 3000);
        };
        var testDummy_1 = function () {
            console.log("Testing the behavior with a set of mock matches...");
            // Need to clear out the DOM so it will mount a new ConversationFiler
            ReactDOM.render(React.createElement("div", null, "Testing..."), document.getElementById("conversationFilerRoot"));
            var mockResults = [{
                    folder: {
                        Id: 'folderId1',
                        DisplayName: 'Folder 1'
                    },
                    message: {
                        Id: 'messageId1',
                        BodyPreview: 'Here\'s a preview of a message body',
                        Sender: 'Foo Bar',
                        ToRecipients: 'Baz Bar',
                        ParentFolderId: 'folderId1'
                    }
                }, {
                    folder: {
                        Id: 'folderId2',
                        DisplayName: 'Folder 2'
                    },
                    message: {
                        Id: 'messageId2',
                        BodyPreview: 'Here\'s another message body',
                        Sender: 'Baz Bar',
                        ToRecipients: 'Foo Bar',
                        ParentFolderId: 'folderId2'
                    }
                }];
            ReactDOM.render(React.createElement(ConversationFiler_1.ConversationFiler, { mailbox: null, mockResults: mockResults }), document.getElementById("conversationFilerRoot"));
        };
        window.setTimeout(testEmpty, 3000);
    }
};


/***/ })
/******/ ]);
//# sourceMappingURL=bundle.js.map