!function(e){function t(n){if(o[n])return o[n].exports;var r=o[n]={i:n,l:!1,exports:{}};return e[n].call(r.exports,r,r.exports,t),r.l=!0,r.exports}var o={};t.m=e,t.c=o,t.i=function(e){return e},t.d=function(e,o,n){t.o(e,o)||Object.defineProperty(e,o,{configurable:!1,enumerable:!0,get:n})},t.n=function(e){var o=e&&e.__esModule?function(){return e.default}:function(){return e};return t.d(o,"a",o),o},t.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},t.p="",t(t.s=10)}([function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});!function(e){!function(e){e[e.GetCallbackToken=0]="GetCallbackToken",e[e.GetConversation=1]="GetConversation",e[e.GetExcludedFolders=2]="GetExcludedFolders",e[e.GetFolderNames=3]="GetFolderNames",e[e.Success=4]="Success",e[e.NotFound=5]="NotFound",e[e.Error=6]="Error"}(e.Progress||(e.Progress={}))}(t.Data||(t.Data={}))},function(e,t){e.exports=React},function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n=o(6),r=o(5);!function(e){function t(e){return e.restUrl?new n.RESTData.Model(e):new r.EWSData.Model(e)}e.getData=t}(t.Factory||(t.Factory={}))},function(e,t,o){"use strict";var n=this&&this.__extends||function(){var e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var o in t)t.hasOwnProperty(o)&&(e[o]=t[o])};return function(t,o){function n(){this.constructor=t}e(t,o),t.prototype=null===o?Object.create(o):(n.prototype=o.prototype,new n)}}();Object.defineProperty(t,"__esModule",{value:!0});var r=o(1),s=o(0),i=o(2),a=o(9),l=o(8),d=o(7),c=function(e){function t(t){var o=e.call(this,t)||this;return o.state={progress:s.Data.Progress.GetCallbackToken},o}return n(t,e),t.prototype.componentDidMount=function(){var e=this;if(this.props.storedResults)return void(this.props.storedResults.length>0?this.setState({progress:s.Data.Progress.Success,matches:this.props.storedResults}):this.setState({progress:s.Data.Progress.NotFound}));if(this.props.mailbox){var t=i.Factory.getData(this.props.mailbox);this.setState({data:t}),t.getItemsAsync(function(t){t.length>0?e.setState({progress:s.Data.Progress.Success,matches:t}):e.setState({progress:s.Data.Progress.NotFound})},function(t){e.setState({progress:t})},function(t){e.setState({progress:s.Data.Progress.Error,error:t})})}},t.prototype.onSelection=function(e){var t=this;if(console.log("Selected a folder: "+e),!this.state.data)return void(this.props.onComplete&&this.props.onComplete(e));this.state.data.moveItemsAsync(e,function(o){t.props.onComplete&&t.props.onComplete(e)},function(e){t.setState({progress:s.Data.Progress.Error,error:e})})},t.prototype.render=function(){return r.createElement("div",null,r.createElement(a.StatusMessage,{progress:this.state.progress,message:this.state.error}),r.createElement(l.SearchResults,{matches:this.state.matches,onSelection:this.onSelection.bind(this)}),r.createElement(d.Feedback,null))},t}(r.Component);t.ConversationFiler=c},function(e,t){e.exports=ReactDOM},function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n=o(0);!function(e){var t=function(){function e(){}return e.getItemsRequest=function(t){var o=[e.beginRequest,"  <m:GetItem>","    <m:ItemShape>","      <t:BaseShape>IdOnly</t:BaseShape>","      <t:BodyType>Text</t:BodyType>","      <t:AdditionalProperties>",'        <t:FieldURI FieldURI="item:ParentFolderId" />','        <t:FieldURI FieldURI="message:Sender" />','        <t:FieldURI FieldURI="message:ToRecipients" />','        <t:FieldURI FieldURI="item:Body" />',"      </t:AdditionalProperties>","    </m:ItemShape>","    <m:ItemIds>"];return t.map(function(e){o.push('      <t:ItemId Id="'+e.itemId.id+'" ChangeKey="'+e.itemId.changeKey+'" />')}),o.push("    </m:ItemIds>","  </m:GetItem>",e.endRequest),o.join("\n")},e.getFolderNamesRequest=function(t){var o=[e.beginRequest,"  <m:GetFolder>","    <m:FolderShape>","      <t:BaseShape>IdOnly</t:BaseShape>","      <t:AdditionalProperties>",'        <t:FieldURI FieldURI="folder:DisplayName" />',"      </t:AdditionalProperties>","    </m:FolderShape>","    <m:FolderIds>"];return t.map(function(e){o.push('      <t:FolderId Id="'+e.folderId.id+'" ChangeKey="'+e.folderId.changeKey+'" />')}),o.push("    </m:FolderIds>","  </m:GetFolder >",e.endRequest),o.join("\n")},e.moveItemsRequest=function(t,o){var n=[e.beginRequest,"  <m:MoveItem>","    <m:ToFolderId>",'      <t:FolderId Id="'+o+'"/>',"    </m:ToFolderId>","    <m:ItemIds>"];return t.map(function(e){n.push('      <t:ItemId Id="'+e.itemId.id+'" ChangeKey="'+e.itemId.changeKey+'" />')}),n.push("    </m:ItemIds>","  </m:MoveItem>",e.endRequest),n.join("\n")},e}();t.beginRequest=['<?xml version="1.0" encoding="utf-8" ?>','<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"','    xmlns:xsd="http://www.w3.org/2001/XMLSchema"','    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"','    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"','    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">',"<soap:Header>",'  <t:RequestServerVersion Version="Exchange2010_SP1" />',"</soap:Header>","<soap:Body>"].join("\n"),t.endRequest=["</soap:Body>","</soap:Envelope>"].join("\n"),t.findConversationRequest=[t.beginRequest,"  <m:FindConversation>",'    <m:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="20" />',"    <m:ParentFolderId>",'      <t:DistinguishedFolderId Id="inbox"/>',"    </m:ParentFolderId>","  </m:FindConversation>",t.endRequest].join("\n"),t.excludedFolderIdsRequest=[t.beginRequest,"  <m:GetFolder>","    <m:FolderShape>","      <t:BaseShape>IdOnly</t:BaseShape>","    </m:FolderShape>","    <m:FolderIds>",'      <t:DistinguishedFolderId Id="inbox"/>','      <t:DistinguishedFolderId Id="drafts"/>','      <t:DistinguishedFolderId Id="sentitems"/>','      <t:DistinguishedFolderId Id="deleteditems"/>',"    </m:FolderIds>","  </m:GetFolder >",t.endRequest].join("\n");var o=function(){function e(e){this.mailbox=e}return e.prototype.loadItems=function(e,o,r){var s=this;this.onLoadComplete=e,this.onProgress=o,this.onError=r,console.log("Finding the conversation with EWS..."),this.onProgress(n.Data.Progress.GetConversation),this.mailbox.makeEwsRequestAsync(t.findConversationRequest,function(e){if(!e.value)return void s.onError(e.error.message);s.conversationXml=$.parseXML(e.value),s.getConversation()})},e.prototype.getConversation=function(){var e=this,t=$(this.conversationXml.querySelectorAll("Conversation > GlobalItemIds > ItemId")).filter('[Id="'+this.mailbox.item.itemId+'"]').parents("Conversation");if(!t.length)return void this.onError("This conversation isn't in your inbox's top 20.");if(parseInt(t.find("MessageCount").text())>=parseInt(t.find("GlobalMessageCount").text()))return void this.onLoadComplete([]);var o=[];t.find("ItemIds > ItemId").each(function(){var e=$(this);o.push({id:e.attr("Id"),changeKey:e.attr("ChangeKey")})});var n=t.find("GlobalItemIds > ItemId");o.map(function(e){n=n.filter('[Id!="'+e.id+'"]')});var r=[];if(n.each(function(){var e=$(this);r.push({id:e.attr("Id"),changeKey:e.attr("ChangeKey")})}),!o.length||!r.length)return void this.onLoadComplete([]);this.conversation={id:this.mailbox.item.conversationId,items:o.map(function(t){return{itemId:t,conversation:e.conversation}}),global:r.map(function(t){return{itemId:t,conversation:e.conversation}})},this.loadExcludedFolders()},e.prototype.loadExcludedFolders=function(){var e=this;this.excludedFolders?this.loadMessages():(console.log("Getting the list of excluded folders"),this.onProgress(n.Data.Progress.GetExcludedFolders),this.mailbox.makeEwsRequestAsync(t.excludedFolderIdsRequest,function(t){if(!t.value)return void e.onError(t.error.message);var o=$.parseXML(t.value),n=[];$(o.querySelectorAll("GetFolderResponseMessage > Folders > Folder > FolderId")).each(function(){var e=$(this);n.push({folderId:{id:e.attr("Id"),changeKey:e.attr("ChangeKey")}})}),e.excludedFolders=n,e.loadMessages()}))},e.prototype.loadMessages=function(){var e=this;this.itemsXml?this.getMessages():(console.log("Getting the messages in other folders: "+this.conversation.global.length),this.onProgress(n.Data.Progress.GetConversation),this.mailbox.makeEwsRequestAsync(t.getItemsRequest(this.conversation.global),function(t){if(!t.value)return void e.onError(t.error.message);e.itemsXml=$.parseXML(t.value),e.getMessages()}))},e.prototype.getMessages=function(){var e=$(this.itemsXml.querySelectorAll("GetItemResponseMessage > Items > Message > ParentFolderId"));this.excludedFolders.map(function(t){e=e.filter('[Id!="'+t.folderId.id+'"]')}),e=e.parent(),this.conversation.global.map(function(t){for(var o=0;o<e.length;o++){var n=e[o];if(n.querySelector('ItemId[Id="'+t.itemId.id+'"]')){var r=n.querySelector("ParentFolderId");t.folder={folderId:{id:r.getAttribute("Id"),changeKey:r.getAttribute("ChangeKey")}},t.from=n.querySelector("Sender > Mailbox > Name").textContent,t.to=n.querySelector("ToRecipients > Mailbox > Name").textContent,t.body=n.querySelector("Body").textContent.slice(0,200);break}}}),this.loadFolderDisplayNames()},e.prototype.loadFolderDisplayNames=function(){var e=this;if(this.folderNamesXml)this.getFolderDisplayNames();else{var o=[];if(this.conversation.global.map(function(e){e.folder&&o.push(e.folder)}),!o.length)return void this.onLoadComplete([]);console.log("Getting the display names of the other folders: "+o.length),this.onProgress(n.Data.Progress.GetFolderNames),this.mailbox.makeEwsRequestAsync(t.getFolderNamesRequest(o),function(t){if(!t.value)return void e.onError(t.error.message);e.folderNamesXml=$.parseXML(t.value),e.getFolderDisplayNames()})}},e.prototype.getFolderDisplayNames=function(){var e=this,t=[];this.conversation.global.map(function(o){if(o.folder){var n=e.folderNamesXml.querySelector('GetFolderResponseMessage > Folders > Folder > FolderId[Id="'+o.folder.folderId.id+'"]').parentNode;o.folder.displayName=n.querySelector("DisplayName").textContent,t.push({message:{Id:o.itemId.id,BodyPreview:o.body,Sender:o.from,ToRecipients:o.to,ParentFolderId:o.folder.folderId.id},folder:{Id:o.folder.folderId.id,DisplayName:o.folder.displayName}})}}),console.log("Finished loading items in other folders: "+t.length),this.onLoadComplete(t)},e.prototype.moveItems=function(e,o,n){var r=this;this.onMoveComplete=o,this.onError=n,console.log("Moving items to folder: "+e),this.mailbox.makeEwsRequestAsync(t.moveItemsRequest(this.conversation.items,e),function(e){if(!e.value)return void r.onError(e.error.message);console.log("Finished moving items to other folder: "+r.conversation.items.length),r.onMoveComplete(r.conversation.items.length)})},e}(),r=function(){function e(e){this.context=new o(e)}return e.prototype.getItemsAsync=function(e,t,o){this.context.loadItems(e,t,o)},e.prototype.moveItemsAsync=function(e,t,o){this.context.moveItems(e,t,o)},e}();e.Model=r}(t.EWSData||(t.EWSData={}))},function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n=o(0);!function(e){var t;!function(e){e[e.Inbox=0]="Inbox",e[e.Drafts=1]="Drafts",e[e.SentItems=2]="SentItems",e[e.DeletedItems=3]="DeletedItems",e[e.Count=4]="Count"}(t||(t={}));var o=function(){function e(e){this.mailbox=e}return e.prototype.loadItems=function(e,t,o){var r=this;this.onLoadComplete=e,this.onProgress=t,this.onError=o,console.log("Requesting the REST callback token..."),this.onProgress(n.Data.Progress.GetCallbackToken),this.mailbox.getCallbackTokenAsync({isRest:!0},function(e){e.status===Office.AsyncResultStatus.Succeeded?r.getConversation(e):r.onError(e.error.message)})},e.prototype.collateRequests=function(e,t,o){var n=this;e.length>1?$.when.apply($,e).done(function(){for(var e=[],o=0;o<arguments.length;o++)e[o]=arguments[o];t(e.map(function(e){return e[0]}))}).fail(function(e){n.onError(e)}):e[0].done(function(e){t([e])}).fail(function(e){n.onError(e)})},e.prototype.getRestId=function(e){return"OutlookIOS"===this.mailbox.diagnostics.hostName?e:this.mailbox.convertToRestId(e,Office.MailboxEnums.RestVersion.v2_0)},e.prototype.getConversation=function(e){var t=this;this.token=e.value;var o=this.mailbox.item.conversationId,r=this.getRestId(o),s=this.mailbox.restUrl+"/v2.0/me/messages?$filter=ConversationId eq '"+r+"'&$select=Id,Subject,BodyPreview,Sender,ToRecipients,ParentFolderId";console.log("Getting the list of items in the conversation: "+s),this.onProgress(n.Data.Progress.GetConversation),$.ajax({url:s,async:!0,dataType:"json",headers:{Authorization:"Bearer "+this.token}}).done(function(e){t.getExcludedFolders(e)}).fail(function(e){t.onError(e)})},e.prototype.getExcludedFolders=function(e){var o=this;if(!e||!e.value||0===e.value.length)return void this.onLoadComplete([]);this.conversationMessages=e.value;for(var r,s=[],i=this.mailbox.item.itemId,a=this.getRestId(i),l=0;l<this.conversationMessages.length;++l)if(this.conversationMessages[l].Id===a){r=this.conversationMessages[l].ParentFolderId,s.push(r);break}for(var d=[],l=0;l<t.Count;++l){var c=t[l],u=this.mailbox.restUrl+"/v2.0/me/mailfolders/"+c+"?$select=Id";console.log("Getting excluded folder ID: "+u),d.push($.ajax({url:u,async:!0,dataType:"json",headers:{Authorization:"Bearer "+this.token}}))}this.onProgress(n.Data.Progress.GetExcludedFolders),this.collateRequests(d,function(e){e.map(function(e){s.push(e.Id)}),o.getFolderNames(r,s)},function(e){o.onError(e)})},e.prototype.getFolderNames=function(e,t){var o=this,r=[];if(this.conversationMessages.map(function(e){for(var o=0;o<t.length;++o)if(t[o]===e.ParentFolderId)return;for(var o=0;o<r.length;++o)if(r[o].folder.Id===e.ParentFolderId)return void r[o].messages.push(e);r.push({folder:{Id:e.ParentFolderId},messages:[e]})}),0===r.length)return void this.onLoadComplete([]);this.currentFolderId=e,this.excludedFolderIds=t;var s=r.map(function(e){var t=o.mailbox.restUrl+"/v2.0/me/mailfolders/"+e.folder.Id+"?$select=Id,DisplayName";return console.log("Getting included folder name: "+t),$.ajax({url:t,async:!0,dataType:"json",headers:{Authorization:"Bearer "+o.token}})});this.onProgress(n.Data.Progress.GetFolderNames),this.collateRequests(s,function(e){e.map(function(e){for(var t=0;t<r.length;++t)if(r[t].folder.Id===e.Id){r[t].folder.DisplayName=e.DisplayName;break}});var t=[];r.map(function(e){e.messages.map(function(o){t.push({message:{Id:o.Id,BodyPreview:o.BodyPreview,Sender:o.Sender.EmailAddress.Name,ToRecipients:o.ToRecipients.map(function(e){return e.EmailAddress.Name}).join("; "),ParentFolderId:o.ParentFolderId},folder:{Id:e.folder.Id,DisplayName:e.folder.DisplayName}})})}),console.log("Finished loading items in other folders: "+t.length),o.onLoadComplete(t)},function(e){o.onError(e)})},e.prototype.moveItems=function(e,t,o){var n=this;this.onMoveComplete=t,this.onError=o,console.log("Moving items to folder: "+e);var r=[];this.conversationMessages.map(function(t){if(t.ParentFolderId===n.currentFolderId){var o=n.mailbox.restUrl+"/v2.0/me/messages/"+t.Id+"/move";console.log("Moving item: "+o),r.push($.ajax({url:o,async:!0,method:"POST",contentType:"application/json",dataType:"json",data:JSON.stringify({DestinationId:e}),headers:{Authorization:"Bearer "+n.token}}))}}),this.collateRequests(r,function(e){console.log("Finished moving items to other folder: "+e.length),n.onMoveComplete(e.length)},function(e){n.onError(e)})},e}(),r=function(){function e(e){this.context=new o(e)}return e.prototype.getItemsAsync=function(e,t,o){this.context.loadItems(e,t,o)},e.prototype.moveItemsAsync=function(e,t,o){this.context.moveItems(e,t,o)},e}();e.Model=r}(t.RESTData||(t.RESTData={}))},function(e,t,o){"use strict";var n=this&&this.__extends||function(){var e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var o in t)t.hasOwnProperty(o)&&(e[o]=t[o])};return function(t,o){function n(){this.constructor=t}e(t,o),t.prototype=null===o?Object.create(o):(n.prototype=o.prototype,new n)}}();Object.defineProperty(t,"__esModule",{value:!0});var r=o(1),s=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return n(t,e),t.prototype.render=function(){return r.createElement("div",{className:"feedback"},r.createElement("a",{href:"https://beandotnet.azurewebsites.net/"},"about this app")," ",r.createElement("a",{href:"mailto:wravery@hotmail.com?Subject=Conversation%20Filer%20v2.0%20App%20for%20Outlook"},"send feedback"))},t}(r.Component);t.Feedback=s},function(e,t,o){"use strict";var n=this&&this.__extends||function(){var e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var o in t)t.hasOwnProperty(o)&&(e[o]=t[o])};return function(t,o){function n(){this.constructor=t}e(t,o),t.prototype=null===o?Object.create(o):(n.prototype=o.prototype,new n)}}();Object.defineProperty(t,"__esModule",{value:!0});var r=o(1),s=function(e){function t(t){var o=e.call(this,t)||this;return o.onClickFolder=o.handleClick.bind(o),o}return n(t,e),t.prototype.render=function(){var e=this;if(!this.props.matches||0===this.props.matches.length)return null;var t=this.props.matches.map(function(t,o){return r.createElement("tr",{key:o},r.createElement("td",null,r.createElement("a",{name:t.folder.Id,onClick:e.onClickFolder},t.folder.DisplayName)),r.createElement("td",null,t.message.Sender),r.createElement("td",null,t.message.ToRecipients),r.createElement("td",null,t.message.BodyPreview))});return r.createElement("table",null,r.createElement("thead",null,r.createElement("tr",null,r.createElement("th",null,"Folder"),r.createElement("th",null,"From"),r.createElement("th",null,"To"),r.createElement("th",null,"Preview"))),r.createElement("tbody",null,t))},t.prototype.handleClick=function(e){this.props.onSelection(e.currentTarget.name),e.preventDefault()},t}(r.Component);t.SearchResults=s},function(e,t,o){"use strict";var n=this&&this.__extends||function(){var e=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var o in t)t.hasOwnProperty(o)&&(e[o]=t[o])};return function(t,o){function n(){this.constructor=t}e(t,o),t.prototype=null===o?Object.create(o):(n.prototype=o.prototype,new n)}}();Object.defineProperty(t,"__esModule",{value:!0});var r=o(1),s=o(0),i=function(e){function t(){return null!==e&&e.apply(this,arguments)||this}return n(t,e),t.prototype.render=function(){switch(this.props.progress){case s.Data.Progress.GetCallbackToken:case s.Data.Progress.GetConversation:case s.Data.Progress.GetExcludedFolders:case s.Data.Progress.GetFolderNames:return r.createElement("h3",null,"Looking for other messages in this conversation...");case s.Data.Progress.Success:return null;case s.Data.Progress.NotFound:return r.createElement("h3",null,"It looks like you haven't filed this conversation anywhere before.");default:return r.createElement("div",null,r.createElement("h3",null,"Sorry, I couldn't figure out where this message should go. :("),r.createElement("span",null,this.props.message))}},t}(r.Component);t.StatusMessage=i},function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n=o(1),r=o(4),s=o(0),i=o(2),a=o(3);Office.initialize=function(){var e=/functions\.html(\?.*)?$/i,t=e.test(window.location.pathname),o=(Office.context||{}).mailbox,l="conversationFilerMatches";if(t)return void(window.fileDialog=function(t){var n=i.Factory.getData(o),r="conversationFilerNotification";console.log("Starting to load the conversation..."),n.getItemsAsync(function(s){if(console.log("Loaded the conversation: "+s.length),0===s.length)return o.item.notificationMessages.replaceAsync(r,{type:Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,message:"It looks like you haven't filed this conversation anywhere before."}),void t.completed();console.log("Showing the dialog..."),window.localStorage.setItem(l,JSON.stringify(s)),Office.context.ui.displayDialogAsync(window.location.href.replace(e,"dialog.html"),{height:25,width:50,displayInIframe:!0},function(e){var s=e.value,i=function(e){o.item.notificationMessages.removeAsync(r),window.localStorage.removeItem(l),e||s.close(),t.completed()};s.addEventHandler(Office.EventType.DialogMessageReceived,function(e){console.log("Moving the items..."),n.moveItemsAsync(e.message,function(e){console.log("Finished moving the items: "+e),i(!1)},function(e){console.log("Error moving the items: "+e),i(!1)})}),s.addEventHandler(Office.EventType.DialogEventReceived,function(){i(!0)})})},function(e){console.log("Progress loading the conversation: "+s.Data.Progress[e]),o.item.notificationMessages.replaceAsync(r,{type:Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,message:"Finding the messages in this conversation..."})},function(e){console.log("Error loading the conversation: "+e),o.item.notificationMessages.replaceAsync(r,{type:Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,message:"Sorry, I couldn't figure out where this message should go."}),t.completed()})});var d,c;/dialog\.html(\?.*)?$/i.test(window.location.pathname)&&(d=function(e){Office.context.ui.messageParent(e)},c=JSON.parse(window.localStorage.getItem(l))),r.render(n.createElement(a.ConversationFiler,{mailbox:o,onComplete:d,storedResults:c}),document.getElementById("conversationFilerRoot"))}}]);
//# sourceMappingURL=bundle.js.map