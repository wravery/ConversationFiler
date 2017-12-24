!function(e){function t(n){if(o[n])return o[n].exports;var s=o[n]={i:n,l:!1,exports:{}};return e[n].call(s.exports,s,s.exports,t),s.l=!0,s.exports}var o={};t.m=e,t.c=o,t.d=function(e,o,n){t.o(e,o)||Object.defineProperty(e,o,{configurable:!1,enumerable:!0,get:n})},t.n=function(e){var o=e&&e.__esModule?function(){return e.default}:function(){return e};return t.d(o,"a",o),o},t.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},t.p="",t(t.s=218)}({10:function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});!function(e){function t(e,t,o){if(t){var n=e.filter(function(e){return e.message.Id===t}).pop();if(n){console.log("Removing duplicates: "+e.length+" original message(s)");var s=e.filter(function(e){return e.message.ParentFolderId===n.message.ParentFolderId});console.log("Messages in the same folder: "+s.length+" message(s)"),e=e.filter(function(e){return e.folder.Id!==n.message.ParentFolderId||(console.log("Removed message in the same folder: "+e.message.Id),!1)}),e=e.filter(function(e){return!o.reduce(function(t,o){return t||o===e.folder.Id},!1)||(console.log("Removed message in an excluded folder: "+e.message.Id),!1)}),e=e.filter(function(e){return!s.reduce(function(t,o){return t||e.message.Sender===o.message.Sender&&e.message.ToRecipients===o.message.ToRecipients&&e.message.BodyPreview===o.message.BodyPreview},!1)||(console.log("Removed duplicate message in other folder: "+n.message.Id),!1)}),console.log("Remaining in other folders: "+e.length+" distinct message(s)")}}return e}!function(e){e[e.GetCallbackToken=0]="GetCallbackToken",e[e.GetConversation=1]="GetConversation",e[e.GetExcludedFolders=2]="GetExcludedFolders",e[e.GetFolderNames=3]="GetFolderNames",e[e.Success=4]="Success",e[e.NotFound=5]="NotFound",e[e.Error=6]="Error"}(e.Progress||(e.Progress={})),e.removeDuplicates=t}(t.Data||(t.Data={}))},168:function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});!function(e){function t(e){window.localStorage.setItem(s,JSON.stringify(e))}function o(){return JSON.parse(window.localStorage.getItem(s))}function n(){window.localStorage.removeItem(s)}var s="conversationFilerMatches";e.saveDialog=t,e.loadDialog=o,e.resetDialog=n}(t.DialogMessages||(t.DialogMessages={}))},218:function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n,s=o(168),i=o(10),r=o(63);!function(e){function t(){return window.location.href.replace(d,"/dialog.html")}function o(){return window.location.href.replace(d,"/about.html")}function n(e){var o=Office.context.mailbox,n=r.Factory.getData(o),a="conversationFilerNotification";console.log("Starting to load the conversation..."),n.getItemsAsync(function(i){if(console.log("Loaded the conversation: "+i.length),0===i.length)return o.item.notificationMessages.replaceAsync(a,{type:Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,message:"It looks like you haven't filed this conversation anywhere before.",icon:"file-icon-16",persistent:!1}),void e.completed();console.log("Showing the dialog..."),o.item.notificationMessages.removeAsync(a),s.DialogMessages.saveDialog(i),Office.context.ui.displayDialogAsync(t(),{height:40,width:50,displayInIframe:!0},function(t){var i=t.value,r=function(){s.DialogMessages.resetDialog(),e.completed()};i.addEventHandler(Office.EventType.DialogMessageReceived,function(e){var t=JSON.parse(e.message);if(t.canceled)return console.log("Dialog canceled"),i.close(),void r();console.log("Moving the items..."),i.close(),o.item.notificationMessages.replaceAsync(a,{type:Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,message:"Moving the items in this conversation..."}),n.moveItemsAsync(t.folderId,function(e){console.log("Finished moving the items: "+e),o.item.notificationMessages.replaceAsync(a,{type:Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,message:"I moved the items in this conversation, but there might be a short delay before that shows up in Outlook.",icon:"file-icon-16",persistent:!1}),r()},function(e){console.log("Error moving the items: "+e),o.item.notificationMessages.replaceAsync(a,{type:Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,message:"Something went wrong, I couldn't move the messages."}),r()})}),i.addEventHandler(Office.EventType.DialogEventReceived,function(){r()})})},function(e){console.log("Progress loading the conversation: "+i.Data.Progress[e]),o.item.notificationMessages.replaceAsync(a,{type:Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,message:"Finding the messages in this conversation..."})},function(t){console.log("Error loading the conversation: "+t),o.item.notificationMessages.replaceAsync(a,{type:Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,message:"Sorry, I couldn't figure out where this message should go."}),e.completed()})}function a(e){Office.context.ui.displayDialogAsync(o(),{height:40,width:25,displayInIframe:!0},function(t){var o=t.value;o.addEventHandler(Office.EventType.DialogMessageReceived,function(){console.log("Dialog closed with button"),o.close(),e.completed()}),o.addEventHandler(Office.EventType.DialogEventReceived,function(){console.log("Dialog closed"),e.completed()})})}function l(e){Office.context.mailbox.displayNewMessageForm({toRecipients:["wravery@hotmail.com"],subject:"Conversation Filer v2.0 App for Outlook"}),e.completed()}var d=/\/functions\.html(\?.*)?$/i;e.fileDialog=n,e.aboutDialog=a,e.sendFeedback=l}(n||(n={})),Office.initialize=function(){window.fileDialog=n.fileDialog,window.aboutDialog=n.aboutDialog,window.sendFeedback=n.sendFeedback}},63:function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n=o(64),s=o(65);!function(e){function t(e){return e.restUrl?new n.RESTData.Model(e):new s.EWSData.Model(e)}e.getData=t}(t.Factory||(t.Factory={}))},64:function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n=o(10);!function(e){var t;!function(e){e[e.Inbox=0]="Inbox",e[e.Drafts=1]="Drafts",e[e.SentItems=2]="SentItems",e[e.DeletedItems=3]="DeletedItems",e[e.Count=4]="Count"}(t||(t={}));var o=function(){function e(e){this.mailbox=e,this.itemId=this.getRestId(this.mailbox.item.itemId)}return e.prototype.loadItems=function(e,t,o){var s=this;this.onLoadComplete=e,this.onProgress=t,this.onError=o,console.log("Requesting the REST callback token..."),this.onProgress(n.Data.Progress.GetCallbackToken),this.mailbox.getCallbackTokenAsync({isRest:!0},function(e){e.status===Office.AsyncResultStatus.Succeeded?s.getConversation(e):s.onError(e.error.message)})},e.prototype.collateRequests=function(e,t,o){var n=this;e.length>1?$.when.apply($,e).done(function(){for(var e=[],o=0;o<arguments.length;o++)e[o]=arguments[o];t(e.map(function(e){return e[0]}))}).fail(function(e,t,o){n.onError(o)}):e[0].done(function(e){t([e])}).fail(function(e,t,o){n.onError(o)})},e.prototype.isOutlookMobile=function(){return"OutlookIOS"===this.mailbox.diagnostics.hostName||"OutlookAndroid"===this.mailbox.diagnostics.hostName},e.prototype.getRestId=function(e){return this.isOutlookMobile()?e:this.mailbox.convertToRestId(e,Office.MailboxEnums.RestVersion.v2_0)},e.prototype.getConversation=function(e){var t=this;this.token=e.value;var o=this.mailbox.item.conversationId,s=this.getRestId(o),i=this.mailbox.restUrl+"/v2.0/me/messages?$filter=ConversationId eq '"+s+"'&$select=Id,Subject,BodyPreview,Sender,ToRecipients,ParentFolderId";console.log("Getting the list of items in the conversation: "+i),this.onProgress(n.Data.Progress.GetConversation),$.ajax({url:i,async:!0,dataType:"json",headers:{Authorization:"Bearer "+this.token}}).done(function(e){t.getExcludedFolders(e)}).fail(function(e,o,n){t.onError(n)})},e.prototype.getExcludedFolders=function(e){var o=this;if(!e||!e.value||0===e.value.length)return void this.onLoadComplete([]);this.conversationMessages=e.value,console.log("Messages in the conversation: "+this.conversationMessages.length);for(var s=this.conversationMessages.filter(function(e){return e.Id===o.itemId}).reduce(function(e,t){return t.ParentFolderId},void 0),i=[],r=0;r<t.Count;++r){var a=t[r],l=this.mailbox.restUrl+"/v2.0/me/mailfolders/"+a+"?$select=Id";console.log("Getting excluded folder ID: "+l),i.push($.ajax({url:l,async:!0,dataType:"json",headers:{Authorization:"Bearer "+this.token}}))}this.onProgress(n.Data.Progress.GetExcludedFolders),this.collateRequests(i,function(e){var t=e.map(function(e){return e.Id});o.getFolderNames(s,t)},function(e){o.onError(e)})},e.prototype.getFolderNames=function(e,t){var o=this,s=this.conversationMessages.reduce(function(e,t){var o=e.filter(function(e){return e.folder.Id===t.ParentFolderId}).pop();return o?o.messages.push(t):e.push({folder:{Id:t.ParentFolderId},messages:[t]}),e},[]);this.currentFolderId=e,this.excludedFolderIds=t;var i=s.filter(function(e){return!o.excludedFolderIds.reduce(function(t,o){return t||o===e.folder.Id},!1)}).map(function(e){var t=o.mailbox.restUrl+"/v2.0/me/mailfolders/"+e.folder.Id+"?$select=Id,DisplayName";return console.log("Getting included folder name: "+t),$.ajax({url:t,async:!0,dataType:"json",headers:{Authorization:"Bearer "+o.token}})});if(0===i.length)return void this.onLoadComplete([]);this.onProgress(n.Data.Progress.GetFolderNames),this.collateRequests(i,function(e){e.forEach(function(e){for(var t=0;t<s.length;++t)if(s[t].folder.Id===e.Id){s[t].folder.DisplayName=e.DisplayName;break}});var t=s.reduce(function(e,t){return e+t.messages.length},0),i=s.length;console.log("Found "+t+" message(s) in "+i+" folder(s)");var r=s.reduce(function(e,t){return e.concat(t.messages.map(function(e){return{message:{Id:e.Id,BodyPreview:e.BodyPreview,Sender:e.Sender.EmailAddress.Name,ToRecipients:e.ToRecipients.map(function(e){return e.EmailAddress.Name}).join("; "),ParentFolderId:e.ParentFolderId},folder:{Id:t.folder.Id,DisplayName:t.folder.DisplayName}}}))},[]);console.log("Finished loading items: "+r.length),o.onLoadComplete(n.Data.removeDuplicates(r,o.itemId,o.excludedFolderIds))},function(e){o.onError(e)})},e.prototype.moveItems=function(e,t,o){var n=this;this.onMoveComplete=t,this.onError=o,console.log("Moving items to folder: "+e);var s=this.conversationMessages.filter(function(e){return e.ParentFolderId!==n.currentFolderId}).map(function(t){var o=n.mailbox.restUrl+"/v2.0/me/messages/"+t.Id+"/move";return console.log("Moving item: "+o),$.ajax({url:o,async:!0,method:"POST",contentType:"application/json",dataType:"json",data:JSON.stringify({DestinationId:e}),headers:{Authorization:"Bearer "+n.token}})});this.collateRequests(s,function(e){console.log("Finished moving items to other folder: "+e.length),n.onMoveComplete(e.length)},function(e){n.onError(e)})},e}(),s=function(){function e(e){this.context=new o(e)}return e.prototype.getItemsAsync=function(e,t,o){this.context.loadItems(e,t,o)},e.prototype.moveItemsAsync=function(e,t,o){this.context.moveItems(e,t,o)},e}();e.Model=s}(t.RESTData||(t.RESTData={}))},65:function(e,t,o){"use strict";Object.defineProperty(t,"__esModule",{value:!0});var n=o(10);!function(e){var t=function(){function e(){}return e.getItemsRequest=function(t){var o=[e.beginRequest,"  <m:GetItem>","    <m:ItemShape>","      <t:BaseShape>IdOnly</t:BaseShape>","      <t:BodyType>Text</t:BodyType>","      <t:AdditionalProperties>",'        <t:FieldURI FieldURI="item:ParentFolderId" />','        <t:FieldURI FieldURI="message:Sender" />','        <t:FieldURI FieldURI="message:ToRecipients" />','        <t:FieldURI FieldURI="item:Body" />',"      </t:AdditionalProperties>","    </m:ItemShape>","    <m:ItemIds>"];return t.map(function(e){o.push('      <t:ItemId Id="'+e.itemId.id+'" ChangeKey="'+e.itemId.changeKey+'" />')}),o.push("    </m:ItemIds>","  </m:GetItem>",e.endRequest),o.join("\n")},e.getFolderNamesRequest=function(t){var o=[e.beginRequest,"  <m:GetFolder>","    <m:FolderShape>","      <t:BaseShape>IdOnly</t:BaseShape>","      <t:AdditionalProperties>",'        <t:FieldURI FieldURI="folder:DisplayName" />',"      </t:AdditionalProperties>","    </m:FolderShape>","    <m:FolderIds>"];return t.map(function(e){o.push('      <t:FolderId Id="'+e.folderId.id+'" ChangeKey="'+e.folderId.changeKey+'" />')}),o.push("    </m:FolderIds>","  </m:GetFolder >",e.endRequest),o.join("\n")},e.moveItemsRequest=function(t,o){var n=[e.beginRequest,"  <m:MoveItem>","    <m:ToFolderId>",'      <t:FolderId Id="'+o+'"/>',"    </m:ToFolderId>","    <m:ItemIds>"];return t.map(function(e){n.push('      <t:ItemId Id="'+e.itemId.id+'" ChangeKey="'+e.itemId.changeKey+'" />')}),n.push("    </m:ItemIds>","  </m:MoveItem>",e.endRequest),n.join("\n")},e.beginRequest=['<?xml version="1.0" encoding="utf-8" ?>','<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"','    xmlns:xsd="http://www.w3.org/2001/XMLSchema"','    xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"','    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"','    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">',"<soap:Header>",'  <t:RequestServerVersion Version="Exchange2010_SP1" />',"</soap:Header>","<soap:Body>"].join("\n"),e.endRequest=["</soap:Body>","</soap:Envelope>"].join("\n"),e.findConversationRequest=[e.beginRequest,"  <m:FindConversation>",'    <m:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="20" />',"    <m:ParentFolderId>",'      <t:DistinguishedFolderId Id="inbox"/>',"    </m:ParentFolderId>","  </m:FindConversation>",e.endRequest].join("\n"),e.excludedFolderIdsRequest=[e.beginRequest,"  <m:GetFolder>","    <m:FolderShape>","      <t:BaseShape>IdOnly</t:BaseShape>","    </m:FolderShape>","    <m:FolderIds>",'      <t:DistinguishedFolderId Id="inbox"/>','      <t:DistinguishedFolderId Id="drafts"/>','      <t:DistinguishedFolderId Id="sentitems"/>','      <t:DistinguishedFolderId Id="deleteditems"/>',"    </m:FolderIds>","  </m:GetFolder >",e.endRequest].join("\n"),e}(),o=function(){function e(e){this.mailbox=e,this.itemId=this.mailbox.item.itemId}return e.prototype.loadItems=function(e,o,s){var i=this;this.onLoadComplete=e,this.onProgress=o,this.onError=s,console.log("Finding the conversation with EWS..."),this.onProgress(n.Data.Progress.GetConversation),this.mailbox.makeEwsRequestAsync(t.findConversationRequest,function(e){if(!e.value)return void i.onError(e.error.message);i.conversationXml=$.parseXML(e.value),i.getConversation()})},e.prototype.getConversation=function(){var e=this,t=$(this.conversationXml.querySelectorAll("Conversation > GlobalItemIds > ItemId")).filter('[Id="'+this.mailbox.item.itemId+'"]').parents("Conversation");if(!t.length)return void this.onError("This conversation isn't in your inbox's top 20.");if(parseInt(t.find("MessageCount").text())>=parseInt(t.find("GlobalMessageCount").text()))return void this.onLoadComplete([]);var o=[];t.find("ItemIds > ItemId").each(function(){var e=$(this);o.push({id:e.attr("Id"),changeKey:e.attr("ChangeKey")})});var n=[];if(t.find("GlobalItemIds > ItemId").each(function(){var e=$(this);n.push({id:e.attr("Id"),changeKey:e.attr("ChangeKey")})}),!o.length||n.length<=o.length)return void this.onLoadComplete([]);this.conversation={id:this.mailbox.item.conversationId,items:o.map(function(t){return{itemId:t,conversation:e.conversation}}),global:n.map(function(t){return{itemId:t,conversation:e.conversation}})},this.loadExcludedFolders()},e.prototype.loadExcludedFolders=function(){var e=this;this.excludedFolders?this.loadMessages():(console.log("Getting the list of excluded folders"),this.onProgress(n.Data.Progress.GetExcludedFolders),this.mailbox.makeEwsRequestAsync(t.excludedFolderIdsRequest,function(t){if(!t.value)return void e.onError(t.error.message);var o=$.parseXML(t.value),n=[];$(o.querySelectorAll("GetFolderResponseMessage > Folders > Folder > FolderId")).each(function(){var e=$(this);n.push({folderId:{id:e.attr("Id"),changeKey:e.attr("ChangeKey")}})}),e.excludedFolders=n,e.loadMessages()}))},e.prototype.loadMessages=function(){var e=this;this.itemsXml?this.getMessages():(console.log("Getting the messages in other folders: "+this.conversation.global.length),this.onProgress(n.Data.Progress.GetConversation),this.mailbox.makeEwsRequestAsync(t.getItemsRequest(this.conversation.global),function(t){if(!t.value)return void e.onError(t.error.message);e.itemsXml=$.parseXML(t.value),e.getMessages()}))},e.prototype.getMessages=function(){var e=$(this.itemsXml.querySelectorAll("GetItemResponseMessage > Items > Message > ParentFolderId")).parent();this.conversation.global.map(function(t){for(var o=0;o<e.length;o++){var n=e[o];if(n.querySelector('ItemId[Id="'+t.itemId.id+'"]')){var s=n.querySelector("ParentFolderId");t.folder={folderId:{id:s.getAttribute("Id"),changeKey:s.getAttribute("ChangeKey")}},t.from=n.querySelector("Sender > Mailbox > Name").textContent,t.to=n.querySelector("ToRecipients > Mailbox > Name").textContent,t.body=n.querySelector("Body").textContent.slice(0,200);break}}}),this.loadFolderDisplayNames()},e.prototype.loadFolderDisplayNames=function(){var e=this;if(this.folderNamesXml)this.getFolderDisplayNames();else{var o=[];if(this.conversation.global.map(function(e){e.folder&&o.push(e.folder)}),!o.length)return void this.onLoadComplete([]);console.log("Getting the display names of the other folders: "+o.length),this.onProgress(n.Data.Progress.GetFolderNames),this.mailbox.makeEwsRequestAsync(t.getFolderNamesRequest(o),function(t){if(!t.value)return void e.onError(t.error.message);e.folderNamesXml=$.parseXML(t.value),e.getFolderDisplayNames()})}},e.prototype.getFolderDisplayNames=function(){var e=this,t=[];this.conversation.global.map(function(o){if(o.folder){var n=e.folderNamesXml.querySelector('GetFolderResponseMessage > Folders > Folder > FolderId[Id="'+o.folder.folderId.id+'"]').parentNode;o.folder.displayName=n.querySelector("DisplayName").textContent,t.push({message:{Id:o.itemId.id,BodyPreview:o.body,Sender:o.from,ToRecipients:o.to,ParentFolderId:o.folder.folderId.id},folder:{Id:o.folder.folderId.id,DisplayName:o.folder.displayName}})}}),console.log("Finished loading items in other folders: "+t.length),this.onLoadComplete(n.Data.removeDuplicates(t,this.itemId,this.excludedFolders.map(function(e){return e.folderId.id})))},e.prototype.moveItems=function(e,o,n){var s=this;this.onMoveComplete=o,this.onError=n,console.log("Moving items to folder: "+e),this.mailbox.makeEwsRequestAsync(t.moveItemsRequest(this.conversation.items,e),function(e){if(!e.value)return void s.onError(e.error.message);console.log("Finished moving items to other folder: "+s.conversation.items.length),s.onMoveComplete(s.conversation.items.length)})},e}(),s=function(){function e(e){this.context=new o(e)}return e.prototype.getItemsAsync=function(e,t,o){this.context.loadItems(e,t,o)},e.prototype.moveItemsAsync=function(e,t,o){this.context.moveItems(e,t,o)},e}();e.Model=s}(t.EWSData||(t.EWSData={}))}});
//# sourceMappingURL=functions.js.map