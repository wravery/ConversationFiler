import { Data } from "./Model";

export module EWSData {
    interface ItemId {
        id: string;
        changeKey: string;
    }

    interface FolderData {
        folderId: ItemId;
        displayName?: string;
    }

    interface MessageData {
        itemId: ItemId;
        conversation: ConversationData;
        folder?: FolderData;
        body?: string;
        from?: string;
        to?: string;
    }

    interface ConversationData {
        id: string;
        items: MessageData[];
        global: MessageData[];
    }

    class RequestBuilder {
        private static beginRequest = [
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

        private static endRequest = [
            '</soap:Body>',
            '</soap:Envelope>'
        ].join('\n');

        static findConversationRequest = [
            RequestBuilder.beginRequest,
            '  <m:FindConversation>',
            '    <m:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="20" />',
            '    <m:ParentFolderId>',
            '      <t:DistinguishedFolderId Id="inbox"/>',
            '    </m:ParentFolderId>',
            '  </m:FindConversation>',
            RequestBuilder.endRequest
        ].join('\n');

        static getItemsRequest(messages: MessageData[]) {
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

            messages.map(message => {
                builder.push(`      <t:ItemId Id="${message.itemId.id}" ChangeKey="${message.itemId.changeKey}" />`);
            });

            builder.push(
                '    </m:ItemIds>',
                '  </m:GetItem>',
                RequestBuilder.endRequest);

            return builder.join('\n');
        }

        static excludedFolderIdsRequest = [
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

        static getFolderNamesRequest(folders: FolderData[]) {
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

            folders.map(folder => {
                builder.push(`      <t:FolderId Id="${folder.folderId.id}" ChangeKey="${folder.folderId.changeKey}" />`);
            });

            builder.push(
                '    </m:FolderIds>',
                '  </m:GetFolder >',
                RequestBuilder.endRequest);

            return builder.join('\n');
        }

        static moveItemsRequest(messages: MessageData[], folderId: string) {
            var builder = [
                RequestBuilder.beginRequest,
                '  <m:MoveItem>',
                '    <m:ToFolderId>',
                '      <t:FolderId Id="' + folderId + '"/>',
                '    </m:ToFolderId>',
                '    <m:ItemIds>',
            ];

            messages.map(message => {
                builder.push(`      <t:ItemId Id="${message.itemId.id}" ChangeKey="${message.itemId.changeKey}" />`);
            });

            builder.push(
                '    </m:ItemIds>',
                '  </m:MoveItem>',
                RequestBuilder.endRequest);

            return builder.join('\n');
        }
    }

    class Context {
        constructor(private mailbox: Office.Mailbox) {
            this.itemId = (<Office.ItemRead>this.mailbox.item).itemId;
        }

        private itemId: string;
        private conversationXml?: XMLDocument;
        private conversation?: ConversationData;
        private excludedFolders?: FolderData[];
        private itemsXml?: XMLDocument;
        private folderNamesXml?: XMLDocument;

        private onLoadComplete?: (results: Data.Match[]) => void;
        private onProgress?: (progress: Data.Progress) => void;
        private onError?: (message: string) => void;
        private onMoveComplete?: (count: number) => void

        loadItems(onLoadComplete: (results: Data.Match[]) => void, onProgress: (progress: Data.Progress) => void, onError: (message: string) => void): void {
            this.onLoadComplete = onLoadComplete;
            this.onProgress = onProgress;
            this.onError = onError;

            console.log('Finding the conversation with EWS...');
            this.onProgress(Data.Progress.GetConversation);

            this.mailbox.makeEwsRequestAsync(RequestBuilder.findConversationRequest, (result) => {
                if (!result.value) {
                    this.onError(result.error.message);
                    return;
                }

                this.conversationXml = $.parseXML(result.value);
                this.getConversation();
            });
        }

        private getConversation() {
            const $conversation = $(this.conversationXml.querySelectorAll('Conversation > GlobalItemIds > ItemId'))
                .filter(`[Id="${(<Office.ItemRead>this.mailbox.item).itemId}"]`)
                .parents('Conversation');

            if (!$conversation.length) {
                this.onError("This conversation isn't in your inbox's top 20.");
                return;
            }

            const messageCount = parseInt($conversation.find('MessageCount').text());
            const globalCount = parseInt($conversation.find('GlobalMessageCount').text());

            if (messageCount >= globalCount) {
                this.onLoadComplete([]);
                return;
            }

            let sameFolderItemIds: ItemId[] = [];

            $conversation.find('ItemIds > ItemId').each(function () {
                const $this = $(this);

                sameFolderItemIds.push({
                    id: $this.attr('Id'),
                    changeKey: $this.attr('ChangeKey')
                });
            });

            let otherFolderItemIds: ItemId[] = [];

            $conversation.find('GlobalItemIds > ItemId').each(function () {
                const $this = $(this);

                otherFolderItemIds.push({
                    id: $this.attr('Id'),
                    changeKey: $this.attr('ChangeKey')
                });
            });

            if (!sameFolderItemIds.length || otherFolderItemIds.length <= sameFolderItemIds.length) {
                this.onLoadComplete([]);
                return;
            }

            this.conversation = {
                id: (<Office.Message>this.mailbox.item).conversationId,
                items: sameFolderItemIds.map(itemId => ({ itemId: itemId, conversation: this.conversation })),
                global: otherFolderItemIds.map(itemId => ({ itemId: itemId, conversation: this.conversation }))
            };

            this.loadExcludedFolders();
        }

        private loadExcludedFolders() {
            if (this.excludedFolders) {
                this.loadMessages();
            } else {
                console.log('Getting the list of excluded folders');
                this.onProgress(Data.Progress.GetExcludedFolders);

                this.mailbox.makeEwsRequestAsync(RequestBuilder.excludedFolderIdsRequest, (result) => {
                    if (!result.value) {
                        this.onError(result.error.message);
                        return;
                    }

                    let foldersXml = $.parseXML(result.value);
                    let excludedFolders: FolderData[] = [];

                    $(foldersXml.querySelectorAll('GetFolderResponseMessage > Folders > Folder > FolderId')).each(function () {
                        var $this = $(this);
                        excludedFolders.push({
                            folderId: {
                                id: $this.attr('Id'),
                                changeKey: $this.attr('ChangeKey')
                            }
                        });
                    });

                    this.excludedFolders = excludedFolders;

                    this.loadMessages();
                });
            }
        }

        private loadMessages() {
            if (this.itemsXml) {
                this.getMessages();
            } else {
                console.log(`Getting the messages in other folders: ${this.conversation.global.length}`);
                this.onProgress(Data.Progress.GetConversation);

                this.mailbox.makeEwsRequestAsync(RequestBuilder.getItemsRequest(this.conversation.global), (result) => {
                    if (!result.value) {
                        this.onError(result.error.message);
                        return;
                    }

                    this.itemsXml = $.parseXML(result.value);
                    this.getMessages();
                });
            }
        }

        private getMessages() {
            let $messages = $(this.itemsXml.querySelectorAll('GetItemResponseMessage > Items > Message > ParentFolderId'));

            this.excludedFolders.map(folder => {
                $messages = $messages.filter(`[Id!="${folder.folderId.id}"]`);
            });

            $messages = $messages.parent();

            this.conversation.global.map(item => {
                for (let i = 0; i < $messages.length; i++) {
                    const msg = $messages[i];

                    if (msg.querySelector(`ItemId[Id="${item.itemId.id}"]`)) {
                        let folderId = msg.querySelector('ParentFolderId');

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
            });

            this.loadFolderDisplayNames();
        }

        private loadFolderDisplayNames() {
            if (this.folderNamesXml) {
                this.getFolderDisplayNames();
            } else {
                let destinations: FolderData[] = [];

                this.conversation.global.map(item => {
                    if (item.folder) {
                        destinations.push(item.folder);
                    }
                });

                if (!destinations.length) {
                    this.onLoadComplete([]);
                    return;
                }

                console.log(`Getting the display names of the other folders: ${destinations.length}`);
                this.onProgress(Data.Progress.GetFolderNames);

                this.mailbox.makeEwsRequestAsync(RequestBuilder.getFolderNamesRequest(destinations), (result) => {
                    if (!result.value) {
                        this.onError(result.error.message);
                        return;
                    }

                    this.folderNamesXml = $.parseXML(result.value);
                    this.getFolderDisplayNames();
                });
            }
        }

        private getFolderDisplayNames() {
            let matches: Data.Match[] = [];

            this.conversation.global.map((item: MessageData) => {
                if (!item.folder) {
                    return;
                }

                const folder = this.folderNamesXml.querySelector(`GetFolderResponseMessage > Folders > Folder > FolderId[Id="${item.folder.folderId.id}"]`).parentNode;
                item.folder.displayName = (<Element>folder).querySelector('DisplayName').textContent;

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

            console.log(`Finished loading items in other folders: ${matches.length}`);
            this.onLoadComplete(Data.removeDuplicates(matches, this.itemId));
        }

        moveItems(folderId: string, onMoveComplete: (count: number) => void, onError: (message: string) => void) {
            this.onMoveComplete = onMoveComplete;
            this.onError = onError;

            console.log(`Moving items to folder: ${folderId}`);

            this.mailbox.makeEwsRequestAsync(RequestBuilder.moveItemsRequest(this.conversation.items, folderId), (result) => {
                if (!result.value) {
                    this.onError(result.error.message);
                    return;
                }

                console.log(`Finished moving items to other folder: ${this.conversation.items.length}`);
                this.onMoveComplete(this.conversation.items.length);
            });
        }
    }

    export class Model implements Data.IModel {
        private context?: Context;

        constructor(mailbox: Office.Mailbox) {
            this.context = new Context(mailbox);
        }

        getItemsAsync(onLoadComplete: (results: Data.Match[]) => void, onProgress: (progress: Data.Progress) => void, onError: (message: string) => void): void {
            this.context.loadItems(onLoadComplete, onProgress, onError);
        }

        moveItemsAsync(folderId: string, onMoveComplete: (count: number) => void, onError: (message: string) => void): void {
            this.context.moveItems(folderId, onMoveComplete, onError);
        }
    }
}