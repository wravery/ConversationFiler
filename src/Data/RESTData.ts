/// <reference path="../_references.ts" />

import { Data } from "./Model"

export module RESTData {
    const Endpoint = '/v2.0/me';

    enum ExcludedFolders {
        Inbox,
        Drafts,
        SentItems,
        DeletedItems,

        // Sentinel value for enumerating the folder names
        Count
    }

    interface EmailAddressJson {
        Name: string;
        Address: string;
    }

    interface AddressJson {
        EmailAddress: EmailAddressJson;
    }

    interface MessageJson {
        Id: string;
        BodyPreview: string;
        Sender: AddressJson;
        ToRecipients: AddressJson[];
        ParentFolderId: string;
    }

    interface MessageJsonCollection {
        value: MessageJson[];
    }

    interface FolderJson {
        Id: string;
        DisplayName?: string;
    }

    class Context {
        constructor(private mailbox: Office.Mailbox) {
        }

        private token?: string;
        private currentFolderId?: string;
        private conversationMessages?: MessageJson[];
        private excludedFolderIds?: string[];

        private onLoadComplete?: (results: Data.Match[]) => void;
        private onProgress?: (progress: Data.Progress) => void;
        private onError?: (message: string) => void;
        private onMoveComplete?: (count: number) => void

        loadItems(onLoadComplete: (results: Data.Match[]) => void, onProgress: (progress: Data.Progress) => void, onError: (message: string) => void): void {
            this.onLoadComplete = onLoadComplete;
            this.onProgress = onProgress;
            this.onError = onError;

            console.log('Requesting the REST callback token...');
            this.onProgress(Data.Progress.GetCallbackToken);

            // Start the chain of requests by getting a callback token.
            this.mailbox.getCallbackTokenAsync({ isRest: true },
                (result: Office.AsyncResult) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        this.getConversation(result);
                    } else {
                        this.onError(result.error.message);
                    }
                });
        }

        // Sometimes we need to make separate REST requests for multiple items. Wait until they all complete and then
        // invoke the callbacks all at once with an array of typed results.
        private collateRequests<T>(requests: JQueryPromise<T>[], onDone: (results: T[]) => void, onFail: (message: string) => void): void {
            if (requests.length > 1) {
                $.when.apply($, requests)
                    .done((...results: any[]) => {
                        let values: T[] = [];

                        results.map((result: any[]) => {
                            values.push(<T>result[0]);
                        });

                        onDone(values);
                    }).fail((message: string) => {
                        this.onError(message);
                    });
            } else {
                requests[0]
                    .done((result: T) => {
                        onDone([result]);
                    }).fail((message: string) => {
                        this.onError(message);
                    });
            }
        }

        // Send a REST request to retrieve a list of messages in this conversation.
        private getConversation(result: Office.AsyncResult) {
            this.token = <string>result.value;

            const conversationId = (<Office.Message>this.mailbox.item).conversationId;
            const restConversationId = this.mailbox.diagnostics.hostName === 'OutlookIOS'
                ? conversationId
                : this.mailbox.convertToRestId(conversationId, Office.MailboxEnums.RestVersion.v2_0);
            const restUrl = `${this.mailbox.restUrl}${Endpoint}/messages?$filter=ConversationId eq '${restConversationId}'&$select=Id,Subject,BodyPreview,Sender,ToRecipients,ParentFolderId`;

            console.log(`Getting the list of items in the conversation: ${restUrl}`);
            this.onProgress(Data.Progress.GetConversation);

            $.ajax({
                url: restUrl,
                async: true,
                dataType: 'json',
                headers: { 'Authorization': `Bearer ${this.token}` }
            }).done((result: MessageJsonCollection) => {
                this.getExcludedFolders(result);
            }).fail((message: string) => {
                this.onError(message);
            });
        }

        // Send a REST request to identify each of the folders we want to exclude in our results.
        private getExcludedFolders(result: MessageJsonCollection) {
            if (!result || !result.value || 0 === result.value.length) {
                this.onLoadComplete([]);
                return;
            }

            this.conversationMessages = result.value;

            let currentFolderId: string;
            let excludedFolderIds: string[] = [];

            // We should ignore any messages in the same folder.
            const itemId = (<Office.ItemRead>this.mailbox.item).itemId;
            const restItemId = this.mailbox.diagnostics.hostName === 'OutlookIOS'
                ? itemId
                : this.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);

            for (let i = 0; i < this.conversationMessages.length; ++i) {
                if (this.conversationMessages[i].Id === restItemId) {
                    currentFolderId = this.conversationMessages[i].ParentFolderId;
                    excludedFolderIds.push(currentFolderId);
                    break;
                }
            }

            // We should also exclude some special folders, but we need to get their folderIds.
            let requests: JQueryXHR[] = [];

            for (let i = 0; i < ExcludedFolders.Count; ++i) {
                const folderId = ExcludedFolders[i];
                const restUrl = `${this.mailbox.restUrl}${Endpoint}/mailfolders/${folderId}?$select=Id`;

                console.log(`Getting excluded folder ID: ${restUrl}`);

                requests.push($.ajax({
                    url: restUrl,
                    async: true,
                    dataType: 'json',
                    headers: { 'Authorization': `Bearer ${this.token}` }
                }));
            }

            this.onProgress(Data.Progress.GetExcludedFolders);

            this.collateRequests(<JQueryPromise<FolderJson>[]>requests, (results: FolderJson[]) => {
                results.map((value: FolderJson) => {
                    excludedFolderIds.push(value.Id);
                });

                this.getFolderNames(currentFolderId, excludedFolderIds);
            }, (message: string) => {
                this.onError(message);
            });
        }

        // Send REST requests to fill in the display names of all the folders we are not excluding.
        private getFolderNames(currentFolderId: string, excludedFolderIds: string[]) {
            let folderMap: {
                folder: FolderJson;
                messages: MessageJson[];
            }[] = [];

            this.conversationMessages.map((message: MessageJson) => {
                for (let i = 0; i < excludedFolderIds.length; ++i) {
                    if (excludedFolderIds[i] === message.ParentFolderId) {
                        // Skip this message.
                        return;
                    }
                }

                for (let i = 0; i < folderMap.length; ++i) {
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

            let requests: JQueryXHR[] = [];

            folderMap.map((entry) => {
                const restUrl = `${this.mailbox.restUrl}${Endpoint}/mailfolders/${entry.folder.Id}?$select=Id,DisplayName`;

                console.log(`Getting included folder name: ${restUrl}`);

                requests.push($.ajax({
                    url: restUrl,
                    async: true,
                    dataType: 'json',
                    headers: { 'Authorization': `Bearer ${this.token}` }
                }));
            });

            this.onProgress(Data.Progress.GetFolderNames);

            this.collateRequests(<JQueryPromise<FolderJson>[]>requests, (results: FolderJson[]) => {
                results.map((value) => {
                    for (let i = 0; i < folderMap.length; ++i) {
                        if (folderMap[i].folder.Id === value.Id) {
                            folderMap[i].folder.DisplayName = value.DisplayName;
                            break;
                        }
                    }
                });

                let matches: Data.Match[] = [];

                folderMap.map((entry) => {
                    entry.messages.map((message) => {
                        let recipients: string[] = [];

                        message.ToRecipients.map((address) => {
                            recipients.push(address.EmailAddress.Name);
                        });

                        let value: Data.Message = {
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

                console.log(`Finished loading items in other folders: ${matches.length}`);
                this.onLoadComplete(matches);
            }, (message: string) => {
                this.onError(message);
            });
        }

        moveItems(folderId: string, onMoveComplete: (count: number) => void, onError: (message: string) => void) {
            this.onMoveComplete = onMoveComplete;
            this.onError = onError;

            console.log(`Moving items to folder: ${folderId}`);

            let requests: JQueryXHR[] = [];

            this.conversationMessages.map((message: MessageJson) => {
                if (message.ParentFolderId !== this.currentFolderId) {
                    // Skip any messages that are not in the current folder.
                    return;
                }

                const restUrl = `${this.mailbox.restUrl}${Endpoint}/messages/${message.Id}/move`;

                console.log(`Moving item: ${restUrl}`);

                requests.push($.ajax({
                    url: restUrl,
                    async: true,
                    method: 'POST',
                    contentType: 'application/json',
                    dataType: 'json',
                    data: JSON.stringify({ DestinationId: folderId }),
                    headers: { 'Authorization': `Bearer ${this.token}` }
                }));
            });

            this.collateRequests(<JQueryPromise<MessageJson>[]>requests, (results: MessageJson[]) => {
                console.log(`Finished moving items to other folder: ${results.length}`);
                this.onMoveComplete(results.length);
            }, (message: string) => {
                this.onError(message);
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