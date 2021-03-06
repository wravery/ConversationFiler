import { Data } from "./Model";

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
        constructor(private mailbox: Data.Mailbox) {
            this.itemId = this.getRestId((<Office.ItemRead>this.mailbox.item).itemId);
        }

        private itemId: string;
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
        private collateRequests<T>(requests: JQuery.jqXHR<T>[], onDone: (results: T[]) => void, onFail: (message: string) => void): void {
            if (requests.length > 1) {
                $.when.apply($, requests)
                    .done((...results: any[]) => {
                        onDone(results.map(result => <T>result[0]));
                    }).fail((request: JQuery.jqXHR, status: JQuery.Ajax.ErrorTextStatus, message: string) => {
                        this.onError(message);
                    });
            } else {
                requests[0]
                    .done((result: T) => {
                        onDone([result]);
                    }).fail((request: JQuery.jqXHR, status: JQuery.Ajax.ErrorTextStatus, message: string) => {
                        this.onError(message);
                    });
            }
        }

        // Test for each of the Diagnostic.hostName values that Outlook Mobile uses.
        private isOutlookMobile() {
            return this.mailbox.diagnostics.hostName === 'OutlookIOS'
                || this.mailbox.diagnostics.hostName === 'OutlookAndroid';
        }

        // Perform any necessary conversion from the EWS ID format to the REST ID format.
        private getRestId(itemId: string) {
            if (this.isOutlookMobile()) {
                // If we're running in Outlook Mobile on iOS or Android, the IDs we get from
                // Office.context.mailbox.item are already REST IDs.
                return itemId;
            }

            // Otherwise we need to convert them from the EWS ID format to the REST ID format.
            return this.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
        }

        // Send a REST request to retrieve a list of messages in this conversation.
        private getConversation(result: Office.AsyncResult) {
            this.token = <string>result.value;

            const conversationId = (<Office.Message>this.mailbox.item).conversationId;
            const restConversationId = this.getRestId(conversationId);
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
            }).fail((request: JQuery.jqXHR, status: JQuery.Ajax.ErrorTextStatus, message: string) => {
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

            console.log(`Messages in the conversation: ${this.conversationMessages.length}`);

            // Get the current folderId.
            const currentFolderId = this.conversationMessages
                .filter(value => value.Id === this.itemId)
                .reduce((previousValue: string, value) => value.ParentFolderId, undefined);

            // We should exclude some special folders, but we need to get their folderIds.
            let requests: JQuery.jqXHR<FolderJson>[] = [];

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

            this.collateRequests(requests, (results) => {
                const excludedFolderIds = results.map(value => value.Id);

                this.getFolderNames(currentFolderId, excludedFolderIds);
            }, (message) => {
                this.onError(message);
            });
        }

        // Send REST requests to fill in the display names of all the folders we are not excluding.
        private getFolderNames(currentFolderId: string, excludedFolderIds: string[]) {
            interface folderMapEntry {
                folder: FolderJson;
                messages: MessageJson[];
            };

            const folderMap = this.conversationMessages
                .reduce((previousValue: folderMapEntry[], message) => {
                    const entry = previousValue
                        .filter(value => value.folder.Id === message.ParentFolderId)
                        .pop();

                    if (entry) {
                        // Add this message to the existing entry.
                        entry.messages.push(message);
                    } else {
                        // Create a new entry for this folder.
                        previousValue.push({ folder: { Id: message.ParentFolderId }, messages: [message] });
                    }

                    return previousValue;
                }, []);

            this.currentFolderId = currentFolderId;
            this.excludedFolderIds = excludedFolderIds;

            const requests: JQuery.jqXHR<FolderJson>[] = folderMap
                .filter(entry => !this.excludedFolderIds.reduce((previousValue, value) =>
                    previousValue || value === entry.folder.Id, false))
                .map((entry) => {
                    const restUrl = `${this.mailbox.restUrl}${Endpoint}/mailfolders/${entry.folder.Id}?$select=Id,DisplayName`;

                    console.log(`Getting included folder name: ${restUrl}`);

                    return $.ajax({
                        url: restUrl,
                        async: true,
                        dataType: 'json',
                        headers: { 'Authorization': `Bearer ${this.token}` }
                    });
                });

            if (requests.length === 0) {
                this.onLoadComplete([]);
                return;
            }

            this.onProgress(Data.Progress.GetFolderNames);

            this.collateRequests(requests, (results: FolderJson[]) => {
                results.forEach((value) => {
                    for (let i = 0; i < folderMap.length; ++i) {
                        if (folderMap[i].folder.Id === value.Id) {
                            folderMap[i].folder.DisplayName = value.DisplayName;
                            break;
                        }
                    }
                });

                const messageCount = folderMap.reduce((previousValue, currentValue) => previousValue + currentValue.messages.length, 0);
                const folderCount = folderMap.length;

                console.log(`Found ${messageCount} message(s) in ${folderCount} folder(s)`);

                const matches = folderMap.reduce((previousValue: Data.Match[], currentValue) =>
                    previousValue.concat(currentValue.messages.map(item => (<Data.Match>{
                        message: {
                            Id: item.Id,
                            BodyPreview: item.BodyPreview,
                            Sender: item.Sender.EmailAddress.Name,
                            ToRecipients: item.ToRecipients.map(address => address.EmailAddress.Name).join('; '),
                            ParentFolderId: item.ParentFolderId
                        },
                        folder: {
                            Id: currentValue.folder.Id,
                            DisplayName: currentValue.folder.DisplayName
                        }
                    }))), []);

                console.log(`Finished loading items: ${matches.length}`);
                this.onLoadComplete(Data.removeDuplicates(matches, this.itemId, this.excludedFolderIds));
            }, (message: string) => {
                this.onError(message);
            });
        }

        moveItems(folderId: string, onMoveComplete: (count: number) => void, onError: (message: string) => void) {
            this.onMoveComplete = onMoveComplete;
            this.onError = onError;

            console.log(`Moving items to folder: ${folderId}`);

            const requests: JQuery.jqXHR<MessageJson>[] = this.conversationMessages
                .filter(message => message.ParentFolderId === this.currentFolderId)
                .map(message => {
                    const restUrl = `${this.mailbox.restUrl}${Endpoint}/messages/${message.Id}/move`;

                    console.log(`Moving item: ${restUrl}`);

                    return $.ajax({
                        url: restUrl,
                        async: true,
                        method: 'POST',
                        contentType: 'application/json',
                        dataType: 'json',
                        data: JSON.stringify({ DestinationId: folderId }),
                        headers: { 'Authorization': `Bearer ${this.token}` }
                    })
                });

            this.collateRequests(requests, (results: MessageJson[]) => {
                console.log(`Finished moving items to other folder: ${results.length}`);
                this.onMoveComplete(results.length);
            }, (message: string) => {
                this.onError(message);
            });
        }
    }

    export class Model implements Data.IModel {
        private context?: Context;

        constructor(mailbox: Data.Mailbox) {
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