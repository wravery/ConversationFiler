export module Data {
    export interface Message {
        Id: string;
        BodyPreview: string;
        Sender: string;
        ToRecipients: string;
        ParentFolderId: string;
    }

    export interface Folder {
        Id: string;
        DisplayName: string;
    }

    export interface Match {
        message: Message;
        folder: Folder;
    }

    export enum Progress {
        GetCallbackToken,
        GetConversation,
        GetExcludedFolders,
        GetFolderNames,
        Success,
        NotFound,
        Error
    }

    export interface IModel {
        getItemsAsync(onLoadComplete: (results: Match[]) => void, onProgress: (progress: Progress) => void, onError: (message: string) => void): void;
        moveItemsAsync(folderId: string, onMoveComplete: (count: number) => void, onError: (message: string) => void): void;
    }

    export function removeDuplicates(results: Match[], itemId: string) {
        if (itemId) {
            const item = results
                .filter(item => item.message.Id === itemId)
                .pop();

            if (item) {
                console.log(`Removing duplicates: ${results.length} original message(s)`);

                // Find all of the items in the same folder, we want to remove all of their duplicates.
                const sameFolderItems = results.filter(result => result.message.ParentFolderId === item.message.ParentFolderId);

                console.log(`Messages in the same folder: ${sameFolderItems.length} message(s)`);

                // Remove all items that are either in the same folder or match an item in the same folder.
                results = results.filter(result => {
                    if (result.message.ParentFolderId === item.message.ParentFolderId) {
                        console.log(`Removed message in the same folder: ${item.message.Id}`);
                        return false;
                    }

                    const isDuplicate = sameFolderItems.reduce((previousValue, value) => previousValue ||
                        (result.message.Sender === value.message.Sender &&
                            result.message.ToRecipients === value.message.ToRecipients &&
                            result.message.BodyPreview === value.message.BodyPreview), false);

                    if (isDuplicate) {
                        console.log(`Removed duplicate message in other folder: ${item.message.Id}`);
                    }

                    return !isDuplicate;
                });

                console.log(`Remaining in other folders: ${results.length} message(s)`);
            }
        }

        return results;
    }
}