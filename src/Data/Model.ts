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
                // Find all of the items in the same folder, we want to remove all of their duplicates.
                const sameFolderItems = results.filter(result => result.message.ParentFolderId === item.message.ParentFolderId);

                // Remove all items that are either in the same folder or match an item in the same folder.
                results = results.filter(result => {
                    if (result.message.ParentFolderId === item.message.ParentFolderId) {
                        return false;
                    }

                    return !sameFolderItems.reduce((previousValue, value) => {
                        if (previousValue) {
                            return true;
                        }

                        return result.message.Sender === value.message.Sender &&
                            result.message.ToRecipients === value.message.ToRecipients &&
                            result.message.BodyPreview === value.message.BodyPreview;
                    }, false);
                });
            }
        }

        return results;
    }
}