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
}