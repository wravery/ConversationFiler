import { Data } from "./Data/Model";

export module DialogMessages {
    const storageKey = "conversationFilerMatches";

    export function saveDialog(storedResults: Data.Match[]) {
        window.localStorage.setItem(storageKey, JSON.stringify(storedResults));
    }

    export function loadDialog() {
        return JSON.parse(window.localStorage.getItem(storageKey)) as Data.Match[];
    }

    export function resetDialog() {
        window.localStorage.removeItem(storageKey);
    }

    export interface FileDialogMessage {
        canceled: boolean;
        folderId?: string;
    }

    export interface AboutDialogMessage {
        canceled: boolean;
    }
}
