/// <reference path="../node_modules/@types/office-js/index.d.ts" />

import { Data } from "./Data/Model";

export module Pages {
    const functionsRegex = /functions\.html(\?.*)?$/i;
    const dialogRegex = /dialog\.html(\?.*)?$/i;
    const storageKey = "conversationFilerMatches";

    export function shouldHaveUI(): boolean {
        return !functionsRegex.test(window.location.pathname);
    }

    export function getDialogUrl(): string {
        return window.location.href.replace(functionsRegex, "dialog.html");
    }

    export interface UIParameters {
        mailbox?: Office.Mailbox,
        onComplete?: (folderId: string) => void;
        storedResults?: Data.Match[];
    }

    export function populateDialog(storedResults: Data.Match[]) {
        window.localStorage.setItem(storageKey, JSON.stringify(storedResults));
    }

    export function getUIParameters(): UIParameters {
        const params: UIParameters = {
            mailbox: (Office.context || (<Office.Context>{})).mailbox,
            onComplete: null,
            storedResults: null
        };

        if (dialogRegex.test(window.location.pathname)) {
            // When we finish moving the items, we want to dismiss the dialog and complete the callback
            params.onComplete = (folderId: string) => {
                Office.context.ui.messageParent(folderId);
            };

            params.storedResults = <Data.Match[]>JSON.parse(window.localStorage.getItem(storageKey));
        }

        return params;
    }

    export function resetDialog() {
        window.localStorage.removeItem(storageKey);
    }
}
