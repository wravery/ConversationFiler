import { Data } from "./Data/Model";

export module Pages {
    const functionsRegex = /\/functions\.html(\?.*)?$/i;
    const dialogRegex = /\/dialog\.html(\?.*)?$/i;
    const storageKey = "conversationFilerMatches";

    export function shouldHaveUI(): boolean {
        return !functionsRegex.test(window.location.pathname);
    }

    export function getDialogUrl(): string {
        return window.location.href.replace(functionsRegex, "/dialog.html");
    }

    export function getAboutUrl(): string {
        return window.location.href.replace(functionsRegex, "/about.html");
    }

    export function populateDialog(storedResults: Data.Match[]) {
        window.localStorage.setItem(storageKey, JSON.stringify(storedResults));
    }

    export function resetDialog() {
        window.localStorage.removeItem(storageKey);
    }

    export interface UIParameters {
        mailbox?: Office.Mailbox,
        onComplete?: (folderId: string) => void;
        onCancel?: () => void;
        storedResults?: Data.Match[];
    }

    export function getUIParameters(): UIParameters {
        if (dialogRegex.test(window.location.pathname)) {
            return {
                onComplete: folderId => { Office.context.ui.messageParent(folderId); },
                onCancel: () => { Office.context.ui.messageParent(""); },
                storedResults: <Data.Match[]>JSON.parse(window.localStorage.getItem(storageKey))
            };
        }
        else {
            return {
                mailbox: Office.context.mailbox,
            };
        }
    }
}
