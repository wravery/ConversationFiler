/// <reference path="_references.ts" />
/// <reference path="./components/ConversationFiler.tsx" />

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Data } from "./Data/Model";
import { Factory } from "./Data/Factory";

import { ConversationFiler } from "./components/ConversationFiler";

Office.initialize = function () {
    const functionsRegex = /functions\.html(\?.*)?$/i;
    const noUI = functionsRegex.test(window.location.pathname);
    const mailbox = (Office.context || ({} as Office.Context)).mailbox;
    const storageKey = "conversationFilerMatches";

    if (noUI) {
        // Add the UI-less function callback if we're loaded from functions.html instead of index.html
        (window as any).fileDialog = function (event: any) {
            const data = Factory.getData(mailbox);

            data.getItemsAsync((results) => {
                window.localStorage.setItem(storageKey, JSON.stringify(results));

                Office.context.ui.displayDialogAsync(window.location.href.replace(functionsRegex, "dialog.html"), { height: 25, width: 50, displayInIframe: true }, (result) => {
                    const dialog = result.value as Office.DialogHandler;

                    const onDialogComplete = () => {
                        dialog.close();
                        window.localStorage.removeItem(storageKey);
                        event.completed();
                    };

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (dialogEvent: { message: string }) => {
                        data.moveItemsAsync(dialogEvent.message, (count) => {
                            onDialogComplete();
                        }, (message) => {
                            // no-op...
                            onDialogComplete();
                        });
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                        event.completed();
                    });
                });
            }, (progress) => {
                // no-op...
            }, (message) => {
                event.completed();
            });
        };

        return;
    }

    // Show the UI...
    let onComplete: (folderId: string) => void;
    let storedResults: Data.Match[];

    if (/dialog\.html(\?.*)?$/i.test(window.location.pathname)) {
        // When we finish moving the items, we want to dismiss the dialog and complete the callback
        onComplete = (folderId: string) => {
            Office.context.ui.messageParent(folderId);
        };

        storedResults = JSON.parse(window.localStorage.getItem(storageKey)) as Data.Match[];
    }

    ReactDOM.render(
        <ConversationFiler mailbox={mailbox} onComplete={onComplete} storedResults={storedResults} />,
        document.getElementById("conversationFilerRoot")
    );
};
