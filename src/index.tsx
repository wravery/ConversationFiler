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
            const notificationKey = 'conversationFilerNotification';

            console.log('Starting to load the conversation...');

            data.getItemsAsync((results) => {
                console.log(`Loaded the conversation: ${results.length}`);

                if (results.length === 0) {
                    // Special case for empty results
                    mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                        message: `It looks like you haven't filed this conversation anywhere before.`
                    });
                    event.completed();

                    return;
                }

                console.log('Showing the dialog...');

                window.localStorage.setItem(storageKey, JSON.stringify(results));
                Office.context.ui.displayDialogAsync(window.location.href.replace(functionsRegex, "dialog.html"), { height: 25, width: 50, displayInIframe: true }, (result) => {
                    const dialog = result.value as Office.DialogHandler;
                    const onDialogComplete = (closed: Boolean) => {
                        mailbox.item.notificationMessages.removeAsync(notificationKey);
                        window.localStorage.removeItem(storageKey);

                        if (!closed) {
                            dialog.close();
                        }

                        event.completed();
                    };

                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (dialogEvent: { message: string }) => {
                        console.log('Moving the items...');

                        data.moveItemsAsync(dialogEvent.message, (count) => {
                            console.log(`Finished moving the items: ${count}`);

                            onDialogComplete(false);
                        }, (message) => {
                            console.log(`Error moving the items: ${message}`);

                            onDialogComplete(false);
                        });
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                        onDialogComplete(true);
                    });
                });
            }, (progress) => {
                console.log(`Progress loading the conversation: ${Data.Progress[progress]}`);

                // Update the progress indicator
                mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                    type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
                    message: 'Finding the messages in this conversation...'
                });
            }, (message) => {
                console.log(`Error loading the conversation: ${message}`);

                // Display an error
                mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                    message: `Sorry, I couldn't figure out where this message should go.`
                });

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
