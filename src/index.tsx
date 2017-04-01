/// <reference path="_references.ts" />
/// <reference path="./components/ConversationFiler.tsx" />

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Data } from "./Data/Model"

import { ConversationFiler } from "./components/ConversationFiler";

Office.initialize = function () {
    const functionsRegex = /functions\.html(\?.*)?$/i;
    const noUI = functionsRegex.test(window.location.pathname);

    if (noUI) {
        // Add the UI-less function callback if we're loaded from functions.html instead of index.html
        (window as any).fileDialog = function (event: any) {
            Office.context.ui.displayDialogAsync(window.location.href.replace(functionsRegex, "dialog.html"), { height: 25, width: 80, displayInIframe: true }, (result) => {
                const dialog = result.value as Office.DialogHandler;

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
                    dialog.close();
                    event.completed();
                });
            });
        };

        return;
    }

    // Show the UI...
    const mailbox = (Office.context || ({} as Office.Context)).mailbox;
    let onComplete: () => void;

    if (mailbox && /dialog\.html(\?.*)?$/i.test(window.location.pathname)) {
        // When we finish moving the items, we want to dismiss the dialog and complete the callback
        onComplete = () => {
            Office.context.ui.messageParent(true);
        };
    }

    ReactDOM.render(
        <ConversationFiler mailbox={mailbox} onComplete={onComplete} />,
        document.getElementById("conversationFilerRoot")
    );

    // ...and if we're running outside of an Outlook client, run through the tests
    if (!mailbox) {
        let testEmpty = function () {
            console.log("Testing the behavior with an empty set of matches...");

            // Need to clear out the DOM so it will mount a new ConversationFiler
            ReactDOM.render(
                <div>Testing...</div>,
                document.getElementById("conversationFilerRoot")
            );

            ReactDOM.render(
                <ConversationFiler mailbox={null} mockResults={[]} />,
                document.getElementById("conversationFilerRoot")
            );

            window.setTimeout(testDummy, 3000);
        }

        let testDummy = function () {
            console.log("Testing the behavior with a set of mock matches...");

            // Need to clear out the DOM so it will mount a new ConversationFiler
            ReactDOM.render(
                <div>Testing...</div>,
                document.getElementById("conversationFilerRoot")
            );

            let mockResults: Data.Match[] = [{
                    folder: {
                        Id: 'folderId1',
                        DisplayName: 'Folder 1'
                    },
                    message: {
                        Id: 'messageId1',
                        BodyPreview: 'Here\'s a preview of a message body',
                        Sender: 'Foo Bar',
                        ToRecipients: 'Baz Bar',
                        ParentFolderId: 'folderId1'
                    }
                }, {
                    folder: {
                        Id: 'folderId2',
                        DisplayName: 'Folder 2'
                    },
                    message: {
                        Id: 'messageId2',
                        BodyPreview: 'Here\'s another message body',
                        Sender: 'Baz Bar',
                        ToRecipients: 'Foo Bar',
                        ParentFolderId: 'folderId2'
                    }
                }];

            ReactDOM.render(
                <ConversationFiler mailbox={null} mockResults={mockResults} />,
                document.getElementById("conversationFilerRoot")
            );
        }

        window.setTimeout(testEmpty, 3000);
    }
};
