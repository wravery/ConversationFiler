import * as React from "react";
import * as ReactDOM from "react-dom";

import { DialogMessages } from "./messages";

import { Data } from "./Data/Model";
import { Factory } from "./Data/Factory";

import { ConversationFilerDialog } from "./components/ConversationFilerDialog";

module ButtonFunctions {
    const functionsRegex = /\/functions\.html(\?.*)?$/i;

    function getDialogUrl(): string {
        return window.location.href.replace(functionsRegex, "/dialog.html");
    }

    function getAboutUrl(): string {
        return window.location.href.replace(functionsRegex, "/about.html");
    }

    export function fileDialog(event: any) {
        const mailbox = Office.context.mailbox;
        const data = Factory.getData(<any>mailbox);
        const notificationKey = 'conversationFilerNotification';

        console.log('Starting to load the conversation...');

        data.getItemsAsync((results) => {
            console.log(`Loaded the conversation: ${results.length}`);

            if (results.length === 0) {
                // Special case for empty results
                mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: `It looks like you haven't filed this conversation anywhere before.`,
                    icon: 'file-icon-16',
                    persistent: false
                });
                event.completed();

                return;
            }

            console.log('Showing the dialog...');

            mailbox.item.notificationMessages.removeAsync(notificationKey);
            DialogMessages.saveDialog(results);
            Office.context.ui.displayDialogAsync(getDialogUrl(), { height: 40, width: 50, displayInIframe: true }, (result) => {
                const dialog = <Office.DialogHandler>result.value;
                const onDialogComplete = () => {
                    DialogMessages.resetDialog();
                    event.completed();
                };

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (dialogEvent: { message: string }) => {
                    const message = <DialogMessages.FileDialogMessage>JSON.parse(dialogEvent.message);

                    if (message.canceled) {
                        console.log('Dialog canceled');

                        dialog.close();
                        onDialogComplete();
                        return;
                    }

                    console.log('Moving the items...');

                    dialog.close();
                    mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
                        message: 'Moving the items in this conversation...'
                    });

                    data.moveItemsAsync(message.folderId, (count) => {
                        console.log(`Finished moving the items: ${count}`);

                        mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                            message: 'I moved the items in this conversation, but there might be a short delay before that shows up in Outlook.',
                            icon: 'file-icon-16',
                            persistent: false
                        });

                        onDialogComplete();
                    }, (message) => {
                        console.log(`Error moving the items: ${message}`);

                        mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                            message: `Something went wrong, I couldn't move the messages.`
                        });
                        onDialogComplete();
                    });
                });

                dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                    onDialogComplete();
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
    }

    export function aboutDialog(event: any) {
        Office.context.ui.displayDialogAsync(getAboutUrl(), { height: 40, width: 25, displayInIframe: true }, (result) => {
            const dialog = <Office.DialogHandler>result.value;
            const onDialogComplete = () => {
                event.completed();
            };

            dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
                console.log('Dialog closed with button');

                dialog.close();
                event.completed();
            });

            dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                console.log('Dialog closed');

                event.completed();
            });
        });
    }

    export function sendFeedback(event: any) {
        Office.context.mailbox.displayNewMessageForm({
            toRecipients: [ 'wravery@hotmail.com' ],
            subject: 'Conversation Filer v2.0 App for Outlook'
        });

        event.completed();
    }
}

Office.initialize = function () {
    // Add the UI-less function callbacks to the window
    (<any>window).fileDialog = ButtonFunctions.fileDialog;
    (<any>window).aboutDialog = ButtonFunctions.aboutDialog;
    (<any>window).sendFeedback = ButtonFunctions.sendFeedback;
};
