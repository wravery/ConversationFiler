/// <reference path="../node_modules/@types/office-js/index.d.ts" />

import { Pages } from "./pages";

import { Data } from "./Data/Model";
import { Factory } from "./Data/Factory";

export module AppFunctions {
    function fileDialog(event: any) {
        const mailbox = Office.context.mailbox;
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

            mailbox.item.notificationMessages.removeAsync(notificationKey);
            Pages.populateDialog(results);
            Office.context.ui.displayDialogAsync(Pages.getDialogUrl(), { height: 25, width: 50, displayInIframe: true }, (result) => {
                const dialog = <Office.DialogHandler>result.value;
                const onDialogComplete = (closed: boolean) => {
                    Pages.resetDialog();

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

                        mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                            message: `Something went wrong, I couldn't move the messages.`
                        });
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
    }

    // Add the UI-less function callbacks to the window
    export function register() {
        (<any>window).fileDialog = fileDialog;
    }
}