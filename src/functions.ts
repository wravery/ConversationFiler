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
                const onDialogComplete = () => {
                    Pages.resetDialog();
                    event.completed();
                };

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (dialogEvent: { message: string }) => {
                    console.log('Moving the items...');

                    dialog.close();
                    mailbox.item.notificationMessages.replaceAsync(notificationKey, {
                        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
                        message: 'Moving the items in this conversation...'
                    });

                    data.moveItemsAsync(dialogEvent.message, (count) => {
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

    // Add the UI-less function callbacks to the window
    export function register() {
        (<any>window).fileDialog = fileDialog;
    }
}
