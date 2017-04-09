import * as React from "react";
import * as ReactDOM from "react-dom";

import { DialogMessages } from "./messages";

import { ConversationFilerDialog } from "./components/ConversationFilerDialog";

module DialogCallbacks {
    export function onComplete(folderId: string) {
        const message: DialogMessages.FileDialogMessage = {
            canceled: false,
            folderId: folderId
        };

        Office.context.ui.messageParent(JSON.stringify(message));
    }

    export function onCancel() {
        const message: DialogMessages.FileDialogMessage = {
            canceled: true
        };

        Office.context.ui.messageParent(JSON.stringify(message));
    }
}

Office.initialize = function () {
    ReactDOM.render(
        <ConversationFilerDialog onComplete={DialogCallbacks.onComplete} onCancel={DialogCallbacks.onCancel} storedResults={DialogMessages.loadDialog()} />,
        document.getElementById("conversationFilerRoot")
    );
};
