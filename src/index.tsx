import * as React from "react";
import * as ReactDOM from "react-dom";

import { Pages } from "./pages";
import { AppFunctions } from "./functions";

import { ConversationFilerPage } from "./components/ConversationFilerPage";
import { ConversationFilerDialog } from "./components/ConversationFilerDialog";

Office.initialize = function () {
    if (Pages.shouldHaveUI()) {
        // Show the UI...
        const params = Pages.getUIParameters();

        if (params.mailbox) {
            ReactDOM.render(
                <ConversationFilerPage mailbox={params.mailbox} />,
                document.getElementById("conversationFilerRoot")
            );
        } else {
            ReactDOM.render(
                <ConversationFilerDialog onComplete={params.onComplete} onCancel={params.onCancel} storedResults={params.storedResults} />,
                document.getElementById("conversationFilerRoot")
            );
        }
    } else {
        AppFunctions.register();
    }
};
