import * as React from "react";
import * as ReactDOM from "react-dom";

import { Pages } from "./pages";
import { AppFunctions } from "./functions";

import { ConversationFiler } from "./components/ConversationFiler";

Office.initialize = function () {
    if (Pages.shouldHaveUI()) {
        // Show the UI...
        const params = Pages.getUIParameters();

        ReactDOM.render(
            <ConversationFiler mailbox={params.mailbox} onComplete={params.onComplete} storedResults={params.storedResults} />,
            document.getElementById("conversationFilerRoot")
        );
    } else {
        AppFunctions.register();
    }
};
