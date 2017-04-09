import * as React from "react";
import * as ReactDOM from "react-dom";

import { ConversationFilerPage } from "./components/ConversationFilerPage";

Office.initialize = function () {
    ReactDOM.render(
        <ConversationFilerPage mailbox={Office.context.mailbox} />,
        document.getElementById("conversationFilerRoot")
    );
};
