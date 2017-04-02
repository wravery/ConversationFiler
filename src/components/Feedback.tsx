import * as React from "react";

export class Feedback extends React.Component<undefined, undefined> {
    render() {
        return (<div className="feedback">
            <a href="https://beandotnet.azurewebsites.net/">about this app</a>
            &nbsp;
            <a href="mailto:wravery@hotmail.com?Subject=Conversation%20Filer%20v2.0%20App%20for%20Outlook">send feedback</a>
        </div>);
    }
}