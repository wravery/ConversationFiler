import * as React from "react";

import { Data } from "../Data/Model";

export interface StatusMessageProps {
    progress: Data.Progress;
    message?: string;
}

export class StatusMessage extends React.Component<StatusMessageProps, undefined> {
    render() {
        let className: string;
        let status: string;

        switch (this.props.progress) {
            case Data.Progress.GetCallbackToken:
            case Data.Progress.GetConversation:
            case Data.Progress.GetExcludedFolders:
            case Data.Progress.GetFolderNames:
                return <h3>Looking for other messages in this conversation...</h3>;

            case Data.Progress.Success:
                return null;

            case Data.Progress.NotFound:
                return <h3>It looks like you haven't filed this conversation anywhere before.</h3>;

            default:
                return (<div>
                    <h3>Sorry, I couldn't figure out where this message should go. :(</h3>
                    <span>{this.props.message}</span>
                </div>);
        }
    }
}
