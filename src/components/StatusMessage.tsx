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
                return <span className="ms-font-l">Looking for other messages in this conversation...</span>;

            case Data.Progress.Success:
                return null;

            case Data.Progress.NotFound:
                return <span className="ms-font-l">It looks like you haven't filed this conversation anywhere before.</span>;

            default:
                return (<div>
                    <span className="ms-font-l">Sorry, I couldn't figure out where this message should go. :(</span>
                    <br />
                    <span className="ms-font-m">{this.props.message}</span>
                </div>);
        }
    }
}
