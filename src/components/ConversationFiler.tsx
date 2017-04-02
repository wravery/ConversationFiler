/// <reference path="../../node_modules/@types/office-js/index.d.ts" />

import * as React from "react";
import * as JQuery from "jquery";

import { Data } from "../Data/Model"
import { Factory } from "../Data/Factory";

import { StatusMessage } from "./StatusMessage";
import { SearchResults } from "./SearchResults";
import { Feedback } from "./Feedback";

export interface ConversationFilerProps {
    mailbox: Office.Mailbox;
    storedResults?: Data.Match[];
    onComplete?: (folderId: string) => void;
}

interface ConversationFilerState {
    progress: Data.Progress;
    data?: Data.IModel;
    error?: string;
    matches?: Data.Match[];
}

export class ConversationFiler extends React.Component<ConversationFilerProps, ConversationFilerState> {
    constructor(props: ConversationFilerProps) {
        super(props);
        this.state = { progress: Data.Progress.GetCallbackToken };
    }

    // Start the chain of requests by getting a callback token.
    componentDidMount() {
        if (this.props.storedResults) {
            if (this.props.storedResults.length > 0) {
                this.setState({ progress: Data.Progress.Success, matches: this.props.storedResults });
            } else {
                this.setState({ progress: Data.Progress.NotFound });
            }

            return;
        } else if (!this.props.mailbox) {
            return;
        }

        const data = Factory.getData(this.props.mailbox);

        this.setState({ data: data });

        data.getItemsAsync((results) => {
            if (results.length > 0) {
                this.setState({ progress: Data.Progress.Success, matches: results });
            } else {
                this.setState({ progress: Data.Progress.NotFound });
            }
        }, (progress) => {
            this.setState({ progress: progress });
        }, (message) => {
            this.setState({ progress: Data.Progress.Error, error: message });
        });
    }

    private onSelection(folderId: string) {
        console.log(`Selected a folder: ${folderId}`);

        if (!this.state.data) {
            // Handle the dialog or test case by just notifying the client
            if (this.props.onComplete) {
                this.props.onComplete(folderId);
            }

            return;
        }

        this.state.data.moveItemsAsync(folderId, (count) => {
            if (this.props.onComplete) {
                this.props.onComplete(folderId);
            }
        }, (message) => {
            this.setState({ progress: Data.Progress.Error, error: message });
        });
    }

    render() {
        return (<div>
            <StatusMessage progress={this.state.progress} message={this.state.error} />
            <SearchResults matches={this.state.matches} onSelection={this.onSelection.bind(this)} />
            <Feedback />
        </div>);
    }
}
