import * as React from "react";
import * as JQuery from "jquery";

import { Data } from "../Data/Model"
import { Factory } from "../Data/Factory";

import { StatusMessage } from "./StatusMessage";
import { SearchResults } from "./SearchResults";
import { Feedback } from "./Feedback";

export interface ConversationFilerPageProps {
    mailbox: Office.Mailbox;
}

interface ConversationFilerPageState {
    progress: Data.Progress;
    data?: Data.IModel;
    error?: string;
    matches?: Data.Match[];
}

export class ConversationFilerPage extends React.Component<ConversationFilerPageProps, ConversationFilerPageState> {
    constructor(props: ConversationFilerPageProps) {
        super(props);
        this.onSelection = this.onSelection.bind(this);

        this.state = { progress: Data.Progress.GetCallbackToken };
    }

    // Start the chain of requests to load the data
    componentDidMount() {
        const data = Factory.getData(this.props.mailbox);
 
        this.setState({ data: data });

        console.log('Starting to load the conversation...');

        data.getItemsAsync((results) => {
            console.log(`Loaded the conversation: ${results.length}`);

            if (results.length > 0) {
                this.setState({ progress: Data.Progress.Success, matches: results });
            } else {
                this.setState({ progress: Data.Progress.NotFound });
            }
        }, (progress) => {
            console.log(`Progress loading the conversation: ${Data.Progress[progress]}`);

            this.setState({ progress: progress });
        }, (message) => {
            console.log(`Error loading the conversation: ${message}`);

            this.setState({ progress: Data.Progress.Error, error: message });
        });
    }

    private onSelection(folderId: string) {
        console.log(`Selected a folder: ${folderId}`);

        this.state.data.moveItemsAsync(folderId, (count) => {
            console.log(`Finished moving the items: ${count}`);
        }, (message) => {
            this.setState({ progress: Data.Progress.Error, error: message });
        });
    }

    render() {
        return (<div>
            <StatusMessage progress={this.state.progress} message={this.state.error} />
            <SearchResults matches={this.state.matches} onSelection={this.onSelection} />
            <Feedback />
        </div>);
    }
}
