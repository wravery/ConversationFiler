import * as React from "react";
import * as JQuery from "jquery";

import { Data } from "../Data/Model"
import { Factory } from "../Data/Factory";

import { StatusMessage } from "./StatusMessage";
import { SearchResults } from "./SearchResults";
import { Label } from "office-ui-fabric-react/lib/Label";
import { DefaultButton, IButtonProps } from "office-ui-fabric-react/lib/Button";

export interface ConversationFilerDialogProps {
    storedResults: Data.Match[];
    onComplete: (folderId: string) => void;
    onCancel: () => void;
}

export class ConversationFilerDialog extends React.Component<ConversationFilerDialogProps, undefined> {
    constructor(props: ConversationFilerDialogProps) {
        super(props);
        this.onSelection = this.onSelection.bind(this);
        this.onCancel = this.onCancel.bind(this);
    }

    private onSelection(folderId: string) {
        console.log(`Selected a folder: ${folderId}`);
        this.props.onComplete(folderId);
    }

    private onCancel() {
        console.log("Cancel button clicked");
        this.props.onCancel();
    }

    render() {
        return (<div>
            <StatusMessage progress={Data.Progress.Success} />
            <SearchResults matches={this.props.storedResults} onSelection={this.onSelection} />
            <div className="dialogButtons">
                <DefaultButton
                    ariaDescription="Cancel the operation without moving any items."
                    onClick={this.onCancel}>
                    Cancel
                </DefaultButton>
            </div>
        </div>);
    }
}
