import * as React from "react";

import { Data } from "../Data/Model"

export interface SearchResultsProps {
    matches: Data.Match[];
    onSelection: (folderId: string) => void;
}

export class SearchResults extends React.Component<SearchResultsProps, undefined> {
    constructor(props: SearchResultsProps) {
        super(props);
        this.onClickFolder = this.handleClick.bind(this);
    }

    render() {
        if (!this.props.matches || this.props.matches.length === 0) {
            return null;
        }

        let rows: JSX.Element[] = [];

        this.props.matches.map((value: Data.Match, index: number) => {
            rows.push(<tr key={index}>
                <td><a name={value.folder.Id} onClick={this.onClickFolder}>{value.folder.DisplayName}</a></td>
                <td>{value.message.Sender}</td>
                <td>{value.message.ToRecipients}</td>
                <td>{value.message.BodyPreview}</td>
            </tr>);
        });

        return (<table>
            <thead>
                <tr>
                    <th>Folder</th>
                    <th>From</th>
                    <th>To</th>
                    <th>Preview</th>
                </tr>
            </thead>
            <tbody>
                {rows}
            </tbody>
        </table>);
    }

    private handleClick(evt: React.MouseEvent<HTMLAnchorElement>) {
        this.props.onSelection(evt.currentTarget.name);
        evt.preventDefault();
    }

    private onClickFolder: React.EventHandler<React.MouseEvent<HTMLAnchorElement>>;
}