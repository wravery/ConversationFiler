import * as React from "react";

import { Data } from "../Data/Model";

import { IColumn } from "office-ui-fabric-react/lib/components/DetailsList/DetailsList.Props";
import { DetailsList } from "office-ui-fabric-react/lib/DetailsList";
import { Link } from "office-ui-fabric-react/lib/Link";

export interface SearchResultsProps {
    matches: Data.Match[];
    onSelection: (folderId: string) => void;
}

export class SearchResults extends React.Component<SearchResultsProps, undefined> {
    constructor(props: SearchResultsProps) {
        super(props);
        this.onClickFolder = this.onClickFolder.bind(this);
        this.onRenderLink = this.onRenderLink.bind(this);
        this.onRenderColumn = this.onRenderColumn.bind(this);
    }

    render() {
        if (!this.props.matches || this.props.matches.length === 0) {
            return null;
        }

        const columns: IColumn[] = [{
                key: 'DisplayName',
                name: 'Folder',
                fieldName: null,
                onRender: this.onRenderLink,
                minWidth: 5
            }, {
                key: 'Sender',
                name: 'From',
                fieldName: null,
                onRender: this.onRenderColumn,
                minWidth: 10
            }, {
                key: 'ToRecipients',
                name: 'To',
                fieldName: null,
                onRender: this.onRenderColumn,
                minWidth: 10
            }, {
                key: 'BodyPreview',
                name: 'Preview',
                fieldName: null,
                onRender: this.onRenderColumn,
                minWidth: 50
            }];

        return (<div>
            <h3>
                I found some items in this conversation filed in other folders. Click on one of the folders listed here to
                automatically reunite this part of the conversation with the ones that came before:
            </h3>
            <DetailsList columns={columns} items={this.props.matches} />
        </div>);
    }

    private onClickFolder(evt: React.MouseEvent<HTMLAnchorElement>) {
        this.props.onSelection(evt.currentTarget.href);
        evt.preventDefault();
    }

    private onRenderLink(item: Data.Match) {
        return <Link onClick={this.onClickFolder} name={item.folder.Id}>{item.folder.DisplayName}</Link>;
    }

    private onRenderColumn(item: Data.Match, index: number, column: IColumn) {
        return (item.message as any)[column.key] as string;
    }
}
