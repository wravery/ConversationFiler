import * as React from "react";
import * as ReactDOM from "react-dom";
import * as TestUtils from "react-dom/test-utils";

import { Data } from "../Data/Model";

import { SearchResults } from "./SearchResults";

test("SearchResults should render", () => {
    const dummyResults: Data.Match[] = [{
        folder: {
            Id: 'folderId1',
            DisplayName: 'Folder 1'
        },
        message: {
            Id: 'messageId1',
            BodyPreview: 'Here\'s a preview of a message body',
            Sender: 'Foo Bar',
            ToRecipients: 'Baz Bar',
            ParentFolderId: 'folderId1'
        }
    }, {
        folder: {
            Id: 'folderId2',
            DisplayName: 'Folder 2'
        },
        message: {
            Id: 'messageId2',
            BodyPreview: 'Here\'s another message body',
            Sender: 'Baz Bar',
            ToRecipients: 'Foo Bar',
            ParentFolderId: 'folderId2'
        }
    }];

    const onSelection = (folderId: string) => {
        fail('Just rendering the dummy data should not invoke a callback.');
    };

    const component = <SearchResults matches={dummyResults} onSelection={onSelection} />;
    const element = TestUtils.renderIntoDocument(component);
    const rendered = ReactDOM.findDOMNode(element as React.ReactInstance);
    expect(rendered.innerHTML).toMatchSnapshot();
});

test("SearchResults onSelection callback should work", () => {
    const folderId = 'folderId3';
    let selectedId: string = null;

    const dummyResults: Data.Match[] = [{
        folder: {
            Id: 'folderId3',
            DisplayName: 'Folder 3'
        },
        message: {
            Id: 'messageId3',
            BodyPreview: 'Click Me!',
            Sender: 'Foo Bar',
            ToRecipients: 'Baz Bar',
            ParentFolderId: 'folderId3'
        }
    }];

    const onSelection = (selected: string) => {
         selectedId = selected;
     };
 
    const component = <SearchResults matches={dummyResults} onSelection={onSelection} />;
    const element = TestUtils.renderIntoDocument(component);
    const rendered = ReactDOM.findDOMNode(element as React.ReactInstance);
    expect(rendered.innerHTML).toMatchSnapshot();

    component.props.onSelection(folderId);
    expect(selectedId).toBe(folderId);
});
