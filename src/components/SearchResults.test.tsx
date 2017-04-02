/// <reference path="../_testReferences.ts" />

import * as React from "react";
import * as renderer from "react-test-renderer";

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

    const component = renderer.create(<SearchResults matches={dummyResults} onSelection={onSelection} />);
    const tree = component.toJSON();
    expect(tree).toMatchSnapshot();
});

test("SearchResults onSelection callback should work", () => {
    const folderId = 'folderId3';
    let selectedId: string = null;

    const onSelection = (selected: string) => {
        selectedId = selected;
    };

    const element = <SearchResults matches={[]} onSelection={onSelection} />;
    const component = renderer.create(element);
    const tree = component.toJSON();
    expect(tree).toMatchSnapshot();

    element.props.onSelection(folderId);
    expect(selectedId).toBe(folderId);
});
