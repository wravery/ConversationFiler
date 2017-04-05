/// <reference path="../../node_modules/@types/jest/index.d.ts" />
/// <reference path="../../node_modules/@types/react-test-renderer/index.d.ts" />

import * as React from "react";
import * as TestUtils from "react-addons-test-utils";

import { Data } from "../Data/Model";

import { ConversationFiler } from "./ConversationFiler";

test("Loading", () => {
    const component = TestUtils.renderIntoDocument(<ConversationFiler mailbox={null} storedResults={null} />) as React.ReactInstance;
    //const tree = component.toJSON();
    //expect(tree).toMatchSnapshot();
});

test("Empty results", () => {
    const component = TestUtils.renderIntoDocument(<ConversationFiler mailbox={null} storedResults={[]} />) as React.ReactInstance;
    //const tree = component.toJSON();
    //expect(tree).toMatchSnapshot();
});

test("Dummy data", () => {
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

    const component = TestUtils.renderIntoDocument(<ConversationFiler mailbox={null} storedResults={dummyResults} />) as React.ReactInstance;
    //const tree = component.toJSON();
    //expect(tree).toMatchSnapshot();
});

test("Folder selection callback works", () => {
    const folderId = 'folderId3';
    let selectedId: string = null;

    const onComplete = (selected: string) => {
        selectedId = selected;
    };

    const element = <ConversationFiler mailbox={null} storedResults={[]} onComplete={onComplete} />;
    const component = TestUtils.renderIntoDocument(element) as React.ReactInstance;
    //const tree = component.toJSON();
    //expect(tree).toMatchSnapshot();

    element.props.onComplete(folderId);
    expect(selectedId).toBe(folderId);
});
