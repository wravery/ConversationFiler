import * as React from "react";
import * as ReactDOM from "react-dom";
import * as TestUtils from "react-addons-test-utils";

import { Data } from "../Data/Model";

import { ConversationFiler } from "./ConversationFiler";

test("Loading", () => {
    const component = TestUtils.renderIntoDocument(<ConversationFiler mailbox={null} storedResults={null} />) as ConversationFiler;
    const rendered = ReactDOM.findDOMNode(component);
    expect(rendered.innerHTML).toMatchSnapshot();
});

test("Empty results", () => {
    const component = TestUtils.renderIntoDocument(<ConversationFiler mailbox={null} storedResults={[]} />) as ConversationFiler;
    const rendered = ReactDOM.findDOMNode(component);
    expect(rendered.innerHTML).toMatchSnapshot();
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

    const component = TestUtils.renderIntoDocument(<ConversationFiler mailbox={null} storedResults={dummyResults} />) as ConversationFiler;
    const rendered = ReactDOM.findDOMNode(component);
    expect(rendered.innerHTML).toMatchSnapshot();
});

test("Folder selection callback works", () => {
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

    const onComplete = (selected: string) => {
        selectedId = selected;
    };

    const component = TestUtils.renderIntoDocument(<ConversationFiler mailbox={null} storedResults={dummyResults} onComplete={onComplete} />) as ConversationFiler;
    const rendered = ReactDOM.findDOMNode(component);
    expect(rendered.innerHTML).toMatchSnapshot();

    component.props.onComplete(folderId);
    expect(selectedId).toBe(folderId);
});
