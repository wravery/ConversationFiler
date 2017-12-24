import * as React from "react";
import * as ReactDOM from "react-dom";
import * as TestUtils from "react-dom/test-utils";

import { Data } from "../Data/Model";

import { ConversationFilerDialog } from "./ConversationFilerDialog";

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

    const onComplete = (selected: string) => {
        fail("Just rendering the dialog should not invoke any callbacks");
    };

    const onCancel = () => {
        fail("The dialog should not be canceled");
    }

    const component = <ConversationFilerDialog storedResults={dummyResults} onComplete={onComplete} onCancel={onCancel} />;
    const element = TestUtils.renderIntoDocument(component);
    const rendered = ReactDOM.findDOMNode(element as React.ReactInstance);
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

    const onCancel = () => {
        fail("The dialog should not be canceled");
    }

    const component = <ConversationFilerDialog storedResults={dummyResults} onComplete={onComplete} onCancel={onCancel} />;
    const element = TestUtils.renderIntoDocument(component);
    const rendered = ReactDOM.findDOMNode(element as React.ReactInstance);
    expect(rendered.innerHTML).toMatchSnapshot();

    component.props.onComplete(folderId);
    expect(selectedId).toBe(folderId);
});

test("Cancel button callback works", () => {
    const folderId = 'folderId3';
    let canceled = false;

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
        fail("No folder should be selected");
    };

    const onCancel = () => {
        canceled = true;
    }

    const component = <ConversationFilerDialog storedResults={dummyResults} onComplete={onComplete} onCancel={onCancel} />;
    const element = TestUtils.renderIntoDocument(component);
    const rendered = ReactDOM.findDOMNode(element as React.ReactInstance);
    expect(rendered.innerHTML).toMatchSnapshot();

    component.props.onCancel();
    expect(canceled).toBe(true);
});
