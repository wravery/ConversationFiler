import * as React from "react";
import * as ReactDOM from "react-dom";
import * as TestUtils from "react-addons-test-utils";

import { Data } from "../Data/Model";
import { Factory } from "../Data/__mocks__/Factory";
jest.mock('../Data/Factory');

import { ConversationFilerPage } from "./ConversationFilerPage";

test("Loading", () => {
    const mockMailbox = {
        restUrl: 'https://foo.bar.com/api'
    } as Factory.MockMailbox;

    const log = console.log;
    console.log = jest.fn((message: string) => {
        expect(message).toMatchSnapshot();
    });

    const component = TestUtils.renderIntoDocument(<ConversationFilerPage mailbox={mockMailbox} />) as ConversationFilerPage;
    const rendered = ReactDOM.findDOMNode(component);
    expect(rendered.innerHTML).toMatchSnapshot();

    console.log = log;
});

test("Empty results", () => {
    const mockMailbox = {
        restUrl: 'https://foo.bar.com/api',
        mockResults: []
    } as Factory.MockMailbox;

    const log = console.log;
    console.log = jest.fn((message: string) => {
        expect(message).toMatchSnapshot();
    });

    const component = TestUtils.renderIntoDocument(<ConversationFilerPage mailbox={mockMailbox} />) as ConversationFilerPage;
    const rendered = ReactDOM.findDOMNode(component);
    expect(rendered.innerHTML).toMatchSnapshot();

    console.log = log;
});

test("Dummy data", () => {
    const mockMailbox = {
        restUrl: 'https://foo.bar.com/api',
        mockResults: [{
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
        }]
    } as Factory.MockMailbox;

    const log = console.log;
    console.log = jest.fn((message: string) => {
        expect(message).toMatchSnapshot();
    });

    const component = TestUtils.renderIntoDocument(<ConversationFilerPage mailbox={mockMailbox} />) as ConversationFilerPage;
    const rendered = ReactDOM.findDOMNode(component);
    expect(rendered.innerHTML).toMatchSnapshot();

    console.log = log;
});
