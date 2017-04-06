import * as React from "react";
import * as renderer from "react-test-renderer";

import { Data } from "../Data/Model";

import { StatusMessage } from "./StatusMessage";

test("StatusMessage (Loading)", () => {
    let component = renderer.create(<StatusMessage progress={Data.Progress.GetCallbackToken} />);
    let tree = component.toJSON();
    expect(tree).toMatchSnapshot();

    // The snapshot should match for all of these progress states
    const snapshot = tree;

    component = renderer.create(<StatusMessage progress={Data.Progress.GetConversation} />);
    tree = component.toJSON();
    expect(tree).toEqual(snapshot);

    component = renderer.create(<StatusMessage progress={Data.Progress.GetExcludedFolders} />);
    tree = component.toJSON();
    expect(tree).toEqual(snapshot);

    component = renderer.create(<StatusMessage progress={Data.Progress.GetFolderNames} />);
    tree = component.toJSON();
    expect(tree).toEqual(snapshot);
});

test("StatusMessage (Success)", () => {
    const component = renderer.create(<StatusMessage progress={Data.Progress.Success} />);
    const tree = component.toJSON();
    expect(tree).toMatchSnapshot();
});

test("StatusMessage (NotFound)", () => {
    const component = renderer.create(<StatusMessage progress={Data.Progress.NotFound} />);
    const tree = component.toJSON();
    expect(tree).toMatchSnapshot();
});
