/// <reference path="../_testReferences.ts" />

import * as React from "react";
import * as renderer from "react-test-renderer";

import { Feedback } from "./Feedback";

test("Feedback (static content)", () => {
    const component = renderer.create(<Feedback />);
    const tree = component.toJSON();
    expect(tree).toMatchSnapshot();
});
