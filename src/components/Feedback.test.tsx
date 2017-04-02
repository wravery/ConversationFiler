/// <reference path="../../node_modules/@types/jest/index.d.ts" />
/// <reference path="../../node_modules/@types/react-test-renderer/index.d.ts" />

import * as React from "react";
import * as renderer from "react-test-renderer";

import { Feedback } from "./Feedback";

test("Feedback (static content)", () => {
    const component = renderer.create(<Feedback />);
    const tree = component.toJSON();
    expect(tree).toMatchSnapshot();
});
