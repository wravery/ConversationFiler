import * as React from "react";

export interface HelloProps {
    compiler: string;
    framework: string;
}

export class Hello extends React.Component<HelloProps, undefined> {
    render() {
        return <h1>Greetings! Hello from {this.props.compiler} and {this.props.framework}!</h1>;
    }
}
