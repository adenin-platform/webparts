import * as React from "react";

export default class ErrorBoundary extends React.Component<any, { hasError: boolean; }> {
    constructor(props) {
        super(props);
        this.state = { hasError: false };
    }

    public componentDidCatch(error: any, errorInfo: any) {
        this.setState({
            hasError: true,
        });
        console.warn(error);
    }

    public render() {
        if (this.state.hasError) {
            return (
                <div style={{
                    backgroundColor: "#f4f4f4",
                    padding: "10px",
                    fontWeight: 100,
                    textAlign: "center"
                }}>
                    <h1>Sorry, something went wrong</h1>
                    <p style={{fontSize: "17px"}}>Please refresh the page to reload this webpart. If the issue persists, please contact your system administrator.</p>
                </div>
            );
        }

        return this.props.children;
    }
}