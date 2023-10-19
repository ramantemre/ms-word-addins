import * as React from "react";
import { Button, ButtonProps, Label } from "@fluentui/react-components";

/* global Word */

export class ButtonExample extends React.Component<ButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertText = async () => {
    // Write text to the document when the button is selected.
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph("Hello Fluent UI React!", Word.InsertLocation.end);
      await context.sync();
    });
  };

  //Defines the Label and Button Fluent React UI components.
  public render() {
    let { disabled } = this.props;
    return (
      <section className="samples">
        <Label weight="semibold">Click the button to say Hello.</Label>
        <br />
        <Button appearance="primary" disabled={disabled} size="large" onClick={this.insertText}>
          Hello text
        </Button>
      </section>
    );
  }
}
