import * as React from "react";
import { Label, Button, ButtonProps } from "@fluentui/react-components";

/* global Word */

export class TableExample extends React.Component<ButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertTable = async () => {
    await Word.run(async (context) => {
      // Use a two-dimensional array to hold the initial table values.
      const data = [
        ["Tokyo", "Beijing", "Seattle"],
        ["Apple", "Orange", "Pineapple"],
      ];
      const table = context.document.body.insertTable(2, 3, "Start", data);
      table.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent2;
      table.styleFirstColumn = false;

      await context.sync();
    });
  };

  //Defines the Label and Table Fluent React UI components.
  public render() {
    let { disabled } = this.props;
    return (
      <section className="samples">
        <Label weight="semibold">Click the Button to Add Table.</Label>
        <br />
        <Button appearance="primary" disabled={disabled} size="large" onClick={this.insertTable}>
          Add Table
        </Button>
      </section>
    );
  }
}
