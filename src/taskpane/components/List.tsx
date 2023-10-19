import * as React from "react";
import { Label, Button, ButtonProps } from "@fluentui/react-components";

/* global Word */

export class ListExample extends React.Component<ButtonProps, {}> {
  public constructor(props) {
    super(props);
  }

  insertList = async () => {
    // This example starts a new list with the second paragraph.
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("$none");

      await context.sync();

      // Start new list using the second paragraph.
      const list = paragraphs.items[1].startNewList();
      list.load("$none");

      await context.sync();

      // To add new items to the list, use Start or End on the insertLocation parameter.
      list.insertParagraph("New list item at the start of the list", "Start");
      const paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

      // Set up list level for the list item.
      paragraph.listItem.level = 4;

      // To add paragraphs outside the list, use Before or After.
      list.insertParagraph("New paragraph goes after (not part of the list)", "After");

      await context.sync();
    });
  };

  setup = async () => {
    await Word.run(async (context) => {
      const body = context.document.body;

      //   body.clear();

      body.insertParagraph(
        "Themes and styles also help keep your document coordinated. When you click design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme. ",
        "Start"
      );
      body.insertParagraph(
        "Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign. ",
        "Start"
      );
      body.insertParagraph(
        "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
        "Start"
      );
      body.paragraphs
        .getLast()
        .insertText(
          "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries. ",
          "Replace"
        );
    });
  };

  public render() {
    let { disabled } = this.props;
    return (
      <>
        <section className="samples">
          <Label weight="semibold">Click the Button to Add Paragraph.</Label>
          <br />
          <Button appearance="primary" disabled={disabled} size="large" onClick={this.setup}>
            Add Paragraph
          </Button>
        </section>
        <section className="samples">
          <Label weight="semibold">Click the Button to Insert List.</Label>
          <br />
          <Button appearance="primary" disabled={disabled} size="large" onClick={this.insertList}>
            Insert List
          </Button>
        </section>
      </>
    );
  }
}
