/* eslint-disable prettier/prettier */
import * as React from "react";
import { Button, ButtonProps, Label } from "@fluentui/react-components";

// export default class ButtonExample extends React.Component {
//     public constructor(props) {
//         super(props);
//     }

//     insertText = async () => {
//         await Excel.run(async (context) => {
//             let sheet = context.workbook.worksheets.getActiveWorksheet();
//             let range = sheet.getRange("A1:B1");
//             range.value = "Hello World";
//             await context.sync();
//         })
//     }
//     render() {
//         let { disabled } = this.props;
//         return (
//           <div className="ms-BasicButtonExample">
//             <Label weight="semibold">Click the button to insert text.</Label>
//             <br />
//             <Button appearance="primary" disabled={disabled} size="large" onClick={this.insertText}>
//               Insert text
//             </Button>
//           </div>
//         );
//     }
// }

// eslint-disable-next-line @typescript-eslint/no-unused-vars
export default function ButtonExample(props) {
  const insertText = async () => {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getCell(2, 2);
      range.values = "Hello World";
      range.value = "Hello World";
      await context.sync();
    });
  };
  return (
    <div className="ms-BasicButtonExample">
      <Label weight="semibold">Click the button to insert text.</Label>
      <br />
      <Button appearance="primary" disabled={props.disabled} size="large" onClick={insertText}>
        Insert text
      </Button>
    </div>
  );
}
