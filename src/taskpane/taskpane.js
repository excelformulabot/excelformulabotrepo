/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      var iframe = document.getElementById("myFrame");
      console.log(iframe);
      var outputTextArea = iframe.contentWindow.document.getElementByClassName("bubble-element MultiLineInput")[0];

      if (outputTextArea != undefined) {
        range.values = outputTextArea.value;
        console.log("outputTextArea.value: " + outputTextArea.value);
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
