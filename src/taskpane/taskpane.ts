/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      const precedents = range.getPrecedents();

      try {
        precedents.ranges.load({
          formulas: true,
          address: true,
        });

        await context.sync();
      } catch (error) {
        // ignore
      }

      console.log(`Addresses of precedents: ${precedents.ranges.items.map((x) => x.address).join(", ")}.`);
      console.log(`Formulas of precedents: ${precedents.ranges.items.map((x) => x.formulas).join(", ")}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
