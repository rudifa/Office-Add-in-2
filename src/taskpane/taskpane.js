/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { base64Image } from "../../base64Image";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("insert-github-users-table").onclick = insertGithubUsersTable;
    document.getElementById("append-github-user-data").onclick = appendGithubUserData;
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    document.getElementById("replace-text").onclick = replaceText;
    document.getElementById("insert-image").onclick = insertImage;
    document.getElementById("insert-html").onclick = insertHTML;
    document.getElementById("insert-table").onclick = insertTable;
    document.getElementById("create-content-control").onclick = createContentControl;
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
  }
});

async function insertParagraph() {
  await Word.run(async (context) => {
    const docBody = context.document.body;
    docBody.insertParagraph(
      "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
      "Start"
    );
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function applyStyle() {
  await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function changeFont() {
  await Word.run(async (context) => {
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
    });
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertTextIntoRange() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    originalRange.load("text");
    await context.sync();
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertTextBeforeRange() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    originalRange.load("text");
    await context.sync();
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function replaceText() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertImage() {
  await Word.run(async (context) => {
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertHTML() {
  await Word.run(async (context) => {
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertTable() {
  await Word.run(async (context) => {
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    const tableData = [
      ["Name", "ID", "Birth City"],
      ["Bob", "434", "Chicago"],
      ["Sue", "719", "Havana"],
    ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function createContentControl() {
  await Word.run(async (context) => {
    // Queue commands to create a content control.
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function replaceContentInControl() {
  await Word.run(async (context) => {
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertGithubUsersTable() {
  await Word.run(async (context) => {
    const tableData = [
      ["login", "name", "location", "bio"],
      ["alison-mk", "Alison", "London", "I'm a developer"],
    ];
    const table = context.document.body.insertTable(tableData.length, tableData[0].length, "Start", tableData);
    table.headerRowCount = 1;
    table.styleBuiltIn = Word.Style.gridTable5Dark_Accent2;
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function appendGithubUserData() {
  // to the first table in the document
  await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();

    //console.log(`First table:`, firstTable);

    const tableData = [["rudifa", "Rudi", "Geneva", "I'm a programmer"]];
    firstTable.addRows("End", tableData.length, tableData);
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

/*
addRows(insertLocation:  "Start" | "End", rowCount: number, values?: string[][]): Word.TableRowCollection;
*/

/* https://xomino.com/category/office-add-in/
async function runAllTables() {
  await Word.run(async (context) => {
    const tableCollection = context.document.body.tables;
    // Queue a commmand to load the results.
    context.load(tableCollection);
    await context.sync();
    //cycle through the tbale collection and test the first cell of each table looking for insects
    for (var i = 0; i < tableCollection.items.length; i++) {
      var theTable = null;
      theTable = tableCollection.items[i];
      var cell1 = theTable.values[0][0];
      if (cell1 == "Insects") {
        //once found, load the table in memory and add a row
        context.load(theTable, "");
        await context.sync();
        let numRows = theTable.rowCount.toString();
        theTable.addRows("End", 1, [[numRows, "Lightning Bug"]]);
      }
    }
  });
}
*/

/*
$("#run").click(() => tryCatch(run));

async function run() {
    await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.font.color = "blue";

        await context.sync();
    });
}

// Default helper for invoking an action and handling errors. 
async function tryCatch(callback) {
  try {
      await callback();
  }
  catch (error) {
      OfficeHelpers.UI.notify(error);
      OfficeHelpers.Utilities.log(error);
  }
}
*/
