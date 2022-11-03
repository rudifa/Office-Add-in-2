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

    document.getElementById("find-github-users-table").onclick = () => tryCatch(findGithubUsersTable);
    document.getElementById("insert-github-users-table").onclick = () => tryCatch(insertGithubUsersTable);
    document.getElementById("append-github-user-data").onclick = () => tryCatch(appendGithubUserData);

    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    document.getElementById("apply-custom-style").onclick = () => tryCatch(applyCustomStyle);
    document.getElementById("change-font").onclick = () => tryCatch(changeFont);
    document.getElementById("insert-text-into-range").onclick = () => tryCatch(insertTextIntoRange);
    document.getElementById("insert-text-outside-range").onclick = () => tryCatch(insertTextBeforeRange);
    document.getElementById("replace-text").onclick = () => tryCatch(replaceText);
    document.getElementById("insert-image").onclick = () => tryCatch(insertImage);
    document.getElementById("insert-html").onclick = () => tryCatch(insertHTML);
    document.getElementById("insert-table").onclick = () => tryCatch(insertTable);
    document.getElementById("create-content-control").onclick = () => tryCatch(createContentControl);
    document.getElementById("replace-content-in-control").onclick = () => tryCatch(replaceContentInControl);
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
  });
}

async function applyStyle() {
  await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    await context.sync();
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    await context.sync();
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
  });
}

async function replaceText() {
  await Word.run(async (context) => {
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    await context.sync();
  });
}

async function insertImage() {
  await Word.run(async (context) => {
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    await context.sync();
  });
}

async function insertHTML() {
  await Word.run(async (context) => {
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    await context.sync();
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
  });
}

async function replaceContentInControl() {
  await Word.run(async (context) => {
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    await context.sync();
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
  });
}

async function appendGithubUserData() {
  const userName = getUserName();
  const url = `https://api.github.com/users/${userName}`;
  const obj = await fetchFrom(url);
  const userData = ["login", "name", "location", "bio"].map((key) => obj[key]);
  console.log(`appendGithubUserData`, userData);
  // append to the first table in the document
  await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const tableData = [userData];
    firstTable.addRows("End", tableData.length, tableData);
    await context.sync();
  });
}

// Helper function to get user name from input element
function getUserName() {
  const userName = document.getElementById("github-user-name").value;
  console.log(`getUserName`, userName);
  return userName;
}

async function findGithubUsersTable() {
  await Word.run(async (context) => {
    const tableCollection = context.document.body.tables;
    // Queue a commmand to load the results.
    context.load(tableCollection);
    await context.sync();
    //cycle through the table collection and test the first cell of each table looking for insects
    for (var i = 0; i < tableCollection.items.length; i++) {
      var theTable = null;
      theTable = tableCollection.items[i];
      var cell00 = theTable.values[0][0];
      if (cell00 == "login") {
        //once found, load the table in memory and add a row
        context.load(theTable, "");
        await context.sync();
        let numRows = theTable.rowCount.toString();
        // theTable.addRows("End", 1, [[numRows, "Lightning Bug"]]);
        displayInfoMessage("Found the table");
      }
    }
  });
}

/**
 * utilities
 */

function displayInfoMessage(message) {
 displayMessage(message, "blue");
}

function displayErrorMessage(message) {
  displayMessage(message, "red");
}

function displayMessage(message, color) {
  const paragraph = document.getElementById("p-message");
  paragraph.innerText = message;
  paragraph.style.color = color;
}

/**
 * Fetch data from a URL
 * @param {*} url
 * @returns promise that resolves to the JSON object returned by the url
 */
async function fetchFrom(url) {
  const response = await fetch(url);
  if (!response.ok) {
    //console.log(response);
    throw new Error(`status: ${response.status}`);
  }
  const data = await response.json();
  return data;
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  displayMessage("");
  try {
    await callback();
  } catch (error) {
    console.log(error);
    displayErrorMessage(error);

    if (error instanceof OfficeExtension.Error) {
      const debugInfo = JSON.stringify(error.debugInfo)
      console.log("Debug info: " + debugInfo);
    }
    // OfficeHelpers.UI.notify(error);
    // OfficeHelpers.Utilities.log(error);
  }
}

// TODO add  input element to get the user name

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
*/
