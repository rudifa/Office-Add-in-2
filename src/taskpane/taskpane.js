/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, Word */

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

    document.getElementById("insert-github-users-table").onclick = () => tryCatch(findOrInsertGithubUsersTable);
    document.getElementById("update-github-user-data").onclick = () => tryCatch(updateGithubUserData);

    document.getElementById("test-async").onclick = () => tryCatch(testAsync);

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

async function insertGithubUsersTable() {
  await Word.run(async (context) => {
    const tableData = [["login", "name", "location", "bio"]];
    const table = context.document.body.insertTable(tableData.length, tableData[0].length, "Start", tableData);
    table.headerRowCount = 1;
    table.styleBuiltIn = Word.Style.gridTable5Dark_Accent2;
    await context.sync();
  });
}

async function _addRow(context, firstTable, userData) {
  const tableData = [userData];
  firstTable.addRows("End", tableData.length, tableData);
  await context.sync();
}

async function _fetchUserData() {
  const userName = getUserName();
  // fetch the user's data from the GitHub API.
  const url = `https://api.github.com/users/${userName}`;
  const obj = await fetchFrom(url);
  // prepare the data for the table.
  const userData = ["login", "name", "location", "bio"].map((key) => obj[key] || "");
  console.log(`updateGithubUserData`, userData);
  return userData;
}

// Helper function to get user name from input element
function getUserName() {
  const userName = document.getElementById("github-user-name").value;
  console.log(`getUserName`, userName);
  return userName;
}

async function findOrInsertGithubUsersTable() {
  await Word.run(async (context) => {
    const tableFound = await findGithubUsersTable();
    console.log(`findOrInsertGithubUsersTable`, tableFound);
    if (!tableFound) {
      await insertGithubUsersTable();
    }
  });
}

async function findGithubUsersTable() {
  var tableFound = false;
  await Word.run(async (context) => {
    // find a table with the tag "githubUsers"
    const tableCollection = context.document.body.tables;
    // Queue a commmand to load the results.
    context.load(tableCollection);
    await context.sync();
    //cycle through the table collection and test the first cell of each table looking for insects
    for (var i = 0; i < tableCollection.items.length; i++) {
      const table = tableCollection.items[i];
      var cell00 = table.values[0][0];
      if (cell00 == "login") {
        tableFound = true;
        break;
      }
    }
    console.log(`findGithubUsersTable`, tableFound);
  });
  return tableFound;
}

async function updateGithubUserData() {
  await Word.run(async (context) => {
    var tableFound = await _findGithubUsersTable(context);
    console.log(`testAsync`, tableFound);
    if (!tableFound) {
      await insertGithubUsersTable();
      tableFound = await _findGithubUsersTable(context);
    }
    context.load(tableFound);
    await context.sync();
    console.log(`testAsync loaded tableFound`, tableFound);

    // await _logAllCellValues(context, tableFound);

    const userData = await _fetchUserData();
    const userName = userData[0];
    const rowFound = await _findMatchingRow(context, tableFound, userName);
    console.log(`testAsync rowFound`, rowFound);

    if (!rowFound) {
      await _addRow(context, tableFound, userData);
    } else {
      await _updateRow(context, rowFound, userData);
    }
  });
}

async function _updateRow(context, rowFound, userData) {
  console.log(`_updateRow`, rowFound, userData);
  rowFound.values = [userData];
  await context.sync();
}

async function _findMatchingRow(context, tableFound, cell0Value) {
  const tableRows = tableFound.rows;
  context.load(tableRows);
  await context.sync();
  console.log(`_findMatchingRow loaded tableRows`, tableRows);
  let rowFound = null;
  for (var i = 0; i < tableRows.items.length; i++) {
    const tableRow = tableRows.items[i];
    console.log(`_findMatchingRow tableRow 1`, tableRow);
    const tableRowCells = tableRow.cells;
    context.load(tableRow);
    context.load(tableRowCells);

    await context.sync();
    console.log(`_findMatchingRow tableRow 2`, tableRow);

    console.log(`_findMatchingRow loaded tableRowCells ${i}`, tableRowCells);
    const tableRowCell0 = tableRowCells.items[0];
    console.log(`_findMatchingRow loaded tableRowCell ${i}`, tableRowCell0.value);
    if (tableRowCell0.value === cell0Value) {
      console.log(`_findMatchingRow found matching row`, tableRow);
      rowFound = tableRow;
      console.log(`_findMatchingRow rowFound`, rowFound);
    }
  }
  console.log(`_findMatchingRow rowFound`, rowFound);
  return rowFound;
}

async function _findGithubUsersTable(context) {
  const tableCollection = context.document.body.tables;
  // Queue a commmand to load the results.
  context.load(tableCollection);
  await context.sync();
  var tableFound = null;
  //cycle through the table collection and test the first cell of each table looking for insects
  for (var i = 0; i < tableCollection.items.length; i++) {
    const table = tableCollection.items[i];
    var cell00 = table.values[0][0];
    if (cell00 == "login") {
      tableFound = table;
      break;
    }
  }
  return tableFound;
}

/**
 * Leftovers from earlier code versions
 */

async function testAsync() {
  await Word.run(async (context) => {
    var tableFound = await _findGithubUsersTable(context);
    console.log(`testAsync`, tableFound);
    if (!tableFound) {
      await insertGithubUsersTable();
      tableFound = await _findGithubUsersTable(context);
    }
    context.load(tableFound);
    await context.sync();
    console.log(`testAsync loaded`, tableFound);
  });
}

async function _logAllCellValues(context, tableFound) {
  const tableRows = tableFound.rows;
  context.load(tableRows);
  await context.sync();
  console.log(`_logAllCellValues loaded tableRows`, tableRows);
  for (var i = 0; i < tableRows.items.length; i++) {
    const tableRow = tableRows.items[i];
    const tableRowCells = tableRow.cells;
    context.load(tableRowCells);
    await context.sync();
    console.log(`_logAllCellValues loaded tableRowCells ${i}`, tableRowCells);
    for (var j = 0; j < tableRowCells.items.length; j++) {
      const tableRowCell = tableRowCells.items[j];
      //const tableRowCellValue = tableRowCell.values[0][0];
      console.log(`_logAllCellValues loaded tableRowCell ${i},${j}`, tableRowCell.value);
    }
  }
}

async function updateGithubUserData_0() {
  // get the user's name from the user input.
  const userData = await _fetchUserData();
  // get the user data from the table (if present)

  // append to the first table in the document
  // TODO find the table
  await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    await _addRow(context, firstTable, userData);
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
      const debugInfo = JSON.stringify(error.debugInfo);
      console.log("Debug info: " + debugInfo);
    }
    // OfficeHelpers.UI.notify(error);
    // OfficeHelpers.Utilities.log(error);
  }
}

/**
 * Functions from the tutorial
 */

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
