/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, Word */

require("!style-loader!css-loader!./taskpane.css");

import { base64Image } from "../../base64Image";

import * as GHConnect from "./github-connect";

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

    document.getElementById("display-addin-info").onclick = () => tryCatch(displayAddinInfo);
    document.getElementById("p-message").onclick = () => tryCatch(clearAddinInfo);

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

/**
 * Interactions with the UI
 */

function getUserName() {
  const userName = document.getElementById("github-user-name").value;
  console.log(`getUserName`, userName);
  return userName;
}

function displayAddinInfo() {
  const message = `
  This sample Word add-in shows how to insert a table into a document, and to update the table with data fetched from the GitHub API.`;
  displayInfoMessage(message);
}

function clearAddinInfo() {
  displayInfoMessage("");
}

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
 * Interactions with the document
 */

async function insertGithubUsersTable() {
  await Word.run(async (context) => {
    const tableHeaderData = [GHConnect.userDataKeys];
    const table = context.document.body.insertTable(
      tableHeaderData.length,
      tableHeaderData[0].length,
      "Start",
      tableHeaderData
    );
    table.headerRowCount = 1;
    table.styleBuiltIn = Word.Style.gridTable5Dark_Accent2;
    await context.sync();
  });
}

async function findOrInsertGithubUsersTable() {
  await Word.run(async (context) => {
    const tableFound = await _findGithubUsersTable(context);
    if (tableFound) {
      displayInfoMessage("Table already exists");
    } else {
      await insertGithubUsersTable();
    }
  });
}

async function updateGithubUserData() {
  await Word.run(async (context) => {
    let tableFound = await _findGithubUsersTable(context);
    if (!tableFound) {
      await insertGithubUsersTable();
      tableFound = await _findGithubUsersTable(context);
    }
    context.load(tableFound);
    await context.sync();

    const userName = getUserName();
    const userData = await GHConnect.fetchUserData(userName);
    const rowFound = await _findMatchingRow(context, tableFound, userName);

    if (!rowFound) {
      await _addRow(context, tableFound, userData);
    } else {
      await _updateRow(context, rowFound, userData);
    }
  });
}

async function _addRow(context, table, userData) {
  const tableData = [userData];
  table.addRows("End", tableData.length, tableData);
  await context.sync();
}

async function _updateRow(context, rowFound, userData) {
  console.log(`_updateRow`, rowFound, userData);
  rowFound.values = [userData];
  await context.sync();
}

async function _findMatchingRow(context, table, cell0Value) {
  const tableRows = table.rows;
  context.load(tableRows);
  await context.sync();
  let rowFound = null;
  for (const tableRow of tableRows.items) {
    const tableRowCells = tableRow.cells;
    context.load(tableRow);
    context.load(tableRowCells);
    // eslint-disable-next-line office-addins/no-context-sync-in-loop
    await context.sync();

    const tableRowCell0 = tableRowCells.items[0];
    if (tableRowCell0.value === cell0Value) {
      rowFound = tableRow;
    }
  }
  return rowFound;
}

async function _findGithubUsersTable(context) {
  const tableCollection = context.document.body.tables;
  context.load(tableCollection);
  await context.sync();
  let tableFound = null;
  // cycle through the table collection and look for a matching cell[0][0] value
  for (const table of tableCollection.items) {
    if (table.values[0][0] == GHConnect.userDataKeys[0]) {
      tableFound = table;
      break;
    }
  }
  return tableFound;
}

/**
 * utilities
 */

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  displayMessage("");
  try {
    await callback();
  } catch (error) {
    console.log(error);
    displayErrorMessage(error);
  }
}

/**
 * Testing functions
 */

async function testAsync() {
  await Word.run(async (context) => {
    const table = await _findGithubUsersTable(context);
    console.log(`testAsync table`, table);
    console.log(`testAsync table.values`, table.values);
    console.log(`testAsync table.values[0]`, table.values[0]);
    console.log(`testAsync table.values[0][0]`, table.values[0][0]);
  });
}

async function testAsync_1() {
  const userData = await GHConnect.fetchUserData(getUserName());
  console.log(`testAsync`, userData);
  displayInfoMessage(`testAsync userData: ${userData}`);
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
