/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;


    document.getElementById("protect").onclick = protect;

    document.getElementById("unprotect").onclick = unprotect;

    document.getElementById("setup").onclick = setup;

  }
});

/*
export async function run() {
  try {
    await Excel.run(async context => {
      
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

*/
/** Protecting the work sheets*/
export async function protect() {
  try {
    let password = "password";
    await Excel.run(async context => {
      var unprotectSheets = context.workbook.worksheets;
      unprotectSheets.load("items");
      await context.sync();

      for (var i = 0; i < unprotectSheets.items.length; i++) {
        unprotectSheets.load("protection/protected");
        await context.sync();
        unprotectSheets.items[i].protection.protect(null, password);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

/** unProtecting the work sheets*/
export async function unprotect() {
  try {
    let password = "password";
    await Excel.run(async context => {
      var protectSheets = context.workbook.worksheets;
      protectSheets.load("items");
      await context.sync();

      for (var i = 0; i < protectSheets.items.length; i++) {
        protectSheets.load("protection/protected");
        await context.sync();
        protectSheets.items[i].protection.unprotect(password);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

/** calling the api ans adding the data to the work sheet */
export async function setup() {
  await Excel.run(async context => {
    context.workbook.worksheets.getItemOrNullObject("Authorization Cover").delete();
    const sheet = context.workbook.worksheets.add("Authorization Cover");

    const expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    const response = await fetch('https://deckofcardsapi.com/api/deck/new/draw/?count=35');
   const myJson = await response.json();

    expensesTable.getHeaderRowRange().values = [["code","image","suit","value"]];

    var transactions = myJson["cards"];

    var newData = transactions.map(item => 
        [item.code, item.image, item.suit, item.value]);

    expensesTable.rows.add(null, newData);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    sheet.activate();
    await context.sync();
  });
}


