/***
-Global Variables. These variables are used through the whole script over and over again. 
***/
// Get current book
const ss = SpreadsheetApp.getActive();
//Get the page "Items"
const items = ss.getSheetByName("Items");
//Get the page "Cosignment Input"
const cosignInput = ss.getSheetByName("Cosignment Input");
//Get the page "Cosignments"
const cosignments = ss.getSheetByName("Cosignments");
//Get the ygncredit
const ygnCredits = ss.getSheetByName("YGN Credits");
//Get the pmncredit
const credits = ss.getSheetByName("credits");
//Get minQty
const minQty = ss.getSheetByName("Min Qty");
//Get Cosignment Costing
const cosCosting = ss.getSheetByName("Cosignment Costing");
//Get Cosignment Check
const cosCheck = ss.getSheetByName("Cosignment Check");
//Get creditDue
const creditDue = ss.getSheetByName("Credit Due");
//Stock movement
//const movementData = ss.getSheetByName("Movement Data");
//Search Transition
//const stockMovement = ss.getSheetByName("Stock Movement");
//Get the items last row
let itemLastRow = items.getLastRow();
//Get the cosignment last row
let nextCosRow = cosignments.getLastRow() + 1;
//Sinomm Inventory Report
const movementData = SpreadsheetApp.openById(
  "1bU2qh_ON1noYu67vt9hu2_OE6roag75uAeiZYhepWhg"
).getSheetByName("Movement Data");
//Get sinomm item
const sinommItem = SpreadsheetApp.openById(
  "1S-Xjj4Gd31NgBwd9MX-tpwC0UUkZWzNpADYRJNjiwHI"
).getSheetByName("Items");
//Get main accounting
const accounts = SpreadsheetApp.openById(
  "1Zy8t1B4hO-kxFVQA_yabrUbdeZgUjTPZr8Q5MMOjW5U"
).getSheetByName("Charts Of Accounts");
//get ui
let ui = SpreadsheetApp.getUi();

/***********************************
- Listen for specific trigger spots to launch certain functions
************************************/
let onEditSinommAccountingTriggers = (e) => {
  return e.source.getActiveSheet().getName() == "Cosignment Input" &&
    e.range.getRow() == 10 &&
    e.range.getColumn() == 3
    ? onEditNewCosSinommAccounting(e)
    : e.source.getActiveSheet().getName() == "Items" && e.range.getColumn() == 7
    ? onEditItems(e)
    : e.source.getActiveSheet().getName() == "Stock Movement" &&
      e.range.getRow() == 2 &&
      e.range.getColumn() == 1
    ? searchMovement(e)
    : e.source.getActiveSheet().getName() == "Cosignment Check" &&
      e.range.getRow() == 4 &&
      e.range.getColumn() == 9
    ? getCosting(e)
    : e.source.getActiveSheet().getName() == "Cosignment Check" &&
      e.range.getRow() == 57 &&
      e.range.getColumn() == 3
    ? submitCosting(e)
    : e.source.getActiveSheet().getName() == "Cosignment Check" &&
      e.range.getRow() == 6 &&
      e.range.getColumn() == 9
    ? addCosting(e)
    : false;
};

let onEditNewCosSinommAccounting = (e) => {
  let signature = cosignInput.getRange(10, 3).getValue();
  e.range.setValue("Waiting...");
  newCos(signature);
  e.range.setValue("Done!");
};

//Update cosignment, update the cosignment that have 0 quantity with the next cosignment
let newCos = (signature) => {
  let date = cosignInput.getRange(10, 2).getValue();
  let cosignmentNum = cosignInput.getRange(10, 5).getValue();
  let cartArrayTwo = cosignInput.getRange(14, 2, 20, 8).getValues();
  let shares = cosignInput.getRange(4, 6, 4, 1).getValues();
  let share = "";
  let cosArray = [];
  let stockMovementArray = [];
  let j = 0;
  if (cosignmentNum == "" || signature == undefined) {
    return false;
  }
  for (i = 0; i < shares.length; i++) {
    if (shares[i] == "") {
      share = share.substr(0, share.length - 1);
      break;
    }
    share += shares[i] + "+";
  }
  for (var i = 0; i < 20; i++) {
    if (cartArrayTwo[i][0].valueOf() == "") {
      break;
    }
    let code = cartArrayTwo[i][0];
    let qty = cartArrayTwo[i][2];
    let cost = cartArrayTwo[i][3];
    let rate = cartArrayTwo[i][5];
    let locationToDrop = cartArrayTwo[i][7];
    let lastCoRow = cosignments.getLastRow() + 1;
    let itemRow = locationToDrop.split("-")[1];
    locationToDrop = locationToDrop.split("-")[0];
    if (locationToDrop == "J6") {
      itemCol = 7;
    }
    if (locationToDrop == "TGN") {
      itemCol = 8;
    }
    if (locationToDrop == "DAGON") {
      itemCol = 9;
    }

    cosArray[j] = [];
    cosArray[j] = [
      ...cosArray[j],
      cosignmentNum + "," + code,
      cosignmentNum,
      code,
      date,
      qty,
      "=SUMIFS('Sinomm Receipt'!H:H,'Sinomm Receipt'!A:A,A" +
        Number(lastCoRow + j) +
        ",'Sinomm Receipt'!T:T,\"Valid\")",
      "=E" + (lastCoRow + j) + "-F" + (lastCoRow + j),
      cost,
      signature,
      "=if(G" + (lastCoRow + j) + "=0,FALSE,TRUE)",
      share,
      locationToDrop,
      rate,
      "=E" + (lastCoRow + j) + "*H" + (lastCoRow + j) + "",
      "=CONCAT(CONCAT(B" +
        Number(lastCoRow + j) +
        ',"|"),G' +
        Number(lastCoRow + j) +
        ")",
      "=H" +
        (lastCoRow + j) +
        "*G" +
        (lastCoRow + j) +
        "/(COUNT(SPLIT(K" +
        (lastCoRow + j) +
        ',"+")))',
      "=H" + (lastCoRow + j) + "*G" + (lastCoRow + j) + "",
    ];

    let initialBalance = sinommItem.getRange(itemRow, itemCol).getValue();
    let currentBalance = initialBalance + qty;

    stockMovementArray[j] = [];
    stockMovementArray[j] = [
      ...stockMovementArray[j],
      "Cosignment",
      code,
      cosignmentNum,
      date,
      locationToDrop,
      initialBalance,
      qty,
      ,
      currentBalance,
    ];

    j++;
    lastCoRow++;
    let initialValue = sinommItem.getRange(itemRow, itemCol).getValue();
    sinommItem.getRange(itemRow, itemCol).setValue(initialValue + qty);
  }

  cosignments.getRange(nextCosRow, 1, cosArray.length, 17).setValues(cosArray);
  movementData
    .getRange(movementData.getLastRow() + 1, 1, stockMovementArray.length, 9)
    .setValues(stockMovementArray);

  cosignInput.getRange(10, 3).setValue("");
  cosignInput.getRange(10, 5).setValue("");
  cosignInput.getRange(7, 5).setValue("");
  cosignInput.getRange(14, 3, 20, 3).setValue("");
  cosignInput.getRange(4, 5, 4, 1).setValue("");
  cosignInput.getRange(14, 8, 20, 1).setValue("");
};

let sendStockEmail = () => {
  let x = minQty.getRange(1, 5).getValue();
  const minQtyItems = minQty.getRange(2, 1, x, 4).getDisplayValues();

  const today = new Date();

  const stockTemplate = HtmlService.createTemplateFromFile("StockTemplate");
  stockTemplate.today = today;
  stockTemplate.minQtyItems = minQtyItems;

  const htmlForEmail = stockTemplate.evaluate().getContent();

  GmailApp.sendEmail(
    "thinthiri@utotools.com,eitun@utotools.com",
    "Low Stocks",
    "Please open this email with a client that supports HTML",
    { htmlBody: htmlForEmail }
  );
};

let sendCreditEmail = (department) => {
  let x =
    department == "ygn"
      ? creditDue.getRange(1, 4).getValue()
      : creditDue.getRange(1, 8).getValue();

  const creditItems =
    department == "ygn"
      ? creditDue.getRange(2, 1, x, 4).getDisplayValues()
      : creditDue.getRange(2, 5, x, 4).getDisplayValues();
  const today = new Date();

  const creditTemplate = HtmlService.createTemplateFromFile("CreditTemplate");
  creditTemplate.today = today;
  creditTemplate.creditItems = creditItems;

  const htmlForEmail = creditTemplate.evaluate().getContent();
  if (department == "ygn") {
    GmailApp.sendEmail(
      "accounting@utotools.com,suhtwe.hs@gmail.com",
      "Credit Vouchers Left",
      "Please open this email with a client that supports HTML",
      { htmlBody: htmlForEmail }
    );
  } else {
    GmailApp.sendEmail(
      "accounting@utotools.com",
      "Credit Vouchers Left",
      "Please open this email with a client that supports HTML",
      { htmlBody: htmlForEmail }
    );
  }
};

let sendCredit = () => {
  sendCreditEmail("ygn");
};

//Update creditLists at the end of every month
let creditTimeInterval = () => {
  let currentMonth = new Date().getMonth() + 1;
  let ygnRange = credits.getRange(1, 1, 400, 2).getValues();
  ygnCredits.getRange(2, currentMonth * 2 - 1, 400, 2).setValues(ygnRange);
};

/***********************************
-specific trigger and Call the addItem Function
************************************/

let onEditItems = (e) => {
  let itemData = items.getRange(e.range.getRow(), 1, 1, 6).getValues();
  let codeArray = items.getRange("A1:A").getValues();
  let codeArrayLastRow = codeArray.filter(String).length;
  codeArray = [].concat(...codeArray);
  let changedData = [];
  let j = 0;
  let response;

  for (data of itemData[0]) {
    let responseText = "";

    if (j == 0) {
      if (data != "") {
        responseText = data;
      } else {
        response = ui.prompt("Editing Code: " + data, ui.ButtonSet.YES_NO);
        if (response.getSelectedButton() == ui.Button.YES) {
          responseText = response.getResponseText();
          if (codeArray.includes(responseText)) {
            ui.alert("Duplicate Code. " + responseText + " already exists!");
            break;
          }
        } else {
          responseText = data;
        }
      }
    }

    if (j == 1) {
      response = ui.prompt("Editing Description: " + data, ui.ButtonSet.YES_NO);
      if (response.getSelectedButton() == ui.Button.YES) {
        responseText = response.getResponseText();
      } else {
        responseText = data;
      }
    }
    if (j == 2) {
      response = ui.prompt("Editing Sale Price: " + data, ui.ButtonSet.YES_NO);
      if (response.getSelectedButton() == ui.Button.YES) {
        responseText = response.getResponseText();
      } else {
        responseText = data;
      }
    }
    if (j == 3) {
      response = ui.prompt("Editing Cost Price: " + data, ui.ButtonSet.YES_NO);
      if (response.getSelectedButton() == ui.Button.YES) {
        responseText = response.getResponseText();
      } else {
        responseText = data;
      }
    }
    if (j == 4) {
      response = ui.prompt("Editing Unit: " + data, ui.ButtonSet.YES_NO);
      if (response.getSelectedButton() == ui.Button.YES) {
        responseText = response.getResponseText();
      } else {
        responseText = data;
      }
    }
    if (j == 5) {
      response = ui.prompt("Editing Packaging: " + data, ui.ButtonSet.YES_NO);
      if (response.getSelectedButton() == ui.Button.YES) {
        responseText = response.getResponseText();
      } else {
        responseText = data;
      }
    }

    changedData.push(responseText);

    j++;
  }
  changedData = [changedData];
  items.getRange(e.range.getRow(), 1, 1, 6).setValues(changedData);
  e.range.setValue(false);
};

/***********************************
- Create Movement Object Constructor
************************************/
function Movement(
  type,
  code,
  invoiceNum,
  date,
  location,
  initialBalance,
  qtyIn,
  qtyOut,
  balance
) {
  this.type = type;
  this.code = code;
  this.invoiceNum = invoiceNum;
  this.date = date;
  this.location = location;
  this.initialBalance = initialBalance;
  this.qtyIn = qtyIn;
  this.qtyOut = qtyOut;
  this.balance = balance;
}

/***********************************
-Search the stock movement data filter
************************************/

let searchMovement = (e) => {
  let filterArray = stockMovement.getRange(2, 3, 1, 5).getValues();
  let movementItems = movementData
    .getRange(2, 1, movementData.getLastRow(), 9)
    .getValues();
  let movementObjectsArray = [];

  // Type Filter
  let typeFilter = (movement) => {
    if (filterArray[0][2] != "") {
      return movement.type == filterArray[0][2];
    } else {
      return movement;
    }
  };

  // Description Filter
  let codeFilter = (movement) => {
    if (filterArray[0][0] != "") {
      return movement.code == filterArray[0][0];
    } else {
      return movement;
    }
  };

  // Invoice Filter
  let locationFilter = (movement) => {
    if (filterArray[0][1] != "") {
      return movement.location == filterArray[0][1];
    } else {
      return movement;
    }
  };

  // Date Start Filter
  let dateStartFilter = (movement) => {
    if (filterArray[0][3] != "") {
      return movement.date >= filterArray[0][3];
    } else {
      return movement;
    }
  };

  // Date End Filter
  let dateEndFilter = (movement) => {
    if (filterArray[0][4] != "") {
      return movement.date >= filterArray[0][4];
    } else {
      return movement;
    }
  };

  for (data of movementItems) {
    let type = data[0];
    let code = data[1];
    let invoiceNum = data[2];
    let date = data[3];
    let location = data[4];
    let initialBalance = data[5];
    let qtyIn = data[6];
    let qtyOut = data[7];
    let balance = data[8];

    let movementObject = new Movement(
      type,
      code,
      invoiceNum,
      date,
      location,
      initialBalance,
      qtyIn,
      qtyOut,
      balance
    );
    movementObjectsArray.push(movementObject);
  }

  let filteredMovement = movementObjectsArray
    .filter(typeFilter)
    .filter(codeFilter)
    .filter(locationFilter)
    .filter(dateStartFilter)
    .filter(dateEndFilter);

  let outputArray = [];

  for (data of filteredMovement) {
    let transaction = [
      data.type,
      data.code,
      data.invoiceNum,
      data.date,
      data.location,
      data.initialBalance,
      data.qtyIn,
      data.qtyOut,
      data.balance,
    ];
    outputArray.push(transaction);
  }
  stockMovement.getRange(5, 2, stockMovement.getLastRow(), 9).setValue("");
  stockMovement.getRange(5, 2, outputArray.length, 9).setValues(outputArray);
  e.range.setValue(false);
};

/***********************************
-Get Cosignment Costing
************************************/
let getCosting = (e) => {
  let cosignment = e.value;
  let cosRow = cosCheck.getRange(4, 11).getValue();

  if (cosRow == "") {
    ui.alert("Cosignment" + cosignment + " does not exist!");
    cosCheck.getRange(4, 3).setValue("");
    cosCheck.getRange(6, 3).setValue("");
    cosCheck.getRange(8, 3).setValue("");
    cosCheck.getRange(4, 13).setValue("");
    cosCheck.getRange(6, 13).setValue("");
    cosCheck.getRange(8, 13).setValue("");
    cosCheck.getRange(11, 2, 45, 10).setValue("");
    return false;
  }

  let costingRecordArray = cosCosting.getRange(cosRow, 1, 45, 14).getValues();
  let from = costingRecordArray[0][1];
  let invoice = costingRecordArray[0][2];
  let by = costingRecordArray[0][3];
  let invoiceDate = costingRecordArray[0][4];
  let receiveDate = costingRecordArray[0][5];
  let goodReceive = costingRecordArray[0][6];
  let costArray = [];
  let j = 0;

  cosCheck.getRange(4, 3).setValue(from);
  cosCheck.getRange(6, 3).setValue(invoice);
  cosCheck.getRange(8, 3).setValue(by);
  cosCheck.getRange(4, 13).setValue(invoiceDate);
  cosCheck.getRange(6, 13).setValue(receiveDate);
  cosCheck.getRange(8, 13).setValue(goodReceive);

  for (data of costingRecordArray) {
    if (data[0].valueOf() != cosignment) {
      break;
    }
    let description = data[7];
    let payment = data[8];
    let date = data[9];
    let memo = data[10];
    let bankCharge = data[11];
    let rate = data[12];
    let kyats = data[13];

    costArray[j] = [];
    costArray[j] = [
      ...costArray[j],
      description,
      ,
      ,
      ,
      payment,
      date,
      memo,
      bankCharge,
      rate,
      kyats,
    ];

    j++;
  }
  cosCheck.getRange(11, 2, 45, 10).setValue("");
  cosCheck.getRange(11, 2, costArray.length, 10).setValues(costArray);
};

/***********************************
-Submit Costing
************************************/
let submitCosting = (e) => {
  let invoice = cosCheck.getRange(4, 9).getValue();
  let totalCost = cosCheck.getRange(8, 11).getValue();
  let allAccounts = accounts
    .getRange(1, 6, accounts.getLastRow(), 1)
    .getValues();
  let sinommContainerCostingAccountRow;
  for (let i = 0; i < allAccounts.length; i++) {
    if (allAccounts[i][0] == "92002") {
      sinommContainerCostingAccountRow = i;
    }
  }
  let initialValue = accounts
    .getRange(sinommContainerCostingAccountRow + 1, 7)
    .getValue();
  let newValue = initialValue + totalCost;
  let cosLastRow = Number(cosCheck.getRange(57, 4).getDisplayValue()) + 1;

  let itemsInInvoice = cosCheck
    .getRange(60, 2, cosLastRow - 60, 11)
    .getValues();
  let cosignmentArray = cosignments
    .getRange(1, 1, cosignments.getLastRow() + 1, 1)
    .getValues();

  let response = ui.alert(
    "Submit cost for invoice: " + invoice + " : " + totalCost + " ?",
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.YES) {
    for (item of itemsInInvoice) {
      let cosignment = invoice + "," + item[0];
      let cost = item[10];
      let j = 1;

      for (data of cosignmentArray) {
        if (cosignment == data[0]) {
          cosignments.getRange(j, 8).setValue(cost);
        }
        j++;
      }
    }

    accounts
      .getRange(sinommContainerCostingAccountRow + 1, 7)
      .setValue(newValue);
    e.range.setValue(false);
  } else {
    e.range.setValue(false);
    return false;
  }
};

/***********************************
-Add Costing
************************************/
let addCosting = (e) => {
  e.range.setValue("Waiting...");

  let invoice = cosCheck.getRange(4, 9).getValue();
  let from = cosCheck.getRange(4, 3).getValue();
  let invoiceNum = cosCheck.getRange(6, 3).getValue();
  let by = cosCheck.getRange(8, 3).getValue();
  let invoiceDate = cosCheck.getRange(4, 13).getValue();
  let receivedDate = cosCheck.getRange(6, 13).getValue();
  let goodReceiveNum = cosCheck.getRange(8, 13).getValue();
  let signature = e.value;
  let costArray = cosCheck.getRange(11, 2, 45, 10).getValues();
  let cosignmentRow = cosCheck.getRange(4, 11).getValue();
  let cosignmentArray = cosCosting.getRange(2, 1, 45, 1).getValues();
  let costRowCount = 0;
  let cosignmentRowCount = 0;
  let cosignmentNewArray = [];
  let j = 0;
  console.log(invoice);
  for (data of costArray) {
    if (data[0] == "") {
      break;
    }
    cosignmentNewArray[j] = [];
    cosignmentNewArray[j] = [
      ...cosignmentNewArray[j],
      invoice,
      from,
      invoiceNum,
      by,
      invoiceDate,
      receivedDate,
      goodReceiveNum,
      data[0],
      data[4],
      data[5],
      data[6],
      data[7],
      data[8],
      data[9],
      signature,
    ];

    j++;
    costRowCount++;
  }

  for (data of cosignmentArray) {
    if (data[0] == invoice) {
      cosignmentRowCount++;
    }
  }

  if (cosignmentRowCount == 0) {
    let cosignmentLastRow = cosCosting.getLastRow() + 1;
    cosCosting
      .getRange(cosignmentLastRow, 1, cosignmentNewArray.length, 15)
      .setValues(cosignmentNewArray);
  } else if (costRowCount > cosignmentRowCount) {
    let cosignmentLastRow = cosignmentRow + cosignmentRowCount;
    cosCosting.insertRowsBefore(
      cosignmentLastRow,
      costRowCount - cosignmentRowCount
    );
    cosCosting
      .getRange(cosignmentRow, 1, cosignmentNewArray.length, 15)
      .setValues(cosignmentNewArray);
  } else {
    cosCosting
      .getRange(cosignmentRow, 1, cosignmentNewArray.length, 15)
      .setValues(cosignmentNewArray);
  }

  e.range.setValue("Done!");
};
