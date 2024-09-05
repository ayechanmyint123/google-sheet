/***********************************
-Global variables. 
************************************/
let ui;
const ss = SpreadsheetApp.getActive();
const items = ss.getSheetByName("Items");
const invoice = ss.getSheetByName("Receipts");
const returnCancel = ss.getSheetByName("ReturnCancel");
const creditPayments = ss.getSheetByName("Credit Payments");
const cashBook = ss.getSheetByName("Cash Book");
const expenseSheet = ss.getSheetByName("Expenses");
const revenueSheet = ss.getSheetByName("Daily Revenue");
const transfers = ss.getSheetByName("Transfers");
const localPurchase = ss.getSheetByName("Local Purchase");
const credits = ss.getSheetByName("Credits");
const categories = ss.getSheetByName("Category");
// const movementData = SpreadsheetApp.openById("199njpCStwK4VN6ViibijH_kLygXxOa0-QpsAnWUBjAY").getSheetByName("Movement Data");
const invCosignments = SpreadsheetApp.openById(
  "1K0k3pbpdvAV4fD_k0Lj2twTDcoh7v_5bRQTx212fcx0"
).getSheetByName("Inv Cosignments");
const transactionsSheet = SpreadsheetApp.openById(
  "1Zy8t1B4hO-kxFVQA_yabrUbdeZgUjTPZr8Q5MMOjW5U"
).getSheetByName("Transactions");
const ygntransfers = SpreadsheetApp.openById(
  "1-BoC0kNN_EBPr47-H4RLXW_QBIHJmtAKeGrLXBSv15E"
).getSheetByName("Transfers");
const ygnItems = SpreadsheetApp.openById(
  "1-BoC0kNN_EBPr47-H4RLXW_QBIHJmtAKeGrLXBSv15E"
).getSheetByName("Items");
const cosignments = SpreadsheetApp.openById(
  "1K0k3pbpdvAV4fD_k0Lj2twTDcoh7v_5bRQTx212fcx0"
).getSheetByName("Cosignments");
// const mdymedicalrevenue = SpreadsheetApp.openById("19p7RwxSVMcaVePlRgUq3jLdI3E_gpJq9erSohFc1fKk").getSheetByName("All Revenue");
let creditPaymentsLastRow = creditPayments.getLastRow();

/***********************************
- Listen for specific trigger spots to launch certain functions
************************************/
let onEditTriggersMDYMedicalShop = (e) => {
  ui = SpreadsheetApp.getUi();
  return e.source.getActiveSheet().getName() == "Cart" &&
    e.range.getRow() == 37 &&
    e.range.getColumn() == 3
    ? onEditCart(e)
    : e.source.getActiveSheet().getName() == "Cart2" &&
      e.range.getRow() == 37 &&
      e.range.getColumn() == 3
    ? onEditCart(e)
    : e.source.getActiveSheet().getName() == "ReturnCancel" &&
      e.range.getColumn() == 1
    ? onEditReturn(e)
    : e.source.getActiveSheet().getName() == "ReturnCancel" &&
      e.range.getRow() == 4 &&
      e.range.getColumn() == 8
    ? onEditCancel(e)
    : e.source.getActiveSheet().getName() == "Credits" &&
      e.range.getRow() == 5 &&
      e.range.getColumn() == 10
    ? onEditCredits(e)
    : e.source.getActiveSheet().getName() == "Cash Book" &&
      e.range.getRow() == 4 &&
      e.range.getColumn() == 14
    ? onEditCashBook(e)
    : e.source.getActiveSheet().getName() == "Transfers" &&
      e.range.getRow() == 2 &&
      e.range.getColumn() == 6
    ? onEditCreateTransfer(e)
    : e.source.getActiveSheet().getName() == "Transfers" &&
      e.range.getColumn() == 9
    ? onEditConfirmTransfer(e)
    : e.source.getActiveSheet().getName() == "Transfers" &&
      e.range.getColumn() == 1
    ? onEditChangeTransfer(e)
    : e.source.getActiveSheet().getName() == "Local Purchase" &&
      e.range.getRow() == 3 &&
      e.range.getColumn() == 6
    ? onEditCreateLocalPurchase(e)
    : e.source.getActiveSheet().getName() == "Local Purchase" &&
      e.range.getColumn() == 1
    ? onEditPayLocalPurchase(e)
    : false;
};

/***********************************
-Calls the newInvoice function 
************************************/
let onEditCart = (e) => {
  let signature = e.value;
  let cart = e.source.getActiveSheet();

  e.range.setValue("Initializing...");

  if (saveCart(signature, cart, e) == false) {
    e.range.setValue("Error!");
    return;
  }

  e.range.setValue("Done!");
};

/***********************************
-Calls the return function 
************************************/
let onEditReturn = (e) => {
  let sheet = e.source.getActiveSheet();
  let code = sheet.getRange(e.range.getRow(), 2).getValue();

  let response = ui.prompt(
    "Enter Return Amount For: " + code,
    ui.ButtonSet.YES_NO
  );

  if (response.getSelectedButton() == ui.Button.YES) {
    let itemAmt = Number(response.getResponseText());
    if (returnItem(invoice, itemAmt, e) == false) {
      return false;
    }
  } else {
    e.range.setValue(false);
  }
};

/***********************************
-Calls the cancel function 
************************************/
let onEditCancel = (e) => {
  let sheet = e.source.getActiveSheet();
  let invoice = sheet.getRange(11, 10).getValue();

  let response = ui.alert(
    "Do you want to cancel: " + invoice,
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    if (cancelInvoice(e) == false) {
      e.range.setValue(false);
      return false;
    }
  } else {
    e.range.setValue(false);
  }
};

/***********************************
-Calls the createTransfer function 
************************************/
let onEditCreateTransfer = (e) => {
  let signature = e.value;

  e.range.setValue("Waiting...");

  if (createTransfer(signature) == false) {
    e.range.setValue("Error!");
    return;
  }

  e.range.setValue("Done!");
};

/***********************************
-Once an invoice is paid it pays in the receipts
************************************/
let onEditCredits = (e) => {
  e.range.setValue("Waiting...");

  let amount = Number(e.value);

  // Display a dialog box with a message, input field, and "Yes" and "No" buttons. The user can
  // also close the dialog by clicking the close button in its title bar.
  let response = ui.alert(
    "Please confirm the amount: " + amount,
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    if (payCredit(amount) == false) {
      e.range.setValue("Error!");
      return;
    }
  } else {
    e.range.setValue("Canceled!");
    return;
  }

  e.range.setValue("Done!");
};

/***********************************
-Calls the newInvoice function 
************************************/
let onEditCashBook = (e) => {
  let signature = e.value;

  e.range.setValue("Waiting...");

  if (dailyExpenses(signature) == false) {
    e.range.setValue("Error!");
    return;
  }

  e.range.setValue("Done!");
};

/***********************************
-Calls the confirm transfer function 
************************************/
let onEditConfirmTransfer = (e) => {
  let invoiceNum = transfers.getRange(e.range.getRow(), 10).getValue();
  let response = ui.alert(
    "Do you want to confirm: " + invoiceNum,
    ui.ButtonSet.YES_NO
  );

  if (invoiceNum == "") {
    ui.alert("Invalid choice!");
    return false;
  }

  if (response == ui.Button.YES) {
    if (confirmTransfer(invoiceNum, e) == false) {
      ui.alert("Not enough stock!");
      e.range.setValue(false);
      return;
    }
  } else {
    e.range.setValue(false);
    return;
  }
};

/***********************************
-Calls the change transfer function 
************************************/
let onEditChangeTransfer = (e) => {
  let trueFalse = e.value;
  if (trueFalse == "FALSE") {
    ui.alert("This transfer is already confirmed!");
    e.range.setValue(true);
    return false;
  }

  let code = transfers.getRange(e.range.getRow(), 6).getValue();
  let response = ui.prompt(
    "Enter Quantity Change For: " + code,
    ui.ButtonSet.YES_NO
  );

  if (response.getSelectedButton() == ui.Button.YES) {
    let qtyAmt = Number(response.getResponseText());
    if (changeTransfer(qtyAmt, e) == false) {
      return false;
    }
  } else {
    e.range.setValue(false);
  }
};

/***********************************
-Calls create local purchase function
************************************/
let onEditCreateLocalPurchase = (e) => {
  let signature = e.value;

  e.range.setValue("Waiting...");

  if (createLocalPurchase(signature) == false) {
    e.range.setValue("Error!");
    return;
  }

  e.range.setValue("Done!");
};

/***********************************
-Calls pay local purchase function
************************************/
let onEditPayLocalPurchase = (e) => {
  let invoiceNum = localPurchase.getRange(e.range.getRow(), 2).getValue();
  let response = ui.alert(
    "Do you want to pay for: " + invoiceNum,
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    if (payLocalPurchase(e) == false) {
      return false;
    }
  } else {
    e.range.setValue(false);
  }
};

/***********************************
- Subtract the items from prespective garage using addItem function
- Save the invoice data form page "Cart" and put it's data into to "Receipts" page
- Split the item amount if the cosignment is split
- Determines whether an invoice is cash or credit. Set variable "Paid" to 0 if credit and "Paid" is set to full amount if cash
- Clear the invoice once data is stored in page "Receipts"
************************************/
let saveCart = (signature, cart, e) => {
  let nextInvoiceRow = invoice.getLastRow() + 1;
  let invoiceNumber = cart.getRange(10, 11).getValue();
  let invoiceDate = cart.getRange(10, 2).getValue();
  let customer = cart.getRange(10, 4).getValue();
  let location = cart.getRange(10, 8).getValue();
  let companyName = cart.getRange(7, 8).getValue();
  let cashCredit = cart.getRange(35, 3).getValue();
  let creditPeriod = cart.getRange(36, 3).getValue();
  let invoiceType = String(cashCredit + "-" + creditPeriod);
  let paidDate = invoiceType == "Cash-0" ? invoiceDate : "";
  let discount = cart.getRange(36, 11).getValue();
  let cartArray = [];
  let cartArrayTwo = cart.getRange(14, 2, 21, 11).getValues();
  let itemColumn = cart.getRange(10, 9).getValue() - 3;
  let stockMovementArray = [];

  //voucher data validation
  e.range.setValue("Validating...");

  if (
    customer.length == 0 ||
    creditPeriod.length == 0 ||
    invoiceDate.length == 0 ||
    signature == undefined ||
    cashCredit.length == 0
  ) {
    ui.alert("Fill in all required fields!");
    return false;
  }

  for (const data of cartArrayTwo) {
    if (data[0].valueOf() == "") {
      break;
    }
    if (data[3].valueOf() == "") {
      ui.alert("Fill in all required fields!");
      return false;
    }
    if (data[3] > items.getRange(data[10], itemColumn).getValue()) {
      ui.alert("Not Enought Stock! " + data[0]);
      return false;
    }

    if (data[3] > invCosignments.getRange(data[10], 20).getValue()) {
      ui.alert("Not Enought Cosignment! " + data[0]);
      return false;
    }
  }

  e.range.setValue("Submitting...");
  let addItemsArray = [];
  let j = 0;
  let l = 0;

  for (const data of cartArrayTwo) {
    if (data[0].valueOf() == "") {
      break;
    }

    let code = data[0];
    let description = data[1];
    let qty = data[3];
    let price = data[7];
    let wholesale = data[8];
    let totalPrice = data[9];
    let itemRow = data[10];
    let finalTotalPrice = totalPrice * (1 - discount);
    let cosignmentArray = invCosignments
      .getRange(itemRow, 2, 1, 10)
      .getValues();
    let twentyTwoStock = items.getRange(itemRow, 7).getValue();
    let eightyFourStock = items.getRange(itemRow, 8).getValue();
    let initialBalance;
    let currentBalance;

    if (location == "22ND") {
      initialBalance = twentyTwoStock;
      currentBalance = twentyTwoStock - qty;
    } else {
      initialBalance = eightyFourStock;
      currentBalance = eightyFourStock - qty;
    }

    addItemsArray[l] = [];
    addItemsArray[l] = [...addItemsArray[l], code, -qty, itemRow, itemColumn];

    stockMovementArray[l] = [];
    stockMovementArray[l] = [
      ...stockMovementArray[l],
      "Sale",
      code,
      invoiceNumber,
      invoiceDate,
      location,
      initialBalance,
      ,
      qty,
      currentBalance,
    ];

    for (cosignment of cosignmentArray) {
      let separateArray = separateCosignment(cosignment, qty);

      for (containers of separateArray) {
        cartArray[j] = [];
        cartArray[j] = [
          ...cartArray[j],
          invoiceNumber + "," + code,
          invoiceNumber,
          location,
          invoiceDate,
          code,
          description,
          containers[1],
          price,
          wholesale,
          discount,
          containers[1] * (finalTotalPrice / qty),
          customer,
          companyName,
          invoiceType == "Cash-0" ? containers[1] * (finalTotalPrice / qty) : 0,
          paidDate,
          signature,
          invoiceType,
          containers[0],
          "Valid",
          containers[0] + "," + code,
        ];
        j++;
      }
    }

    l++;
  }

  invoice
    .getRange(nextInvoiceRow, 1, cartArray.length, 20)
    .setValues(cartArray);

  let invoiceCheck =
    cart.getName() == "Cart"
      ? ss.getSheetByName("Invoice Check")
      : ss.getSheetByName("Invoice Check2");

  invoiceCheck.getRange(11, 10).setValue(invoiceNumber);
  ss.setActiveSheet(invoiceCheck);

  addItems(addItemsArray);

  cart.getRange(10, 4, 1, 1).setValue("");
  cart.getRange(14, 3, 21, 3).setValue("");
  cart.getRange(14, 10, 21, 1).setValue("");
  cart.getRange(35, 3).setValue("");
  cart.getRange(36, 3).setValue("");
  cart.getRange(7, 8).setValue("");
  cart.getRange(36, 11).setValue(0);

  movementData
    .getRange(movementData.getLastRow() + 1, 1, stockMovementArray.length, 9)
    .setValues(stockMovementArray);
};

/***********************************
- Separate Cosignment
************************************/
let separateCosignment = (cosignment, qty) => {
  let cosignmentOneName = cosignment[0].split("|")[0];
  let cosignmentTwoName = cosignment[1].split("|")[0];
  let cosignmentThreeName = cosignment[2].split("|")[0];
  let cosignmentFourName = cosignment[3].split("|")[0];
  let cosignmentFiveName = cosignment[4].split("|")[0];
  let cosignmentSixName = cosignment[5].split("|")[0];
  let cosignmentSevenName = cosignment[6].split("|")[0];
  let cosignmentEightName = cosignment[7].split("|")[0];
  let cosignmentNineName = cosignment[8].split("|")[0];
  let cosignmentTenName = cosignment[9].split("|")[0];

  qty = Number(qty);
  let cosignmentOneQty = Number(
    cosignment[0].split("|")[1] == undefined ? 0 : cosignment[0].split("|")[1]
  );
  let cosignmentTwoQty = Number(
    cosignment[1].split("|")[1] == undefined ? 0 : cosignment[1].split("|")[1]
  );
  let cosignmentThreeQty = Number(
    cosignment[2].split("|")[1] == undefined ? 0 : cosignment[2].split("|")[1]
  );
  let cosignmentFourQty = Number(
    cosignment[3].split("|")[1] == undefined ? 0 : cosignment[3].split("|")[1]
  );
  let cosignmentFiveQty = Number(
    cosignment[4].split("|")[1] == undefined ? 0 : cosignment[4].split("|")[1]
  );
  let cosignmentSixQty = Number(
    cosignment[5].split("|")[1] == undefined ? 0 : cosignment[5].split("|")[1]
  );
  let cosignmentSevenQty = Number(
    cosignment[6].split("|")[1] == undefined ? 0 : cosignment[6].split("|")[1]
  );
  let cosignmentEightQty = Number(
    cosignment[7].split("|")[1] == undefined ? 0 : cosignment[7].split("|")[1]
  );
  let cosignmentNineQty = Number(
    cosignment[8].split("|")[1] == undefined ? 0 : cosignment[8].split("|")[1]
  );
  let cosignmentTenQty = Number(
    cosignment[9].split("|")[1] == undefined ? 0 : cosignment[9].split("|")[1]
  );

  if (
    cosignmentOneQty +
      cosignmentTwoQty +
      cosignmentThreeQty +
      cosignmentFourQty +
      cosignmentFiveQty +
      cosignmentSixQty +
      cosignmentSevenQty +
      cosignmentEightQty +
      cosignmentNineQty <
    qty
  ) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, cosignmentThreeQty],
      [cosignmentFourName, cosignmentFourQty],
      [cosignmentFiveName, cosignmentFiveQty],
      [cosignmentSixName, cosignmentSixQty],
      [cosignmentSevenName, cosignmentSevenQty],
      [cosignmentEightName, cosignmentEightQty],
      [cosignmentNineName, cosignmentNineQty],
      [
        cosignmentTenName,
        qty -
          (cosignmentOneQty +
            cosignmentTwoQty +
            cosignmentThreeQty +
            cosignmentFourQty +
            cosignmentFiveQty +
            cosignmentSixQty +
            cosignmentSevenQty +
            cosignmentEightQty +
            cosignmentNineQty),
      ],
    ];
  } else if (
    cosignmentOneQty +
      cosignmentTwoQty +
      cosignmentThreeQty +
      cosignmentFourQty +
      cosignmentFiveQty +
      cosignmentSixQty +
      cosignmentSevenQty +
      cosignmentEightQty <
    qty
  ) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, cosignmentThreeQty],
      [cosignmentFourName, cosignmentFourQty],
      [cosignmentFiveName, cosignmentFiveQty],
      [cosignmentSixName, cosignmentSixQty],
      [cosignmentSevenName, cosignmentSevenQty],
      [cosignmentEightName, cosignmentEightQty],
      [
        cosignmentNineName,
        qty -
          (cosignmentOneQty +
            cosignmentTwoQty +
            cosignmentThreeQty +
            cosignmentFourQty +
            cosignmentFiveQty +
            cosignmentSixQty +
            cosignmentSevenQty +
            cosignmentEightQty),
      ],
    ];
  } else if (
    cosignmentOneQty +
      cosignmentTwoQty +
      cosignmentThreeQty +
      cosignmentFourQty +
      cosignmentFiveQty +
      cosignmentSixQty +
      cosignmentSevenQty <
    qty
  ) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, cosignmentThreeQty],
      [cosignmentFourName, cosignmentFourQty],
      [cosignmentFiveName, cosignmentFiveQty],
      [cosignmentSixName, cosignmentSixQty],
      [cosignmentSevenName, cosignmentSevenQty],
      [
        cosignmentEightName,
        qty -
          (cosignmentOneQty +
            cosignmentTwoQty +
            cosignmentThreeQty +
            cosignmentFourQty +
            cosignmentFiveQty +
            cosignmentSixQty +
            cosignmentSevenQty),
      ],
    ];
  } else if (
    cosignmentOneQty +
      cosignmentTwoQty +
      cosignmentThreeQty +
      cosignmentFourQty +
      cosignmentFiveQty +
      cosignmentSixQty <
    qty
  ) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, cosignmentThreeQty],
      [cosignmentFourName, cosignmentFourQty],
      [cosignmentFiveName, cosignmentFiveQty],
      [cosignmentSixName, cosignmentSixQty],
      [
        cosignmentSevenName,
        qty -
          (cosignmentOneQty +
            cosignmentTwoQty +
            cosignmentThreeQty +
            cosignmentFourQty +
            cosignmentFiveQty +
            cosignmentSixQty),
      ],
    ];
  } else if (
    cosignmentOneQty +
      cosignmentTwoQty +
      cosignmentThreeQty +
      cosignmentFourQty +
      cosignmentFiveQty <
    qty
  ) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, cosignmentThreeQty],
      [cosignmentFourName, cosignmentFourQty],
      [cosignmentFiveName, cosignmentFiveQty],
      [
        cosignmentSixName,
        qty -
          (cosignmentOneQty +
            cosignmentTwoQty +
            cosignmentThreeQty +
            cosignmentFourQty +
            cosignmentFiveQty),
      ],
    ];
  } else if (
    cosignmentOneQty +
      cosignmentTwoQty +
      cosignmentThreeQty +
      cosignmentFourQty <
    qty
  ) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, cosignmentThreeQty],
      [cosignmentFourName, cosignmentFourQty],
      [
        cosignmentFiveName,
        qty -
          (cosignmentOneQty +
            cosignmentTwoQty +
            cosignmentThreeQty +
            cosignmentFourQty),
      ],
    ];
  } else if (cosignmentOneQty + cosignmentTwoQty + cosignmentThreeQty < qty) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, cosignmentThreeQty],
      [
        cosignmentFourName,
        qty - (cosignmentOneQty + cosignmentTwoQty + cosignmentThreeQty),
      ],
    ];
  } else if (cosignmentOneQty + cosignmentTwoQty < qty) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, cosignmentTwoQty],
      [cosignmentThreeName, qty - (cosignmentOneQty + cosignmentTwoQty)],
    ];
  } else if (cosignmentOneQty < qty) {
    return [
      [cosignmentOneName, cosignmentOneQty],
      [cosignmentTwoName, qty - cosignmentOneQty],
    ];
  } else {
    return [[cosignmentOneName, qty]];
  }
};

/***********************************
- Add Items function
************************************/
let addItems = (itemArray) => {
  for (data of itemArray) {
    let newAmt = Number(items.getRange(data[2], data[3]).getValue()) + data[1];
    if (newAmt < 0) {
      return false;
    }
  }
  for (data of itemArray) {
    let newAmt = Number(items.getRange(data[2], data[3]).getValue()) + data[1];
    items.getRange(data[2], data[3]).setValue(newAmt);
  }
};

/***********************************
- Return function
************************************/
let returnItem = (invoice, itemAmount, e) => {
  let date = returnCancel.getRange(11, 2).getValue();
  let location = returnCancel.getRange(11, 3).getValue();
  let customer = returnCancel.getRange(11, 4).getValue();
  let invoiceNumber = returnCancel.getRange(11, 10).getValue();
  let nextInvoiceRow = invoice.getLastRow() + 1;
  let leftamount = returnCancel.getRange(5, 13).getValue();
  let payamount = returnCancel.getRange(6, 13).getValue();
  let flagCredit = returnCancel.getRange(46, 3).getValue();
  flagCredit = payamount == leftamount ? false : true;
  let code = returnCancel.getRange(e.range.getRow(), 2).getValue();
  let description = returnCancel.getRange(e.range.getRow(), 3).getValue();
  let qty = returnCancel.getRange(e.range.getRow(), 5).getValue();
  let price = returnCancel.getRange(e.range.getRow(), 8).getValue();
  let finalTotalPrice = returnCancel.getRange(e.range.getRow(), 11).getValue();
  let discount = returnCancel.getRange(47, 11).getValue();
  let wholesale = returnCancel.getRange(e.range.getRow(), 9).getValue();
  finalTotalPrice = (finalTotalPrice * (1 - discount)) / qty;
  let company = returnCancel.getRange(11, 5).getValue();
  let cosignment = returnCancel.getRange(e.range.getRow(), 12).getValue();
  let signature = returnCancel.getRange(47, 3).getValue();
  let invoiceArray = returnCancel.getRange(15, 2, 31, 3).getValues();
  let sumAmt = returnCancel.getRange(e.range.getRow(), 13).getValue();
  let itemRow = returnCancel.getRange(e.range.getRow(), 14).getValue();
  let itemColumn = location == "22ND" ? 7 : 8;
  let creditLocation = returnCancel.getRange(11, 13).getValue();
  let newAmount = credits.getRange(creditLocation, 12).getValue();
  let dateReturn = returnCancel.getRange(11, 7).getValue();
  let twentyTwoStock = items.getRange(itemRow, 7).getValue();
  let eightyFourStock = items.getRange(itemRow, 8).getValue();
  let allAmount = Math.round(finalTotalPrice * itemAmount);
  let stockMovementArray;

  let initialBalance;
  let currentBalance;

  if (location == "22ND") {
    initialBalance = twentyTwoStock;
    currentBalance = twentyTwoStock + itemAmount;
  } else {
    initialBalance = eightyFourStock;
    currentBalance = eightyFourStock + itemAmount;
  }

  if (qty < 0) {
    e.range.setValue(false);
    ui.alert("This item has already been returned.");
    return false;
  }

  if (sumAmt == 0 || itemAmount > sumAmt) {
    e.range.setValue(false);
    ui.alert("Return amount is larger than sold amount.");
    return false;
  }

  let returnArray = [
    [
      invoiceNumber + "," + code,
      invoiceNumber,
      location,
      dateReturn,
      code,
      description,
      -itemAmount,
      price,
      wholesale,
      discount,
      -allAmount,
      customer,
      company,
      -allAmount,
      dateReturn,
      signature,
      flagCredit ? "Return-Credit" : "Return-Cash",
      cosignment,
      "Valid",
      cosignment + "," + code,
    ],
  ];

  stockMovementArray = [
    [
      "Return",
      code,
      invoiceNumber,
      dateReturn,
      location,
      initialBalance,
      itemAmount,
      ,
      currentBalance,
    ],
  ];

  invoice
    .getRange(nextInvoiceRow, 1, returnArray.length, 20)
    .setValues(returnArray);
  itemArray = [[code, itemAmount, itemRow, itemColumn]];
  addItems(itemArray);

  creditPaymentArray = [
    [invoiceNumber, dateReturn, customer, "", "", allAmount, "Return"],
  ];

  if (flagCredit) {
    creditPayments
      .getRange(creditPaymentsLastRow + 1, 1, creditPaymentArray.length, 7)
      .setValues(creditPaymentArray);
    credits.getRange(creditLocation, 12).setValue(newAmount + allAmount);
  }

  movementData
    .getRange(movementData.getLastRow() + 1, 1, stockMovementArray.length, 9)
    .setValues(stockMovementArray);
  e.range.setValue(false);
};

/***********************************
- Cancel function
************************************/
let cancelInvoice = (e) => {
  let startInvoiceRow = returnCancel.getRange(11, 14).getValue();
  let invoiceNum = returnCancel.getRange(11, 10).getValue();
  let receiptArray = invoice.getRange(startInvoiceRow, 2, 21, 18).getValues();
  let payingCredit = returnCancel.getRange(3, 13).getValue();
  let date = returnCancel.getRange(11, 7).getValue();
  let customerName = returnCancel.getRange(11, 4).getValue();
  let j = 0;
  let l = 0;
  let alreadyCancel =
    invoice.getRange(startInvoiceRow, 19).getValue() == "Canceled"
      ? true
      : false;
  let cartArray = returnCancel.getRange(15, 2, 31, 13).getValues();
  let location = returnCancel.getRange(11, 3).getValue();
  let addItemsArray = [];
  let stockMovementArray = [];
  if (alreadyCancel) {
    ui.alert("Error, This invoice is already canceled.");
    return false;
  }

  if (payingCredit != "") {
    ui.alert("Error, This credit invoice is currently being paid.");
    return false;
  }

  for (let i = 0; i < 21; i++) {
    if (receiptArray[i][0] != invoiceNum) {
      break;
    }
    j++;
  }

  for (data of cartArray) {
    if (data[0] == "") {
      break;
    }

    let code = data[0];
    let qty = data[3];
    let itemRow = data[12];
    let itemCol = location == "22ND" ? 7 : 8;
    let twentyTwoStock = items.getRange(itemRow, 7).getValue();
    let eightyFourStock = items.getRange(itemRow, 8).getValue();
    let initialBalance;
    let currentBalance;

    if (location == "22ND") {
      initialBalance = twentyTwoStock;
      currentBalance = twentyTwoStock + qty;
    } else {
      initialBalance = eightyFourStock;
      currentBalance = eightyFourStock + qty;
    }

    addItemsArray[l] = [];
    addItemsArray[l] = [...addItemsArray[l], code, qty, itemRow, itemCol];

    stockMovementArray[l] = [];
    stockMovementArray[l] = [
      ...stockMovementArray[l],
      "Cancel",
      code,
      invoiceNum,
      date,
      location,
      initialBalance,
      qty,
      ,
      currentBalance,
    ];
    l++;
  }
  addItems(addItemsArray);
  invoice.getRange(startInvoiceRow, 19, j, 1).setValue("Canceled");
  e.range.setValue(false);
  movementData
    .getRange(movementData.getLastRow() + 1, 1, stockMovementArray.length, 9)
    .setValues(stockMovementArray);
};

/***********************************
- Pay Credit 
************************************/
let payCredit = (amount) => {
  let date = credits.getRange(2, 3).getValue();
  let customer = credits.getRange(2, 6).getValue();
  let amtLeft = credits.getRange(5, 9).getValue();
  let customerPosition = credits.getRange(2, 4).getValue();
  let currentAmt = credits.getRange(customerPosition, 12).getValue();
  let totalAmt = credits.getRange(5, 7).getValue();
  let sumAmt = credits.getRange(5, 8).getValue();
  let creditInvoiceRowInReceipt = credits.getRange(8, 10).getValue();
  let creditInvoiceNum = credits.getRange(5, 6).getValue();
  let creditPaymentsLastRow = creditPayments.getLastRow();
  //let creditMovementLastRow=creditMovement.getLastRow();

  if (amount > amtLeft) {
    ui.alert("Pay amount is greater than Amount Left!");
    return false;
  } else if (amount == totalAmt) {
    for (let j = 0; j < 21; j++) {
      let invoiceNum = invoice
        .getRange(creditInvoiceRowInReceipt + j, 2)
        .getValue();
      let salePrice = invoice
        .getRange(creditInvoiceRowInReceipt + j, 11)
        .getValue();

      if (invoiceNum != creditInvoiceNum) {
        break;
      }

      invoice.getRange(creditInvoiceRowInReceipt + j, 14).setValue(salePrice);
      invoice.getRange(creditInvoiceRowInReceipt + j, 15).setValue(date);
    }

    creditArray = [
      [creditInvoiceNum, date, customer, totalAmt, amtLeft, amount, "Pay"],
    ];
    creditPayments
      .getRange(creditPaymentsLastRow + 1, 1, 1, 7)
      .setValues(creditArray);
    //creditMovement.getRange(creditMovementLastRow+1,1,1,7).setValues(creditArray);
  } else if (amount == amtLeft) {
    for (let j = 0; j < 21; j++) {
      let invoiceNum = invoice
        .getRange(creditInvoiceRowInReceipt + j, 2)
        .getValue();
      let salePrice = invoice
        .getRange(creditInvoiceRowInReceipt + j, 11)
        .getValue();

      if (invoiceNum != creditInvoiceNum) {
        break;
      }

      invoice.getRange(creditInvoiceRowInReceipt + j, 14).setValue(salePrice);
      invoice.getRange(creditInvoiceRowInReceipt + j, 15).setValue(date);
    }

    creditArray = [
      [creditInvoiceNum, date, customer, totalAmt, amtLeft, amount, "Pay"],
    ];
    creditPayments
      .getRange(creditPaymentsLastRow + 1, 1, 1, 7)
      .setValues(creditArray);
    // creditMovement.getRange(creditMovementLastRow+1,1,1,7).setValues(creditArray);

    if (sumAmt == currentAmt) {
      credits.getRange(customerPosition, 12).setValue(0);
    } else {
      credits.getRange(customerPosition, 12).setValue(currentAmt - sumAmt);
    }
  } else if (amount < amtLeft) {
    credits.getRange(customerPosition, 12).setValue(currentAmt + amount);
    creditArray = [
      [creditInvoiceNum, date, customer, totalAmt, amtLeft, amount, "Pay"],
    ];
    creditPayments
      .getRange(creditPaymentsLastRow + 1, 1, 1, 7)
      .setValues(creditArray);
    //creditMovement.getRange(creditMovementLastRow+1,1,1,7).setValues(creditArray);
  }
  credits.getRange(5, 6).setValue("");
};

/***********************************
-Record Daily Expenses
************************************/
let dailyExpenses = (signature) => {
  let date = cashBook.getRange(4, 7).getValue();
  let currentDate = new Date();
  let expenses = cashBook.getRange(8, 12, 28, 4).getValues();
  let expenseLastRow = expenseSheet.getLastRow() + 1;
  let revenueLastRow = revenueSheet.getLastRow() + 1;
  let expenseDate = expenseSheet.getRange(2, 1, expenseLastRow, 1).getValues();
  let expenseType = expenseSheet.getRange(2, 2, expenseLastRow, 1).getValues();
  date =
    new Date(date).getYear() +
    "," +
    new Date(date).getMonth() +
    "," +
    new Date(date).getDate();

  for (let i = 0; i < expenseDate.length; i++) {
    let checkDate =
      new Date(expenseDate[i][0]).getYear() +
      "," +
      new Date(expenseDate[i][0]).getMonth() +
      "," +
      new Date(expenseDate[i][0]).getDate();
    let checkExpense = expenseType[i][0];

    if (date == checkDate && checkExpense == "MDY-MEDICAL") {
      ui.alert("Cashbook for this date has already been recorded!");
      return false;
    }
  }

  let expenseArray = [];
  let j = 0;

  for (let i = 0; i < expenses.length; i++) {
    if (expenses[i][0] == "" && expenses[i][1] == "" && expenses[i][2] == "") {
      break;
    }
    let category = expenses[i][0];
    let invoiceNumber = expenses[i][1];
    let expense = expenses[i][2];
    let amount = expenses[i][3];

    if (expense.length == 0 || amount.length == 0) {
      ui.alert("Fill in all required fields!");
      return false;
    }

    expenseArray[j] = [];
    expenseArray[j] = [
      ...expenseArray[j],
      cashBook.getRange(4, 7).getValue(),
      "MDY-MEDICAL",
      expense,
      amount,
      signature,
      category,
      currentDate,
      invoiceNumber,
    ];

    j++;
  }
  expenseSheet
    .getRange(expenseLastRow, 1, expenseArray.length, 8)
    .setValues(expenseArray);
  cashBook.getRange(8, 12, 28, 4).setValue("");
  let revenue = cashBook.getRange(12, 11).getValue();
  dailyRevenue(
    signature,
    cashBook.getRange(4, 7).getValue(),
    revenue,
    revenueLastRow
  );
  return true;
};

/***********************************
-Record Daily Revenue In Main Accounting
************************************/
let dailyRevenue = (signature, date, revenue, revenueLastRow) => {
  let dayRevenue = cashBook.getRange(4, 11).getValue();
  let dayIncome = cashBook.getRange(8, 11).getValue();
  let system =
    revenueSheet.getRange(revenueSheet.getLastRow(), 7).getValue() + 1;
  let revenueArray = [
    [date, "MDY-MEDICAL", revenue, dayRevenue, dayIncome, signature, system],
  ];
  revenueSheet.getRange(revenueLastRow, 1, 1, 7).setValues(revenueArray);

  let transactionNumArray = transactionsSheet
    .getRange(transactionsSheet.getLastRow() - 20, 1, 20, 1)
    .getValues();
  let transactionNum = transactionNumArray.reduce((a, b) => {
    return Math.max(a, b);
  });

  let systemRevenue = "System-Revenue-" + revenueSheet.getLastRow();

  let array = [
    [transactionNum + 1, date, systemRevenue, "Daily Revenue", , , ,],
    [, , , 96008, "Open Trade Debtor-MDY Medical", dayRevenue, ,],
    [, , , 70029, "Revenue-Open Medical (MDY)", , dayRevenue],
    [transactionNum + 2, date, systemRevenue, "Daily Income", , , ,],
    [, , , 10001, "Cash in Hand-MDY Medical", dayIncome, ,],
    [, , , 96008, "Open Trade Debtor-MDY Medical", , dayIncome],
  ];

  let arrayRevenue = [
    [date, systemRevenue, "Daily Revenue", , , ,],
    [, , 96008, "Open Trade Debtor-MDY Medical", dayRevenue, ,],
    [, , 70029, "Revenue-Open Medical (MDY)", , dayRevenue],
    [date, systemRevenue, "Daily Income", , , ,],
    [, , 10001, "Cash in Hand-MDY Medical", dayIncome, ,],
    [, , 96008, "Open Trade Debtor-MDY Medical", , dayIncome],
  ];

  transactionsSheet
    .getRange(transactionsSheet.getLastRow() + 1, 1, 6, 7)
    .setValues(array);
  mdymedicalrevenue
    .getRange(mdymedicalrevenue.getLastRow() + 1, 1, 6, 6)
    .setValues(arrayRevenue);
};

/***********************************
-Create Transfer
************************************/
let createTransfer = (signature) => {
  let transferArrayOne = transfers.getRange(2, 2, 1, 5).getValues();
  let transferArrayTwo = transfers.getRange(4, 2, 15, 6).getValues();
  let transferLastRow = transfers.getRange(1, 8).getValue() + 1;

  let invoiceNum = transferArrayOne[0][3];
  let newInvoiceNum = Number(invoiceNum.split("-")[1]) + 1;
  newInvoiceNum = "mdymedicaltransfer-" + newInvoiceNum;
  let date = transferArrayOne[0][2];
  let from = transferArrayOne[0][0];
  let to = transferArrayOne[0][1];

  let itemArray = [];
  let j = 0;

  for (item of transferArrayTwo) {
    let code = item[0];
    let description = item[1];
    let qty = item[4];
    let remark = item[5];

    if (code == "") {
      break;
    }

    itemArray[j] = [];
    itemArray[j] = [
      ...itemArray[j],
      ,
      invoiceNum,
      date,
      from,
      to,
      code,
      description,
      qty,
      signature,
      remark,
      "No",
      "=MATCH(F" + Number(transferLastRow + j) + ",Items!A:A,0)",
    ];

    j++;
  }

  transfers
    .getRange(transferLastRow, 1, itemArray.length, 12)
    .setValues(itemArray);
  transfers
    .getRange(transferLastRow, 1, itemArray.length, 1)
    .insertCheckboxes();
  transfers.getRange(2, 5).setValue(newInvoiceNum);
  transfers.getRange(2, 2, 1, 2).setValue("");
  transfers.getRange(4, 3, 15, 5).setValue("");
};

/***********************************
-Confirm Transfer
************************************/
let confirmTransfer = (invoiceNum, e) => {
  let invoiceStartRow = transfers.getRange(e.range.getRow(), 13).getValue();
  let invoiceRowArray = [];
  let invoiceArray = transfers.getRange(invoiceStartRow, 2, 15, 12).getValues();
  let addItemsArray = [];
  let addItemsArrayTwo = [];
  let itemArrayTwo = [];
  let stockMovementArray = [];
  let stockMovementArrayTwo = [];
  let l = 0;
  let ygnTransferLastRow = ygntransfers.getLastRow() + 1;

  for (data of invoiceArray) {
    if (data[0] != invoiceNum) {
      break;
    }

    let date = data[1];
    let code = data[4];
    let qty = data[6];
    let itemRow = data[10];
    let from = data[2];
    let locationOne = from;
    let to = data[3];
    let locationTwo = to;
    let description = data[5];
    let signature = data[7];
    let remark = data[8];
    let twentyTwoStock = items.getRange(itemRow, 7).getValue();
    let eightyFourStock = items.getRange(itemRow, 8).getValue();
    let tgn = ygnItems.getRange(itemRow, 7).getValue();
    let dagon = ygnItems.getRange(itemRow, 8).getValue();
    let kyimyindine = ygnItems.getRange(itemRow, 9).getValue();
    let jsix = ygnItems.getRange(itemRow, 10).getValue();

    itemArrayTwo[l] = [];
    itemArrayTwo[l] = [
      ...itemArrayTwo[l],
      true,
      invoiceNum,
      date,
      from,
      to,
      code,
      description,
      qty,
      signature,
      remark,
      "No",
      "=MATCH(F" + Number(ygnTransferLastRow + l) + ",Items!A:A,0)",
    ];

    if (from == "22ND" || from == "84TH") {
      if (from == "22ND") {
        from = 7;
      }
      if (from == "84TH") {
        from = 8;
      }

      addItemsArray[l] = [];
      addItemsArray[l] = [...addItemsArray[l], code, -qty, itemRow, from];

      let initialBalanceOne;
      let currentBalanceOne;

      if (locationOne == "22ND") {
        initialBalanceOne = twentyTwoStock;
        currentBalanceOne = twentyTwoStock - qty;
      } else {
        initialBalanceOne = eightyFourStock;
        currentBalanceOne = eightyFourStock - qty;
      }

      stockMovementArray[l] = [];
      stockMovementArray[l] = [
        ...stockMovementArray[l],
        "Transfer",
        code,
        invoiceNum,
        date,
        locationOne,
        initialBalanceOne,
        ,
        qty,
        currentBalanceOne,
      ];
    }

    if (to == "22ND" || to == "84TH") {
      if (to == "22ND") {
        to = 7;
      }
      if (to == "84TH") {
        to = 8;
      }

      addItemsArrayTwo[l] = [];
      addItemsArrayTwo[l] = [...addItemsArrayTwo[l], code, qty, itemRow, to];

      let initialBalanceTwo;
      let currentBalanceTwo;

      if (locationTwo == "22ND") {
        initialBalanceTwo = twentyTwoStock;
        currentBalanceTwo = twentyTwoStock + qty;
      } else {
        initialBalanceTwo = eightyFourStock;
        currentBalanceTwo = eightyFourStock + qty;
      }

      stockMovementArrayTwo[l] = [];
      stockMovementArrayTwo[l] = [
        ...stockMovementArrayTwo[l],
        "Transfer",
        code,
        invoiceNum,
        date,
        locationTwo,
        initialBalanceTwo,
        qty,
        ,
        currentBalanceTwo,
      ];
    }

    invoiceRowArray.push(invoiceStartRow + l);
    l++;
  }

  if (addItemsArray.length != 0) {
    if (addItems(addItemsArray) == false) {
      return false;
    }
  }

  if (addItemsArrayTwo.length != 0) {
    if (addItems(addItemsArrayTwo) == false) {
      return false;
    }
  }

  let to = itemArrayTwo[0][4];
  if (to == "TGN" || to == "DAGON" || to == "KyiMyinDine" || to == "JSIX") {
    ygntransfers
      .getRange(ygnTransferLastRow, 1, itemArrayTwo.length, 12)
      .setValues(itemArrayTwo);
    ygntransfers
      .getRange(ygnTransferLastRow, 1, itemArrayTwo.length, 1)
      .insertCheckboxes();
  }

  for (let i = 0; i < invoiceRowArray.length; i++) {
    transfers.getRange(Number(invoiceRowArray[i]), 11).setValue("Yes");
    transfers.getRange(Number(invoiceRowArray[i]), 1).setValue(true);
  }

  e.range.setValue(false);
  if (stockMovementArray.length != 0) {
    movementData
      .getRange(movementData.getLastRow() + 1, 1, stockMovementArray.length, 9)
      .setValues(stockMovementArray);
  }

  if (stockMovementArrayTwo.length != 0) {
    movementData
      .getRange(
        movementData.getLastRow() + 1,
        1,
        stockMovementArrayTwo.length,
        9
      )
      .setValues(stockMovementArrayTwo);
  }
};

/***********************************
-Change Transfer Quantity
************************************/
let changeTransfer = (qtyAmt, e) => {
  transfers.getRange(e.range.getRow(), 8).setValue(qtyAmt);
  e.range.setValue(false);
};

/***********************************
-Create Local Purchase
************************************/
let createLocalPurchase = (signature) => {
  let dataArray = localPurchase.getRange(3, 2, 1, 5).getValues();
  let date = dataArray[0][0];
  let supplier = dataArray[0][1];
  let invoiceNum = dataArray[0][2];
  let newInvoiceNum = Number(invoiceNum.split("-")[1]) + 1;
  newInvoiceNum = "mdyMedLocalPurchase-" + newInvoiceNum;
  let cashCredit = dataArray[0][3];
  let cosignmentArray = [];
  let recordArray = [];
  let expenseArray = [];
  let addItemsArray = [];
  let stockMovementArray = [];
  let j = 0;
  let lastCoRow = cosignments.getLastRow() + 1;
  let lastPurchaseRow = localPurchase.getRange(1, 9).getValue() + 1;
  let totalExpense = 0;

  let invoiceArray = localPurchase.getRange(5, 3, 10, 7).getValues();

  //for cosignments
  for (data of invoiceArray) {
    let concat = invoiceNum + "," + data[0];
    let code = data[0];
    let description = data[1];
    let qty = data[3];
    let cost = data[4];
    let totalCost = data[5];
    let itemRow = data[6];

    totalExpense += totalCost;

    if (code == "") {
      break;
    }

    let twentyTwoStock = items.getRange(itemRow, 7).getValue();

    initialBalance = twentyTwoStock;
    currentBalance = twentyTwoStock + qty;

    cosignmentArray[j] = [];
    cosignmentArray[j] = [
      ...cosignmentArray[j],
      concat,
      invoiceNum,
      code,
      date,
      qty,
      "=SUMIF('YGN Receipt'!A:E,A" +
        Number(lastCoRow) +
        ",'YGN Receipt'!F:F)+SUMIF('MDY Receipt'!A:E,A" +
        Number(lastCoRow) +
        ",'MDY Receipt'!F:F)",
      "=E" + lastCoRow + "-F" + lastCoRow,
      cost,
      signature,
      "=if(G" + lastCoRow + "=0,FALSE,TRUE)",
      "32000",
      "22ND",
      ,
      "=E" + lastCoRow + "*H" + lastCoRow + "",
      "=CONCAT(CONCAT(B" +
        Number(lastCoRow) +
        ',"|"),G' +
        Number(lastCoRow) +
        ")",
      "=H" +
        lastCoRow +
        "*G" +
        lastCoRow +
        "/(COUNT(SPLIT(K" +
        lastCoRow +
        ',"+")))',
      "=H" + lastCoRow + "*G" + lastCoRow + "",
    ];

    recordArray[j] = [];
    recordArray[j] = [
      ...recordArray[j],
      invoiceNum,
      date,
      supplier,
      code,
      description,
      qty,
      cost,
      totalCost,
      cashCredit,
      cashCredit == "Cash" ? totalCost : 0,
      cashCredit == "Cash" ? date : "",
      signature,
    ];

    addItemsArray[j] = [];
    addItemsArray[j] = [...addItemsArray[j], code, qty, itemRow, 7];

    stockMovementArray[j] = [];
    stockMovementArray[j] = [
      ...stockMovementArray[j],
      "Local Purchase",
      code,
      invoiceNum,
      date,
      "Garage1",
      initialBalance,
      qty,
      ,
      currentBalance,
    ];

    expenseArray[0] = [
      date,
      invoiceNum,
      " from " + supplier,
      totalExpense,
      signature,
    ];

    j++;
    lastCoRow++;
  }

  lastCoRow = cosignments.getLastRow() + 1;

  if (addItems(addItemsArray) == false) {
    return false;
  }
  cosignments
    .getRange(lastCoRow, 1, cosignmentArray.length, 17)
    .setValues(cosignmentArray);
  localPurchase
    .getRange(lastPurchaseRow, 10, recordArray.length, 12)
    .setValues(recordArray);
  localPurchase.getRange(3, 4).setValue(newInvoiceNum);
  if (cashCredit == "Cash") {
    expenseSheet
      .getRange(expenseSheet.getLastRow() + 1, 1, 1, 5)
      .setValues(expenseArray);
  }
  localPurchase.getRange(5, 4, 10, 4).setValue("");
  localPurchase.getRange(3, 3).setValue("");
  localPurchase.getRange(3, 5).setValue("");
  movementData
    .getRange(movementData.getLastRow() + 1, 1, stockMovementArray.length, 9)
    .setValues(stockMovementArray);
};

/***********************************
-Pay Local Purchase Credit
************************************/
let payLocalPurchase = (e) => {
  let invoiceArray = localPurchase
    .getRange(e.range.getRow(), 2, 1, 5)
    .getValues();
  let invoiceNum = invoiceArray[0][0];
  let supplier = invoiceArray[0][2];
  let amount = invoiceArray[0][3];
  let startRow = invoiceArray[0][4];
  let date = localPurchase.getRange(3, 2).getValue();
  let localPurchaseArray = localPurchase
    .getRange(startRow, 10, 10, 12)
    .getValues();
  let payArray = [];
  let expenseArray = [];
  let totalExpense = 0;
  let j = 0;

  for (data of localPurchaseArray) {
    if (data[0] != invoiceNum) {
      break;
    }

    let totalCost = data[7];
    totalExpense += totalCost;

    payArray[j] = [];
    payArray[j] = [...payArray[j], totalCost, date];

    j++;
  }
  expenseArray[0] = [
    date,
    invoiceNum,
    " from " + supplier,
    totalExpense,
    "Pay Credit LP",
  ];
  expenseSheet
    .getRange(expenseSheet.getLastRow() + 1, 1, 1, 5)
    .setValues(expenseArray);
  localPurchase.getRange(startRow, 19, payArray.length, 2).setValues(payArray);
  e.range.setValue(false);
};

/***********************************
-Record Daily Expenses In Main Accounting
************************************/
let dailyExpensesCategories = () => {
  let dayExpenses = expenseSheet.getRange("A2:H").getValues();
  let categoriesFilter = categories.getRange("A2:B").getValues();

  let currentDate = new Date();
  let currentExpensesByDate = dayExpenses.filter((expense) => {
    return (
      new Date(expense[6]).getYear() === currentDate.getYear() &&
      new Date(expense[6]).getMonth() === currentDate.getMonth() &&
      new Date(expense[6]).getDate() === currentDate.getDate()
    );
  });

  let transactionNumArray = transactionsSheet
    .getRange(transactionsSheet.getLastRow() - 20, 1, 20, 1)
    .getValues();
  let transactionNum = transactionNumArray.reduce((a, b) => {
    return Math.max(a, b);
  });

  let creditAccountNum = 10001;
  let creditAccount = "Cash in Hand-MDY Medical";
  let transactionArray = [];
  currentExpensesByDate.forEach((expense, idx) => {
    const accountCode = categoriesFilter.find(
      (category) => category[0] === expense[5]
    );

    let array = [
      transactionNum + idx + 1,
      expense[0],
      expense[7],
      expense[2],
      "",
      "",
      "",
    ];
    let debitArray = [
      "",
      "",
      "",
      accountCode[1],
      accountCode[0],
      expense[3],
      "",
    ];
    let creditArray = [
      "",
      "",
      "",
      creditAccountNum,
      creditAccount,
      "",
      expense[3],
    ];

    transactionArray.push(array, debitArray, creditArray);
  });

  if (transactionArray.length >= 3) {
    transactionsSheet
      .getRange(
        transactionsSheet.getLastRow() + 1,
        1,
        transactionArray.length,
        7
      )
      .setValues(transactionArray);
  }
  //console.log(transactionNum)
  return;
};
