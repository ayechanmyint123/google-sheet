/***********************************
-Global variables. 
************************************/
const ss = SpreadsheetApp.getActive();
const trial = ss.getSheetByName("Trial Balance");
const accounts = ss.getSheetByName("Charts Of Accounts");
const newJournal = ss.getSheetByName("New Journal");
const newJournal2 = ss.getSheetByName("New Journal 2");
const transactions = ss.getSheetByName("Transactions");
const expenses = ss.getSheetByName("Shop Expenses");
const accountReport = ss.getSheetByName("Account Report");
//Get All Revenues
const allRevenues = SpreadsheetApp.openById(
  "19p7RwxSVMcaVePlRgUq3jLdI3E_gpJq9erSohFc1fKk"
).getSheetByName("All Revenue");
const ui = SpreadsheetApp.getUi();
let transactionsLastRow = transactions.getLastRow() + 1;

/***********************************
- Listen for specific trigger spots to launch certain functions
************************************/

let onEditMainAccountingTriggers = (e) => {
  return e.source.getActiveSheet().getName() == "New Journal" &&
    e.range.getColumn() == 3
    ? linkAccount(e)
    : e.source.getActiveSheet().getName() == "New Journal 2" &&
      e.range.getColumn() == 3
    ? linkAccount(e)
    : e.source.getActiveSheet().getName() == "New Journal" &&
      e.range.getRow() == 2 &&
      e.range.getColumn() == 10
    ? divideTransaction(e)
    : e.source.getActiveSheet().getName() == "New Journal 2" &&
      e.range.getRow() == 2 &&
      e.range.getColumn() == 10
    ? divideTransaction(e)
    : e.source.getActiveSheet().getName() == "Trial Balance" &&
      e.range.getRow() == 3 &&
      e.range.getColumn() == 6
    ? checkTrial(e)
    : e.source.getActiveSheet().getName() == "New Journal" &&
      e.range.getRow() == 5 &&
      e.range.getColumn() == 10
    ? createTransaction(e)
    : e.source.getActiveSheet().getName() == "New Journal 2" &&
      e.range.getRow() == 5 &&
      e.range.getColumn() == 10
    ? createTransaction(e)
    : e.source.getActiveSheet().getName() == "Account Report" &&
      e.range.getRow() == 2 &&
      e.range.getColumn() == 8
    ? generateReport(e)
    : e.source.getActiveSheet().getName() == "Account Report 2" &&
      e.range.getRow() == 2 &&
      e.range.getColumn() == 8
    ? generateReport(e)
    : e.source.getActiveSheet().getName() == "Account Report 3" &&
      e.range.getRow() == 2 &&
      e.range.getColumn() == 8
    ? generateReport(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 1
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 7
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 13
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 19
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 25
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 31
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 37
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Shop Expenses" &&
      e.range.getColumn() == 43
    ? categorization(e)
    : e.source.getActiveSheet().getName() == "Account Report" &&
      e.range.getColumn() == 9
    ? loadTransactionForm(e)
    : e.source.getActiveSheet().getName() == "Transactions" &&
      e.range.getColumn() == 9
    ? syncShopRevenues(e)
    : false;
};

/***********************************
- Get Sub accounts from main account
************************************/
let linkAccount = (e) => {
  let newJournal = e.source.getActiveSheet();
  let mainAccount = e.range.getValue();
  let accountsArray = accounts.getRange("D5:E").getValues();

  let filteredArray = accountsArray.filter((data) => {
    return data[0] == mainAccount;
  });
  filteredArray = filteredArray.map((value) => {
    return value[1];
  });
  let validateRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(filteredArray)
    .setAllowInvalid(false)
    .build();
  newJournal.getRange(e.range.getRow(), 4).setDataValidation(validateRule);
};

/***********************************
- Sync Shop Revenues
************************************/
let syncShopRevenues = (e) => {
  let allRevenuesLastRow = allRevenues.getLastRow();
  let allRevenuesArray = allRevenues
    .getRange(2, 1, allRevenuesLastRow - 1, 6)
    .getValues();
};

/***********************************
- Create Transaction
************************************/
let divideTransaction = (e) => {
  let newJournal = e.source.getActiveSheet();
  newJournal.insertRowAfter(newJournal.getLastRow());
  let lastRow = newJournal.getLastRow() + 1;
  newJournal
    .getRange(lastRow, 9)
    .setValue(
      "=iferror(VLOOKUP(D" +
        lastRow +
        ",'Charts Of Accounts'!E:F,2,false),\"\")"
    );
  e.range.setValue(false);
};

/***********************************
- Create Transaction
************************************/
let createTransaction = (e) => {
  let newJournal = e.source.getActiveSheet();
  let lastRow = newJournal.getLastRow();
  let newJournalArray = newJournal.getRange(6, 2, lastRow - 5, 8).getValues();
  let transactionNum = newJournal.getRange(2, 4).getValue();
  let debit = newJournal.getRange(3, 7).getValue();
  let credit = newJournal.getRange(3, 8).getValue();
  let date = newJournalArray[0][0];
  let invoiceNum = newJournalArray[0][3];
  let memo = newJournalArray[0][4];

  if (transactionNum == "" || date == "" || invoiceNum == "") {
    ui.alert("Fill in all required fields");
    return false;
  }

  if (credit == debit) {
    let transactionArray = [
      [transactionNum, date, invoiceNum, memo, "", "", ""],
    ];

    for (data of newJournalArray) {
      if (
        (Number(data[7]) >= 20000 && Number(data[7]) < 35000) ||
        (Number(data[7]) >= 70000 && Number(data[7]) < 70015) ||
        (Number(data[7]) >= 75000 && Number(data[7]) < 75006)
      ) {
        ui.alert("Invalid Credit or Debit Account");
        return false;
      }
      let transaction = ["", "", "", data[7], data[2], data[5], data[6]];
      transactionArray.push(transaction);
    }

    transactions
      .getRange(transactions.getLastRow() + 1, 1, transactionArray.length, 7)
      .setValues(transactionArray);

    newJournal.getRange(6, 3, 2, 6).setValue("");
    newJournal.deleteRows(8, lastRow - 7);
  } else {
    ui.alert("Credit needs to equal Debit");
  }
  e.range.setValue(false);
};

/***********************************
- Change Filter
************************************/
let generateReport = (e) => {
  let accountReport = e.source.getActiveSheet();
  let lastRow = transactions.getLastRow();

  if (lastRow == 3) {
    ui.alert("There are no Transactions!");
    e.range.setValue(false);
    return false;
  }

  let filterArray = accountReport.getRange(5, 2, 1, 4).getValues();
  let transactionArray = transactions
    .getRange(4, 1, lastRow - 3, 7)
    .getValues();
  let transactionObjectsArray = [];
  console.log(filterArray);

  // Account Filter
  let accountFilter = (transaction) => {
    if (filterArray[0][0].includes("-All")) {
      let accountType = filterArray[0][0].split("-")[0];
      return transaction.accountDescription.includes(accountType);
    } else if (filterArray[0][0] != "") {
      return transaction.accountDescription == filterArray[0][0];
    } else {
      return transaction;
    }
  };
  // Invoice Filter
  let invoiceFilter = (transaction) => {
    if (filterArray[0][1] != "") {
      return transaction.invoiceNum == filterArray[0][1];
    } else {
      return transaction;
    }
  };
  // Date Start Filter
  let dateStartFilter = (transaction) => {
    if (filterArray[0][2] != "") {
      return transaction.date >= filterArray[0][2];
    } else {
      return transaction;
    }
  };
  // Date End Filter
  let dateEndFilter = (transaction) => {
    if (filterArray[0][3] != "") {
      return transaction.date <= filterArray[0][3];
    } else {
      return transaction;
    }
  };

  let transactionNum;
  let date;
  let invoiceNum;
  let memo;
  let debit;
  let credit;
  let balance;
  let accountNum;
  let accountDescription;
  let totalBalance = 0;

  for (data of transactionArray) {
    if (data[0] != "") {
      transactionNum = data[0];
      date = data[1];
      invoiceNum = data[2];
      memo = data[3];
      continue;
    } else {
      accountNum = data[3];
      accountDescription = data[4];
      debit = data[5];
      credit = data[6];
      balance = debit == "" ? Number(-credit) : Number(debit);
      let transactionObject = new Transaction(
        date,
        memo,
        invoiceNum,
        transactionNum,
        accountNum,
        accountDescription,
        debit,
        credit,
        balance
      );
      transactionObjectsArray.push(transactionObject);
    }
  }

  let totalBalanceObjects = transactionObjectsArray.filter(accountFilter);

  let sortedByDateObjects = totalBalanceObjects.sort((a, b) => a.date - b.date);

  let outputObjectArray = [];

  for (data of sortedByDateObjects) {
    totalBalance += data.balance;
    data.balance = totalBalance;
    let transactionObject = new Transaction(
      data.date,
      data.memo,
      data.invoiceNum,
      data.transactionNum,
      data.accountNum,
      data.accountDescription,
      data.debit,
      data.credit,
      data.balance
    );
    outputObjectArray.push(transactionObject);
  }

  let filteredTransactionObjects = outputObjectArray
    .filter(invoiceFilter)
    .filter(dateStartFilter)
    .filter(dateEndFilter);

  if (filteredTransactionObjects.length == 0) {
    ui.alert("There are no Transactions in this filter!");
    e.range.setValue(false);
    return false;
  }

  let outputArray = [];

  for (data of filteredTransactionObjects) {
    let transaction = [
      data.accountNum,
      data.date,
      data.memo,
      data.invoiceNum,
      data.transactionNum,
      data.debit,
      data.credit,
      data.balance,
    ];
    outputArray.push(transaction);
  }
  console.log(outputArray);

  outputArray.sort((a, b) => a[0] - b[0]);

  accountReport
    .getRange(8, 1, accountReport.getLastRow(), 8)
    .removeCheckboxes();
  accountReport.getRange(8, 1, accountReport.getLastRow(), 8).setValue("");
  accountReport.getRange(8, 1, outputArray.length, 8).setValues(outputArray);

  //Put checkboxes
  let enforceCheckbox = SpreadsheetApp.newDataValidation();
  enforceCheckbox.requireCheckbox();
  enforceCheckbox.setAllowInvalid(false);
  enforceCheckbox.build();
  accountReport
    .getRange(8, 9, outputArray.length, 1)
    .setDataValidation(enforceCheckbox);
  e.range.setValue(false);
};

/***********************************
- Create Transaction Object Constructor
************************************/
function Transaction(
  date,
  memo,
  invoiceNum,
  transactionNum,
  accountNum,
  accountDescription,
  debit,
  credit,
  balance,
  row
) {
  this.date = date;
  this.memo = memo;
  this.invoiceNum = invoiceNum;
  this.transactionNum = transactionNum;
  this.accountNum = accountNum;
  this.accountDescription = accountDescription;
  this.debit = debit;
  this.credit = credit;
  this.balance = balance;
  this.row = row;
}

/***********************************
- Categorization of Shop Expenses
************************************/
let categorization = (e) => {
  let row = e.range.getRow();
  let col = e.range.getColumn();
  let memo = expenses.getRange(row, col + 3).getValue();
  let transactionNum = newJournal.getRange(2, 4).getValue();
  let date = expenses.getRange(row, col + 1).getValue();
  let amount = expenses.getRange(row, col + 4).getValue();
  let debitAccount = e.range.getValue();
  let accountsArray = accounts.getRange("E5:F").getValues();

  let filteredArray = accountsArray.filter((data) => {
    return data[0] == debitAccount;
  });

  let debitAccountNum = filteredArray[0][1];
  let creditAccount;
  let creditAccountNum;

  switch (col) {
    case 1:
      creditAccountNum = 10000;
      creditAccount = "Cash in Hand-YGN Medical";
      break;
    case 7:
      creditAccountNum = 10001;
      creditAccount = "Cash in Hand-MDY Medical";
      break;
    case 13:
      creditAccountNum = 10002;
      creditAccount = "Cash in Hand-Main";
      break;
    case 19:
      creditAccountNum = 10003;
      creditAccount = "Cash in Hand-MDY Tools";
      break;
    case 25:
      creditAccountNum = 10004;
      creditAccount = "Cash in Hand-PMN Tools";
      break;
    case 31:
      creditAccountNum = 10005;
      creditAccount = "Cash in Hand-SINOMM";
      break;
    case 37:
      creditAccountNum = 10017;
      creditAccount = "Cash in Hand-YGN Medical ( Get Well )";
      break;
    case 43:
      creditAccountNum = 10018;
      creditAccount = "Cash in Hand-MDY Medical ( Get Well )";
      break;
  }

  let response = ui.prompt(
    "Enter Invoice Number For: " + memo,
    ui.ButtonSet.YES_NO
  );

  if (response.getSelectedButton() == ui.Button.YES) {
    let invoiceNum = response.getResponseText();
    let transactionArray = [
      [transactionNum, date, invoiceNum, memo, "", "", ""],
    ];
    let debitArray = ["", "", "", debitAccountNum, debitAccount, amount, ""];
    let creditArray = ["", "", "", creditAccountNum, creditAccount, "", amount];
    transactionArray.push(debitArray, creditArray);
    transactions
      .getRange(transactions.getLastRow() + 1, 1, transactionArray.length, 7)
      .setValues(transactionArray);
  } else {
    e.range.setValue("");
  }
};

/***********************************
- Check trial on the specific date
************************************/

let checkTrial = (e) => {
  let startDate = trial.getRange(3, 4).getValue();
  let endDate = trial.getRange(3, 5).getValue();
  let accounts = trial.getRange(5, 2, 1000, 1).getValues();
  let lastRow = transactions.getLastRow();
  let transactionArray = transactions
    .getRange(4, 1, lastRow - 3, 7)
    .getValues();
  let transactionObjectsArray = [];

  trial.getRange(2, 3).setValue("Waiting...");

  if (lastRow == 3) {
    trial.getRange(2, 3).setValue("There are no Transactions!");
    //ui.alert('There are no Transactions!');
  }

  let startDateFilter = (transaction) => {
    if (startDate != "") {
      return transaction.date >= startDate;
    } else {
      return transaction;
    }
  };

  let endDateFilter = (transaction) => {
    if (endDate != "") {
      return transaction.date <= endDate;
    } else {
      return transaction;
    }
  };

  let transactionNum;
  let date;
  let invoiceNum;
  let memo;
  let debit;
  let credit;
  let balance;
  let accountNum;
  let accountDescription;

  for (data of transactionArray) {
    if (data[0] != "") {
      transactionNum = data[0];
      date = data[1];
      invoiceNum = data[2];
      memo = data[3];
      continue;
    } else {
      accountNum = data[3];
      accountDescription = data[4];
      debit = data[5];
      credit = data[6];
      balance = debit == "" ? Number(-credit) : Number(debit);
      let transactionObject = new Transaction(
        date,
        memo,
        invoiceNum,
        transactionNum,
        accountNum,
        accountDescription,
        debit,
        credit,
        balance
      );
      transactionObjectsArray.push(transactionObject);
    }
  }

  let filteredTransactionObjects = transactionObjectsArray
    .filter(startDateFilter)
    .filter(endDateFilter);

  let trialBalance = [];

  for (data of accounts) {
    if (data[0] == "") {
      trialBalance.push(["", ""]);
      continue;
    } else {
      let filteredByAccount = filteredTransactionObjects.filter(
        (transaction) => transaction.accountNum == data[0]
      );

      if (filteredByAccount.length == 0) {
        trialBalance.push(["", ""]);
        continue;
      }

      let creditOfFilteredAccount = filteredByAccount.reduce((a, b) => {
        if (a.credit == "") {
          a.credit == 0;
        }
        if (b.credit == "") {
          b.credit == 0;
        }
        a.credit = Number(a.credit);
        b.credit = Number(b.credit);
        return { credit: a.credit + b.credit };
      });
      let debitOfFilteredAccount = filteredByAccount.reduce((a, b) => {
        if (a.debit == "") {
          a.debit == 0;
        }
        if (b.debit == "") {
          b.debit == 0;
        }
        a.debit = Number(a.debit);
        b.debit = Number(b.debit);
        return { debit: a.debit + b.debit };
      });
      trialBalance.push([
        debitOfFilteredAccount.debit,
        creditOfFilteredAccount.credit,
      ]);
    }
  }

  trial.getRange(5, 4, trialBalance.length, 2).setValues(trialBalance);
  trial.getRange(3, 6).setValue(false);
  trial.getRange(2, 3).setValue("Done!");
};

/***********************************
- Load form when editing a transaction
************************************/
let loadTransactionForm = (e) => {
  var protection = transactions.getProtections(
    SpreadsheetApp.ProtectionType.SHEET
  )[0];
  if (protection != undefined) {
    protection.remove();
  }
  let transactionNum = accountReport
    .getRange(e.range.getRow(), e.range.getColumn() - 4)
    .getValue();
  let transactionAccounts = accounts
    .getRange(5, 5, accounts.getLastRow(), 1)
    .getValues();
  let resultArray = getDataTransaction(e, transactionNum);
  const htmlServ = HtmlService.createTemplateFromFile("main");
  htmlServ.data = {
    transactionArray: resultArray,
    account: transactionAccounts,
  };
  const html = htmlServ.evaluate();
  html.setWidth(850).setHeight(600);
  ui.showModalDialog(html, "Transaction");

  e.range.setValue(false);
};

let getDataTransaction = (e, transNum) => {
  let lastRow = transactions.getLastRow();
  let transactionArray = transactions
    .getRange(4, 1, lastRow - 3, 8)
    .getValues();
  let transactionObjectsArray = [];
  let resultArray = [];

  let transactionNum;
  let date;
  let invoiceNum;
  let memo;
  let debit;
  let credit;
  let balance;
  let accountNum;
  let accountDescription;
  let totalBalance = 0;

  for (data of transactionArray) {
    if (data[0] != "") {
      transactionNum = data[0];
      date = data[1];
      invoiceNum = data[2];
      memo = data[3];
      continue;
    } else {
      accountNum = data[3];
      accountDescription = data[4];
      debit = data[5];
      credit = data[6];
      balance = debit == "" ? Number(-credit) : Number(debit);
      let transactionObject = new Transaction(
        date,
        memo,
        invoiceNum,
        transactionNum,
        accountNum,
        accountDescription,
        debit,
        credit,
        balance
      );
      transactionObjectsArray.push(transactionObject);
    }
  }

  let transactionNumFilter = (transaction) => {
    return transaction.transactionNum == transNum;
  };

  let filteredTransactionNumObjects =
    transactionObjectsArray.filter(transactionNumFilter);

  resultArray.push([
    filteredTransactionNumObjects[0].transactionNum,
    filteredTransactionNumObjects[0].date,
    filteredTransactionNumObjects[0].invoiceNum,
    filteredTransactionNumObjects[0].memo,
    "",
    "",
    "",
  ]);

  for (data of filteredTransactionNumObjects) {
    resultArray.push([
      "",
      "",
      "",
      data.accountNum,
      data.accountDescription,
      data.debit,
      data.credit,
    ]);
  }

  return resultArray;
};

let editDataTransaction = (values) => {
  let transactionNum = values[0];
  let date = values[1];
  let invoiceNum = values[2];
  let memo = values[3];

  let transactionArray = [[transactionNum, date, invoiceNum, memo, "", "", ""]];
  let j = 2;
  for (let i = 4; i < values.length - 2; i += 3) {
    let account = values[i];
    let debit = values[i + 1];
    let credit = values[i + 2];

    let transaction = [
      "",
      "",
      "",
      "=vlookup(E" +
        Number(transactions.getLastRow() + j) +
        ",'Charts Of Accounts'!E:F,2,false)",
      account,
      debit,
      credit,
    ];
    transactionArray.push(transaction);
    j++;
  }
  transactions
    .getRange(transactions.getLastRow() + 1, 1, transactionArray.length, 7)
    .setValues(transactionArray);
  deleteDataTransaction(transactionNum);
  var protection = transactions.protect();
  protection.removeEditors(protection.getEditors());
};

let deleteDataTransaction = (transactionNum) => {
  let transactionRange = transactions
    .getRange(1, 1, transactions.getLastRow(), 1)
    .getValues();
  let j = 1;
  let rowStart;
  let rowEnd;

  for (data of transactionRange) {
    if (data[0] == transactionNum) {
      rowStart = j;
      break;
    }
    j++;
  }
  for (let i = rowStart; i < transactionRange.length; i++) {
    if (transactionRange[i] != "") {
      rowEnd = i;
      break;
    }
  }

  transactions.deleteRows(rowStart, rowEnd - rowStart + 1);
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index");
}

function getUnreadEmails() {
  return GmailApp.getInboxUnreadCount();
}
