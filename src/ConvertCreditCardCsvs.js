const creditCardCsvSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Credit Card CSV');
const creditCardBalancesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Credit Cards');

function ConvertCreditCardCsv() {
  const selectCreditCardHtml = HtmlService.createHtmlOutputFromFile('helpers/CreditCardSelect');
  SpreadsheetApp.getUi().showModalDialog(selectCreditCardHtml, 'Select Credit Card to Convert');
}

function submitCreditCardAndConvert(formObject) {
  const creditCard = formObject['credit-card'];

  if ('Venture' === creditCard) {
    convertCapitalOneCsvToBalances('Venture');
  } else if ('Spark' === creditCard) {
    convertCapitalOneCsvToBalances('Spark');
  } else if ('Sapphire Reserve' === creditCard) {
    convertChaseCsvToBalances('Sapphire Reserve');
  } else if ('Freedom Unlimited' === creditCard) {
    convertChaseCsvToBalances('Freedom Unlimited');
  }
}

function convertChaseCsvToBalances(cardName) {
  const transactions = getTransactionsSortedByDate('postDate');
  const creditCardBalancesByDate = getCreditCardBalancesByDate();

  const camelizedCardName = camelize(cardName);

  const preUpdateBalance = getLastBalanceOfCreditCardCsvColumn(creditCardBalancesByDate, camelizedCardName);
  const newBalancesByDate = getNewBalancesByDate(preUpdateBalance, transactions, 'postDate');

  const startBalance = creditCardBalancesByDate.find(balance => '' === balance[camelizedCardName]);
  const startDate = startBalance.date;
  const startRowIndex = startBalance.rowNum + 1;

  insertBalancesIntoCreditCardSheet(newBalancesByDate, cardName, startDate, preUpdateBalance, startRowIndex);
}

// https://stackoverflow.com/questions/2970525/converting-any-string-into-camel-case
function camelize(string) {
  return string.replace(
    /(?:^\w|[A-Z]|\b\w)/g,
    function(word, index) {
      return index === 0 ? word.toLowerCase() : word.toUpperCase();
    }
  ).replace(/\s+/g, '');
}

function convertCapitalOneCsvToBalances(cardName) {
  const transactions = getTransactionsSortedByDate('postedDate');
  const creditCardBalancesByDate = getCreditCardBalancesByDate();

  const preUpdateBalance = getLastBalanceOfCreditCardCsvColumn(creditCardBalancesByDate, cardName);
  const newBalancesByDate = getNewBalancesByDate(
    preUpdateBalance,
    transactions,
    'postedDate',
    'credit',
    'debit'
  );

  const startBalance = creditCardBalancesByDate.find(balance => '' === balance[cardName.toLowerCase()]);
  const startDate = startBalance.date;
  const startRowIndex = startBalance.rowNum + 1;

  insertBalancesIntoCreditCardSheet(newBalancesByDate, cardName, startDate, preUpdateBalance, startRowIndex);
}

function getCreditCardBalancesByDate() {
  let balances = convertSpreadSheetToObjectArray(creditCardBalancesSheet);
  return convertDatesToIso8601(balances, ['date']);
}

function getTransactionsSortedByDate(postedDateColumnName) {
  let transactions = convertSpreadSheetToObjectArray(creditCardCsvSpreadsheet);
  transactions = convertDatesToIso8601(transactions, ['transactionDate', postedDateColumnName]);
  transactions.sort(sortTransactionsByProperty(postedDateColumnName));

  return transactions;
}

function insertBalancesIntoCreditCardSheet(
  balances,
  columnName,
  startDate,
  preUpdateBalance,
  firstNewDateRow
) {
  const creditCardColumnIndex = getColumnIndexByName(columnName, creditCardBalancesSheet);

  balances = fillOutBlankDateValues(balances, startDate, preUpdateBalance);
  const dateCount = Object.keys(balances).length;

  const balancesRange = creditCardBalancesSheet.getRange(firstNewDateRow, creditCardColumnIndex, dateCount);

  const balanceValues = Object.values(balances);
  const balanceSpreadsheetValues = []
  for (const balanceValue of balanceValues) {
    balanceSpreadsheetValues.push([balanceValue]);
  }

  balancesRange.setValues(balanceSpreadsheetValues);
}

function fillOutBlankDateValues(balances, startDateIso8601, preUpdateBalance) {
  // All dates are in UTC, even if Spreadsheet is in different time zone

  const startDate = new Date(startDateIso8601);
  const yesterday = getYesterday();

  if (!(startDateIso8601 in balances)) {
    balances[startDateIso8601] = preUpdateBalance;
  }

  let previousBalance = balances[startDate];

  const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;

  for (let day = startDate; day <= yesterday; day = new Date(day.getTime() + MILLIS_PER_DAY)) {
    const dayIso8601 = convertDateToIso8601(day);

    if (!(dayIso8601 in balances)) {
      balances[dayIso8601] = previousBalance;
    }

    previousBalance = balances[dayIso8601];
  }

  return Object.keys(balances).sort().reduce(
    (obj, key) => { 
      obj[key] = balances[key];
      return obj;
    }, 
    {}
  );
}

function getColumnIndexByName(name, sheet) {
  const data = sheet.getDataRange().getValues();
  const columnIndex = data[ 0 ].indexOf( name );
  
  return columnIndex !== -1 ? columnIndex + 1 : null;
}

function convertDatesToIso8601(objects, dateAttrNames = []) {
  for (const object of objects) {
    for (const dateAttrName of dateAttrNames) {
      const date = object[dateAttrName];

      if ('object' === typeof date) { // Date objects only
        object[dateAttrName] = convertDateToIso8601(date);
      }
    }
  }

  return objects;
}

function getLastBalanceOfCreditCardCsvColumn(creditCardBalancesByDate, columnName) {
  columnName = columnName.charAt(0).toLowerCase() + columnName.slice(1); // Change first letter lowercase
  const dayBalance = creditCardBalancesByDate.slice().reverse().find(balance => '' !== balance[columnName]);
  const cardBalance = dayBalance[columnName];

  return roundCurrency(cardBalance);
}

function getNewBalancesByDate(
  previousBalance,
  transactions,
  postedDateName,
  creditName = false,
  debitName = false
) {
  const balances = {};

  for (const [transactionIndex, transaction] of Object.entries(transactions)) {
    const postedDate = transaction[postedDateName];
    const transactionAmount = getTransactionAmount(creditName, debitName, transaction); // If positive, is credit, if negative, is debit

    if (0 === parseInt(transactionIndex) || !dateExistsInBalances(postedDate, balances)) {
      const newBalance = previousBalance + transactionAmount;
      balances[postedDate] = roundCurrency(newBalance);
    } else {
      balances[postedDate] += transactionAmount;
    }

    previousBalance = balances[postedDate];
  }

  return balances;
}

function roundCurrency(amount) {
  return Math.round(amount * 100) / 100;
}

function dateExistsInBalances(date, balances) {
  return balances.hasOwnProperty(date);
}

function getTransactionAmount(creditName, debitName, transaction) {
  let transactionAmount;
  if (creditName && debitName) {
    transactionAmount = getTransactionAmountFromCreditOrDebit(transaction.credit, transaction.debit);
  } else {
    transactionAmount = transaction.amount;
  }

  return parseFloat(transactionAmount);
}

function getTransactionAmountFromCreditOrDebit(credit, debit) {
  var amount = 0;

  if ('' === debit && parseFloat(credit)) {
    amount += credit;
  } else if ('' === credit && parseFloat(debit)) {
    amount += -(debit);
  }

  return amount.toFixed(2);
}

function sortTransactionsByProperty(property) {
  return function (a, b) {
    const propertyA = a[property];
    const propertyB = b[property];

    let comparison = 0;

    if (propertyA > propertyB) {
      comparison = 1;
    } else if (propertyA < propertyB) {
      comparison = -1;
    }

    return comparison;
  }
}

// function sortTransactionsByProperty(a, b) {
  // const postedDateA = a.postedDate;
  // const postedDateB = b.postedDate;

  // let comparison = 0;

  // if (postedDateA > postedDateB) {
  //   comparison = 1;
  // } else if (postedDateA < postedDateB) {
  //   comparison = -1;
  // }

  // return comparison;
// }





































