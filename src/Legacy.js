const cashSheet = getSheetByName( 'Cash' );

function ConvertCreditCardSheetsToCashBalances() {
  convertCapitalOneCreditCardsToCashBalances();
  convertChaseCreditCardsToCashBalances();
}

function convertChaseCreditCardsToCashBalances() {
  convertChaseCreditCardToCashBalances( 'Sapphire Reserve' );
}

function convertCapitalOneCreditCardsToCashBalances() {
  // This assumes that the name of the column name in 'Cash' is the same as the sheet name for the CSV
  convertCapitalOneCsvToCashBalances( 'Venture' );
  convertCapitalOneCsvToCashBalances( 'Spark' );
}

function convertChaseCreditCardToCashBalances( chaseCardName ) {
  const transactionsSummaryByDate = getTransactionsSummaryByPostedDate( chaseCardName );

  Logger.log( 'Transactions Summary by date', transactionsSummaryByDate );
}

function convertCapitalOneCsvToCashBalances( capitalOneCardName ) {
  var transactionsSummaryByDate = getTransactionsSummaryByPostedDate( capitalOneCardName );
  var balancesByDate = getBalancesByDate( transactionsSummaryByDate, capitalOneCardName );

  if ( null !== balancesByDate ) {
    insertBalancesIntoCashSheet( balancesByDate, capitalOneCardName );
  }
}

function getSheetByName( sheetName ) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName( sheetName );
}

function getCurrentTransactionDebitOrCredit( transaction, debitIndex, creditIndex ) {
  var transactionValue = transaction[ creditIndex ];
  
  if ( '' === transactionValue ) {
    transactionValue = transaction[ debitIndex ] * -1;
  }
  
  return transactionValue;
}

function roundValueToSecondDecimal( value ) {
  return Math.round( ( value + Number.EPSILON ) * 100 ) / 100;
}

function getTransactionsSummaryByPostedDate( sheetName ) {
  var csvSheet = getSheetByName( sheetName );
  var csvTransactions = csvSheet.getDataRange().getValues();
  
  var debitIndex = getColumnIndexByName( csvTransactions, 'Debit' );
  var creditIndex = getColumnIndexByName( csvTransactions, 'Credit' );
  var postedDateIndex = getColumnIndexByName( csvTransactions, 'Posted Date' );
  
  var transactionsSummary = [];
  var transactionsDateSubtotal = 0;
  for ( var transaction in csvTransactions ) {
    
    if ( 0 != transaction ) { // Skip header row
    
      var currentTransaction = csvTransactions[ transaction ];
      var currentTransactionValue = getCurrentTransactionDebitOrCredit( currentTransaction, debitIndex, creditIndex );
      var currentPostedDate = convertDateToIso8601( currentTransaction[ postedDateIndex ] );
      
      if ( ! transactionsSummary[ currentPostedDate ] ) {
        transactionsSummary[ currentPostedDate ] = currentTransactionValue;
      } else {
        transactionsSummary[ currentPostedDate ] += currentTransactionValue;
      }
      
      transactionsSummary[ currentPostedDate ] = roundValueToSecondDecimal( transactionsSummary[ currentPostedDate ] );
      
    }
  }
  
  return transactionsSummary;
}

// function getColumnIndexByName( sheetValues, columnName ) {
//   return sheetValues[ 0 ].indexOf( columnName ); // Assumes headers are in first row '0' - an optional parameter could be added for header row if this is not the case
// }

function getNextRowInSingleColumnInSheet( column, sheetName ) {
  var sheet = getSheetByName( sheetName );
  var lastAbsoluteRow = sheet.getMaxRows();
  var sheetValues = sheet.getRange( 1, column, lastAbsoluteRow ).getValues();
  
  var nextRowInSingleColumn = null;
  for ( var row = lastAbsoluteRow; "" == sheetValues[ row - 1 ] && row > 0; row-- ) {
    nextRowInSingleColumn = row;
  }
  
  return nextRowInSingleColumn;
}

function getBalancesByDate( transactionsSummary, cardName ) {
  var cashSheetValues = cashSheet.getDataRange().getValues();
  
  var cardNameColumnIndex = getColumnIndexByName( cashSheetValues, cardName );
  var dateColumnIndex = getColumnIndexByName( cashSheetValues, 'Date' );
  var nextRowInCardColumn = getNextRowInSingleColumnInSheet( ( cardNameColumnIndex + 1 ), 'Cash' );
  var nextRowInDateColumn = getNextRowInSingleColumnInSheet( ( dateColumnIndex + 1 ), 'Cash' );
  
  var firstDate = cashSheet.getRange( nextRowInCardColumn, ( dateColumnIndex + 1 ) ).getValues()[ 0 ][ 0 ];

  if ( '' !== firstDate ) {
    firstDate = convertDateToIso8601( firstDate );
  } else {
    return null; // No transactions
  }
  
  var lastDate = cashSheet.getRange( ( nextRowInDateColumn - 1 ), ( dateColumnIndex + 1 ) ).getValues()[ 0 ][ 0 ];
  lastDate = convertDateToIso8601( lastDate );
  
  var previousBalance = cashSheet.getRange( ( nextRowInCardColumn - 1 ), ( cardNameColumnIndex + 1 ) ).getValues()[ 0 ][ 0 ];
  
  var balancesByDate = [];
  for ( var row = nextRowInCardColumn; row < nextRowInDateColumn; row++ ) {
    var date = cashSheet.getRange( row, ( dateColumnIndex + 1 ) ).getValues()[ 0 ][ 0 ];
    date = convertDateToIso8601( date );
    var transactionSummary = ( date in transactionsSummary ) ? ( transactionsSummary[ date ] ) : ( 0 );
    var balance = previousBalance + transactionSummary;
    balance = roundValueToSecondDecimal( balance );
    
    balancesByDate[ date ] = balance;
    
    previousBalance = balance;
  }
  
  return balancesByDate;
}

function convertBalancesToValueSettingFormat( balances ) {
  var convertedBalances = [];
  
  var i = 0;
  for ( var balance in balances ) {
    convertedBalances[ i ] = [ balances[ balance ] ];
    i++;
  }
  
  return convertedBalances;
}

function insertBalancesIntoCashSheet( balances, columnName ) {
  var cashSheet = getSheetByName( 'Cash' );
  var cashSheetValues = cashSheet.getDataRange().getValues();
  var cardNameColumnIndex = getColumnIndexByName( cashSheetValues, columnName );
  var dateColumnIndex = getColumnIndexByName( cashSheetValues, 'Date' );
  var nextRowInCardColumn = getNextRowInSingleColumnInSheet( ( cardNameColumnIndex + 1 ), 'Cash' );
  var nextRowInDateColumn = getNextRowInSingleColumnInSheet( ( dateColumnIndex + 1 ), 'Cash' );
  var numberRowsFromLastCardValueToLastDate = nextRowInDateColumn - nextRowInCardColumn;
  
  var cardRange = cashSheet.getRange( nextRowInCardColumn, ( cardNameColumnIndex + 1 ), numberRowsFromLastCardValueToLastDate );
  var balances = convertBalancesToValueSettingFormat( balances );
  
  cardRange.setValues( balances );
}

function FillInSelectedCashBlanks() {
  
  // Declare standard vars
  var ss = SpreadsheetApp;
  var actSs = ss.getActiveSpreadsheet();
  var actSht = actSs.getActiveSheet();
  
  // Get selected range
  var transData = actSht.getActiveRange().getValues();
  
  // Run through selected range and fill in blank data from the row before it
  for ( var account in transData ) {
    for ( var date in transData ) {
      if ( '' === transData[date][account] ) {
        transData[date][account] = transData[date - 1][account]; // Make date value equal to the day before if blank
      }
    }
  }
  
  // Set processed data in selected range
  actSht.getActiveRange().setValues( transData );
  
}

































