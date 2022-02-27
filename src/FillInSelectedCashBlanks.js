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