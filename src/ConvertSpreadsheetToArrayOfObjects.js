/**
 * Converts a spreadsheet to an array of objects in Google Apps Script
 * Lightly modified version of code taken from here - https://sites.google.com/site/scriptsexamples/custom-methods/gs-objects/source
 * 
 * @param {class} spreadsheet
 * @returns {array}
 */
function convertSpreadSheetToObjectArray( sheet, allowDigitAsFirstCharacter = false ){
  const values = sheet.getRange( 1, 1, sheet.getLastRow(), sheet.getLastColumn() ).getValues();
  const headers = values[ 0 ];
  
  const objects = [];
  for ( var i = 1; i < values.length; i++ ) {
    var row = {};
    
    row.rowNum = i;
    
    for ( var j = 0; j < headers.length; j++ ) {
      row[ camelString_( headers[ j ], allowDigitAsFirstCharacter ) ] = values[ i ][ j ];
    }
    
    objects.push( row );
  }
  
  return objects;
}

function camelString_( header, allowDigitAsFirstCharacter = false ) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter) && '$' !== letter) {
      continue;
    }
    if ( key.length == 0 && isDigit_(letter) && ! allowDigitAsFirstCharacter ) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

function isDigit_(char) {
  return char >= '0' && char <= '9';
}
