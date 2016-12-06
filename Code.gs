/*******************************
 * Google Apps Script for Google Scheets that provides custom functions for the MIK
 * Metadata Mappings Helper (https://github.com/MarcusBarnes/mik/wiki/Metadata-Mappings-Helper).
 *
 * This code released to the public domain by Mark Jordan.
 */

/**
 * Adds a custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Check snippet syntax', functionName: 'getSnippet'}
  ];
  spreadsheet.addMenu('Move to Islandora Kit', menuItems);
}

/**
 * Gets the snippet and passes it off to the syntax checker function.
 */
function getSnippet() {
  var currentValue = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
  if (checkSnippetSyntax(currentValue)) {
    Browser.msgBox('Looks good!', 'Snippet ' + currentValue + ' has no syntax errors', Browser.Buttons.OK);
  }
  // If there are errors, XmlService.parse() displays its own errors,
  // so we don't need to do anything here.
  return;
}

/**
 * Checks the syntax / well-formedness of the MODS snippet.
 *
 * @param string
 *   The XML snippet from the current cell.
 */
function checkSnippetSyntax(snippet) {
  var doc;
  if (doc = XmlService.parse(snippet)) {
    return true;
  }
  else {
   return false; 
  }
}
