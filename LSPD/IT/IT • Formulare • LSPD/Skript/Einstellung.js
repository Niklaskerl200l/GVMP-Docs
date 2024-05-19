function Einstellung(e)
{
  var error = undefined;
  var Fehler = false;

  try
  {
    Direction_Einstellung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 1).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 1).setValue(true).insertCheckboxes();
  }

  error = undefined;

  try
  {
    Training_Einstellung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 2).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 2).setValue(true).insertCheckboxes();
  }

  error = undefined;
  
  try
  {
    Zeitsystem_Einstellung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 3).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 3).setValue(true).insertCheckboxes();
  }

  try
  {
    Loyal_Einstellung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 4).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Einstellung").getRange(e.range.getRow(),e.range.getLastColumn() + 4).setValue(true).insertCheckboxes();
  }

  MailApp.sendEmail(e.namedValues.Email.toString(),"Einladung LSPD Dienstblatt","Dienstblatt - LSPD: https://docs.google.com/spreadsheets/d/" + LSPD.ID_Dienstblatt + " \n\n In bis zu 10 Minuten dürftest du Zugriff erhalten.\nAls Tipp stellt man den Browser Zoom bei Google Docs auf 90% dann siehst man das Dienstblatt auf einem Screen.");

  if(Fehler == true)
  {
    Problem;
  }
}