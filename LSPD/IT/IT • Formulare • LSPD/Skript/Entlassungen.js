function Entlassung(e)
{
  var error = undefined;
  var Fehler = false;

  try
  {
    Direction_Entlassung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 1).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 1).setValue(true).insertCheckboxes();
  }

  error = undefined;

  try
  {
    Training_Entlassung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 2).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 2).setValue(true).insertCheckboxes();
  }

  error = undefined;
  
  try
  {
    Zeitsystem_Entlassung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 3).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 3).setValue(true).insertCheckboxes();
  }

  try
  {
    Loyal_Entlassung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 4).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("LSPD Entlassungen").getRange(e.range.getRow(),e.range.getLastColumn() + 4).setValue(true).insertCheckboxes();
  }

  if(Fehler == true)
  {
    Problem;
  }
}
