function Telefonnummeraenderung(e) 
{
  var error = undefined;
  var Fehler = false;

  try
  {
    Direction_Telefonnummer(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    SpreadsheetApp.getActive().getSheetByName("Telefonnummeränderung").getRange(e.range.getRow(),e.range.getLastColumn() + 2).setValue(false).insertCheckboxes();
    Fehler = true;
  }
  if(error == undefined)
  {
    SpreadsheetApp.getActive().getSheetByName("Telefonnummeränderung").getRange(e.range.getRow(),e.range.getLastColumn() + 2).setValue(true).insertCheckboxes();
  }

  if(Fehler == true)
  {
    Problem;
  }
}
