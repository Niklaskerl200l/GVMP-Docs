function onOpen(e)
{
  LSPD.onOpen();
  
  try
  {
    if(SpreadsheetApp.getActiveSpreadsheet().getName() != "Ausbildungen Vorlage")
    {
      SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns").getRange("E3").setValue(ScriptApp.getScriptId());
    }
  }
  catch(err)
  {
    Logger.log(err);
  }
}