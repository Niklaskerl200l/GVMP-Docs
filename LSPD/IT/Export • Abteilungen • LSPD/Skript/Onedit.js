function onEdit(e) 
{
  LSPD.onEdit(e);
  
  SpreadsheetApp.getActive().getSheetByName("Import Abteilungen").getRange("A3").setValue("WAHR");
}

LSPD.Eingabe_Test();

function onOpen()
{
  LSPD.onOpen();
}