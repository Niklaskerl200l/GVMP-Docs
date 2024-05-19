function onOpen()
{
  LSPD.onOpen();
  
  SpreadsheetApp.getUi().createMenu("Start Import").addItem("Start","Start_Manuell").addToUi()
}

function Start_Manuell()
{
  Uebertragung(SpreadsheetApp.getActive().getActiveSheet().getName());
}