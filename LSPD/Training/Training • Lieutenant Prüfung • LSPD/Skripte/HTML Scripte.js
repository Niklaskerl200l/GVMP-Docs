function LT_Start()  //Starten des Benutzerinterface für LT Eingaben
{
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('LT UI').setWidth(1114).setHeight(200).setSandboxMode(HtmlService.SandboxMode.IFRAME),"Prüfungs Daten eingeben");
}

function RCT_Alle()   //Ausgabe aller RCT für LT Benutzerinterface
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Import Personaltabelle");
  var Array = Sheet.getRange("D4:D" + Sheet.getRange("D1").getValue()).getValues();
  return Array;
}

function Pruefer_Alle()   //Ausgabe aller Prüfer für LT Benutzerinterface
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Import Personaltabelle");
  var Array = Sheet.getRange("D4:D" + Sheet.getRange("D1").getValue()).getValues()
  return Array;
}

function Email()    //Ausgabe des Namen des Aktiven Benutzers für LT Benutzerinterface
{
  return LSPD.Umwandeln();
}
