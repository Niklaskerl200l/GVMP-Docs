function Stempeluhr(Status = 0, Mitarbeiter = Umwandeln())
{
  SpreadsheetApp.getActive().getSheetByName("Log Stempeluhr").appendRow(["", Mitarbeiter, Status, new Date(), ""]);
}

function Logout_Alle()
{
  var SS_Dienstblatt = SpreadsheetApp.getActive();

  var Sheet_Stempeluhr = SS_Dienstblatt.getSheetByName("Stempeluhr");
  var Array_Streifendienst = Sheet_Stempeluhr.getRange("D3:E202").getValues();

  for(var i = 0; i < Array_Streifendienst.length; i++)
  {
    if(Array_Streifendienst[i][0] != "" && Array_Streifendienst[i][1] != "")
    {
      Einsatz_Eintragung(Array_Streifendienst[i][0], "");
    }
  }

  var Sheet_Log = SS_Dienstblatt.getSheetByName("Log Stempeluhr");
  Sheet_Log.getRange("B3:D752").clearContent();
}

function Stempeluhr_Dritte()
{
  var UI = SpreadsheetApp.getUi();

  var Name = UI.prompt("Name angeben...", "Geben Sie den Vorname-Vorname Nachname-Nachname der Person an...", UI.ButtonSet.OK).getResponseText();

  var Sheet_Import = SpreadsheetApp.getActive().getSheetByName("Import Dienstblatt LSPD");
  var Array_Import = Sheet_Import.getRange("I5:I128").getValues();

  var Gefunden = false;

  for(var i = 0; i < Array_Import.length; i++)
  {
    if(Array_Import[i][0].toString().toUpperCase().includes(Name.toString().toUpperCase()) == true)
    {
      Gefunden = true;
      break;
    }
  }

  if(Gefunden == false)
  {
    return SpreadsheetApp.getUi().alert("Fehler! Person nicht gefunden...");
  }

  if(Name != undefined && Name != null && Name != "")
  {
    SpreadsheetApp.getActive().getSheetByName("Log Stempeluhr").appendRow(["", Array_Import[i][0], 1, new Date(), ""]);
  }
}