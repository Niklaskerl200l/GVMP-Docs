function Einsatz_Dokumentation(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("H") && Zeile >= 5 && Value == "TRUE")
  {
    var Sheet_Einsatz = SpreadsheetApp.getActive().getSheetByName("Einsatz Archiv");

    var Array_Einsatz = Sheet_Einsatz.getRange("B5:F5").getValues();
    var Array_Letzter_Einsatz = Sheet_Einsatz.getRange("B9:H9").getValues();

    var Cooldown = new Date();

    Cooldown.setMinutes(Cooldown.getMinutes() - 30);

    if(Array_Letzter_Einsatz[0][0] == Array_Einsatz[0][0] && Array_Letzter_Einsatz[0][6] == Array_Einsatz[0][6] && Array_Letzter_Einsatz[0][5] >= Cooldown)
    {
      Sheet_Einsatz.getRange("B5:H5").setValue("");
      SpreadsheetApp.getUi().alert("Dieser Einsatz wurde schon Eingetragen");
    }
    else
    {
      Sheet_Einsatz.insertRowAfter(8);

      Sheet_Einsatz.getRange("B9:I9").setValues([[Array_Einsatz[0][0], Array_Einsatz[0][1],Array_Einsatz[0][2],Array_Einsatz[0][3],Array_Einsatz[0][4], new Date(), LSPD.Umwandeln(), Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "MM")]]);

      Sheet_Einsatz.getRange("B5:H5").setValue("");
    }
  }
}

function Einsatz_Eintragung(Benutzer, Einsatz = "")
{
  var Sheet_Einsatz = SpreadsheetApp.getActive().getSheetByName("Log Einsatz");

  if(Einsatz == "Abteilungsarbeit" || Einsatz == "Unmarked")
  {
    var Abteilungen = LSPD.Propertie_Lesen("LSPD_Abteilung").toString();
    
    if((Abteilungen.includes("Detective") == false && Abteilungen.includes("Direction") == false) && Einsatz == "Unmarked")
    {
      return SpreadsheetApp.getUi().alert("Fehler!\nKeinen Zugriff!");
    }

    if((Abteilungen.includes("Gang Task Force") == false && Abteilungen.includes("Direction") == false) && Einsatz == "Abteilungsarbeit")
    {
      return SpreadsheetApp.getUi().alert("Fehler!\nKeinen Zugriff!");
    }
  }

  Sheet_Einsatz.appendRow(["", Benutzer, Einsatz, new Date(), ""]);
}

function Einsatz_Beenden(Einsatz)
{
  if(Einsatz == "")
  {
    return 0;
  }

  var Sheet_Stempeluhr = SpreadsheetApp.getActive().getSheetByName("Stempeluhr");

  var Array_Einsatz =  Sheet_Stempeluhr.getRange("D3:E").getValues();

  for(var y = 0; y < Array_Einsatz.length; y++)
  {
    if(Array_Einsatz[y][1] == Einsatz)
    {
      Einsatz_Eintragung(Array_Einsatz[y][0],"");
    }
  }
}

function Einsatz_Archivieren()
{
  var Sheet_Log_Einsatz = SpreadsheetApp.getActive().getSheetByName("Log Einsatz");
  var Sheet_Archiv_Einsatz = SpreadsheetApp.openById(LSPD.ID_Archiv_Einsatz_Logs).getSheetByName("Archiv Einsatz")

  var Letzte_Zeile = Sheet_Log_Einsatz.getLastRow();

  if(Letzte_Zeile <= 2)
  {
    Letzte_Zeile = 3;
  }

  var Array_Log = Sheet_Log_Einsatz.getRange("B3:D" + Letzte_Zeile).getValues();

  Sheet_Log_Einsatz.getRange("B3:D" + Letzte_Zeile).setValue("");

  Sheet_Archiv_Einsatz.getRange(Sheet_Archiv_Einsatz.getLastRow() + 1,2,Array_Log.length,Array_Log[0].length).setValues(Array_Log);
}

function Einsatz_Auto_Archivieren()
{
  var Dienstblatt = SpreadsheetApp.getActive().getSheetByName("Einsatz Archiv");
  var Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_Einsaetze).getSheetByName("Archiv");

  var Daten = Dienstblatt.getRange("B9:H").getValues();

  var NeunzigTage = new Date();
  NeunzigTage.setDate(NeunzigTage.getDate() - 30);

  var Archiv_Array = [];

  for(var i = Daten.length - 1; i >= 0; i--)
  {
    if(Daten[i][0] != "" && new Date(Daten[i][5]) <= NeunzigTage)
    {
      Logger.log(i + 9);
      Archiv_Array.push(Daten[i]);

      Dienstblatt.deleteRow(i + 9);
    }
  }

  if(Archiv_Array.length > 0)
  {
    Archiv.insertRowsAfter(3, Archiv_Array.length);
    Archiv.getRange("B4:H" + (Archiv_Array.length + 3)).setValues(Archiv_Array);
  }
}