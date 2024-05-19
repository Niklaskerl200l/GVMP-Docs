function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  LSPD.onEdit(e);

  if(Spalte == Spalte_in_Index("L") && SheetName != "Übersicht" && SheetName != "Auswertungsgedöns" && Zeile >= 9 && Zeile <= 23 && Zeile % 2 == 1 && Value == "TRUE")
  {
    Sheet.getRange("H" + Zeile + ":J" + Zeile).setValues([[new Date(),,LSPD.Umwandeln()]]);
  }
  else if(Spalte == Spalte_in_Index("L") && SheetName != "Übersicht" && SheetName != "Auswertungsgedöns" && Zeile >= 9 && Zeile <= 23 && Zeile % 2 == 1 && Value == "FALSE")
  {
    Sheet.getRange("H" + Zeile + ":J" + Zeile).setValue("");
  }

  
  else if(Spalte == Spalte_in_Index("J") && SheetName == "Übersicht" && Zeile >= 7 && Zeile <= 23 && Zeile % 2 == 1 && Value == "TRUE")
  {
    var Modul_Name = Sheet.getRange("B" + Zeile).getValue();

    try
    {
      var Sheet_Modul = SpreadsheetApp.getActive().getSheetByName(Modul_Name);

      Sheet_Modul.getName();
    }
    catch(err)
    {
      var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

      var Letzte_Zeile = Sheet_Auswertung.getRange("B1").getValue();

      var Array_Uebersetzer = Sheet_Auswertung.getRange("B3:C" + Letzte_Zeile).getValues();

      for(var i = 0; i < Array_Uebersetzer.length; i++)
      {
        if(Array_Uebersetzer[i][0] == Modul_Name)
        {
          var Sheet_Modul = SpreadsheetApp.getActive().getSheetByName(Array_Uebersetzer[i][1]);

          var Sonder_Modul = Modul_Name;

          break;
        }
      }
    }

    var Array_Module = Sheet_Modul.getRange("B9:B" + Sheet_Modul.getLastRow()).getValues();

    var Name = LSPD.Umwandeln();

    for(var i = 0; i < Array_Module.length; i++)
    {
      if(Sonder_Modul != undefined)
      {
        if(Array_Module[i][0] != "" && Sonder_Modul == Array_Module[i][0])
        {
          Sheet_Modul.getRange("H" + (9 + i) + ":L" + (9 + i)).setValues([[new Date(),,Name,,true]]);
        }
      }
      else
      {
        if(Array_Module[i][0] != "")
        {
          Sheet_Modul.getRange("H" + (9 + i) + ":L" + (9 + i)).setValues([[new Date(),,Name,,true]]);
        }
      }
    }
  }


  else if(Spalte == Spalte_in_Index("J") && SheetName == "Übersicht" && Zeile >= 7 && Zeile <= 23 && Zeile % 2 == 1 && Value == "FALSE")
  {
    var Modul_Name = Sheet.getRange("B" + Zeile).getValue();

    try
    {
      var Sheet_Modul = SpreadsheetApp.getActive().getSheetByName(Modul_Name);

      Sheet_Modul.getName();
    }
    catch(err)
    {
      var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

      var Letzte_Zeile = Sheet_Auswertung.getRange("B1").getValue();

      var Array_Uebersetzer = Sheet_Auswertung.getRange("B3:C" + Letzte_Zeile).getValues();

      for(var i = 0; i < Array_Uebersetzer.length; i++)
      {
        if(Array_Uebersetzer[i][0] == Modul_Name)
        {
          var Sheet_Modul = SpreadsheetApp.getActive().getSheetByName(Array_Uebersetzer[i][1]);

          var Sonder_Modul = Modul_Name;

          break;
        }
      }
    }

    var Array_Module = Sheet_Modul.getRange("B9:B" + Sheet_Modul.getLastRow()).getValues();

    var Name = LSPD.Umwandeln();

    for(var i = 0; i < Array_Module.length; i++)
    {
      if(Sonder_Modul != undefined)
      {
        if(Array_Module[i][0] != "" && Sonder_Modul == Array_Module[i][0])
        {
          Sheet_Modul.getRange("H" + (9 + i) + ":L" + (9 + i)).setValue("");
        }
      }
      else
      {
        if(Array_Module[i][0] != "")
        {
          Sheet_Modul.getRange("H" + (9 + i) + ":L" + (9 + i)).setValue("");
        }
      }
    }
  }
}

