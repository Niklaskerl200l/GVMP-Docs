function Fraktionsfahrzeuge(e) 
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue; 

  var Sheet_Fraktionsfahrzeuge = SpreadsheetApp.getActive().getSheetByName("Fraktionsfahrzeuge");

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 8 && Value != undefined && Value != "")
  {
    Sheet_Fraktionsfahrzeuge.getRange("H" + Zeile + ":I" + Zeile).setValues([[Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy"), LSPD.Umwandeln()]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 8 && (Value == "" || Value == undefined))
  {
    Sheet_Fraktionsfahrzeuge.getRange("B" + Zeile + ":I" + Zeile).setValue("");
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 5 && Zeile <= 8 && Value == "TRUE")
  {
    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName("Logs");

    var Array_Fahrzeug = Sheet_Fraktionsfahrzeuge.getRange("B" + Zeile + ":I" + Zeile).getValues();
    var Array_Fahrzeuge = Sheet_Fraktionsfahrzeuge.getRange("B17:J").getValues();

    var Letzte_Zeile = Sheet_Log.getRange("L1").getValue();

    for(var i = 0; i < Array_Fahrzeuge.length; i ++)
    {
      if(Array_Fahrzeuge[i][2] == Array_Fahrzeug[0][2])
      {
        Sheet_Fraktionsfahrzeuge.deleteRow(i + 17);
        break;
      }
    }

    Sheet_Fraktionsfahrzeuge.insertRowAfter(16);
    Sheet_Fraktionsfahrzeuge.getRange("B17:I17").setValues(Array_Fahrzeug)

    Sheet_Log.getRange("L" + Letzte_Zeile + ":O" + Letzte_Zeile).setValues([[Array_Fahrzeug[0][0], Array_Fahrzeug[0][2], LSPD.Umwandeln(), new Date()]]);

    Sheet_Fraktionsfahrzeuge.getRange("B" + Zeile + ":J" + Zeile).setValue("");
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 17 && Value == "TRUE")
  {
    Sheet_Fraktionsfahrzeuge.deleteRow(Zeile);
  }
  else if(Spalte >= Spalte_in_Index("B") && Spalte <= Spalte_in_Index("J") && Zeile == 13)
  {
    var Array_Suche = Sheet_Fraktionsfahrzeuge.getRange("B13:J13").getValues();

    var Fraktion = Array_Suche[0][0];
    var Fahrzeugtyp = Array_Suche[0][1];
    var Seriennummer = Array_Suche[0][2];
    var Kennzeichen = Array_Suche[0][3];
    var Beweis = Array_Suche[0][4];
    var Notiz = Array_Suche[0][5];
    var Datum = Array_Suche[0][6];
    var Beamter = Array_Suche[0][7];

    try
    {
      var Filter = Sheet_Fraktionsfahrzeuge.getFilter();

      if(Fraktion != "")  Filter.setColumnFilterCriteria(2,SpreadsheetApp.newFilterCriteria().whenTextContains(Fraktion).build());
      else if(Filter.getColumnFilterCriteria(2) != null) Filter.removeColumnFilterCriteria(2);

      if(Fahrzeugtyp != "")  Filter.setColumnFilterCriteria(3,SpreadsheetApp.newFilterCriteria().whenTextContains(Fahrzeugtyp).build());
      else if(Filter.getColumnFilterCriteria(3) != null) Filter.removeColumnFilterCriteria(3);

      if(Seriennummer != "")  Filter.setColumnFilterCriteria(4,SpreadsheetApp.newFilterCriteria().whenTextContains(Seriennummer).build());
      else if(Filter.getColumnFilterCriteria(4) != null) Filter.removeColumnFilterCriteria(4);

      if(Kennzeichen != "")  Filter.setColumnFilterCriteria(5,SpreadsheetApp.newFilterCriteria().whenTextContains(Kennzeichen).build());
      else if(Filter.getColumnFilterCriteria(5) != null) Filter.removeColumnFilterCriteria(5);

      if(Beweis != "")  Filter.setColumnFilterCriteria(6,SpreadsheetApp.newFilterCriteria().whenTextContains(Beweis).build());
      else if(Filter.getColumnFilterCriteria(6) != null) Filter.removeColumnFilterCriteria(6);

      if(Notiz != "")  Filter.setColumnFilterCriteria(7,SpreadsheetApp.newFilterCriteria().whenTextContains(Notiz).build());
      else if(Filter.getColumnFilterCriteria(7) != null) Filter.removeColumnFilterCriteria(7);

      if(Datum != "")  Filter.setColumnFilterCriteria(8,SpreadsheetApp.newFilterCriteria().whenDateEqualTo(Datum).build());
      else if(Filter.getColumnFilterCriteria(8) != null) Filter.removeColumnFilterCriteria(8);
      
      if(Beamter != "")  Filter.setColumnFilterCriteria(9,SpreadsheetApp.newFilterCriteria().whenTextContains(Beamter).build());
      else if(Filter.getColumnFilterCriteria(9) != null) Filter.removeColumnFilterCriteria(9);
    }
    catch(err)
    {
      Logger.log(err.stack);
    }
  }
}
