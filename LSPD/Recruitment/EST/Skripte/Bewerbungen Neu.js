function Bewerbungen_Neu(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile >= 5 && Zeile <= 35 && Spalte == Spalte_in_Index("C") && Value == "Evaluierung")
  {
    var Sheet_Bewerbungen_Neu = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Neu");

    Sheet_Bewerbungen_Neu.getRange("K" + Zeile).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM") + " " + LSPD.Umwandeln());
  }
  else if(Zeile >= 5 && Zeile <= 35 && Spalte == Spalte_in_Index("C") && Value == "Eingeladen")
  {
    var Sheet_Bewerbungen_Neu = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Neu");

    Sheet_Bewerbungen_Neu.getRange("L" + Zeile).setValue(Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM") + " " + LSPD.Umwandeln());
  }
  else if(Zeile >= 5 && Zeile <= 35 && Spalte == Spalte_in_Index("N") && Value == "TRUE")
  {
    var Sheet_Bewerbungen_Neu = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Neu");
    var Sheet_Bewerbungen_Aktuell = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Aktuell");
    var Sheet_Bewerbungen_Archiv = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Archiv");

    var Array_Neu = Sheet_Bewerbungen_Neu.getRange("B" + Zeile + ":L" + Zeile).getValues();

    if(Array_Neu[0][1] == "Abgelehnt")
    {
      Sheet_Bewerbungen_Neu.getRange(Zeile, Spalte).setValue(false);
      Array_Neu = Array_Neu[0];

      var Array_Archiv = 
      [[
        Array_Neu[0],
        "Abgelehnt",
        "-",
        "-",
        Array_Neu[2],
        Array_Neu[3],
        Array_Neu[4],
        Array_Neu[5],
        Array_Neu[6],
        "-",
        Array_Neu[7],
        Array_Neu[8],
        Array_Neu[9],
        Array_Neu[10],
        "-"
      ]];

      Sheet_Bewerbungen_Archiv.insertRowAfter(5);
      Sheet_Bewerbungen_Archiv.getRange("B6:P6").setValues(Array_Archiv);

      Sheet_Bewerbungen_Neu.getRange("B" + Zeile + ":N" + Zeile).clearContent();
      Sheet_Bewerbungen_Neu.getRange("B5:N32").sort({column: Spalte_in_Index("B"), ascending: true});

      Sheet_Bewerbungen_Archiv.setActiveSelection("B6");
    }
    else
    {
      var Array_Aktuell =
      [[
        Array_Neu[0][0],
        "Bewerbungsgespräch",
        "",
        LSPD.Umwandeln(),
        Array_Neu[0][2],
        Array_Neu[0][3],
        Array_Neu[0][4],
        Array_Neu[0][5],
        Array_Neu[0][6],
        Array_Neu[0][7],
        Array_Neu[0][8],
        Array_Neu[0][9],
        Array_Neu[0][10]
      ]];

      Logger.log(Array_Aktuell)

      Sheet_Bewerbungen_Neu.getRange("B" + Zeile + ":L" + Zeile).setValue("");
      Sheet_Bewerbungen_Neu.getRange("N" + Zeile).setValue("");
      Sheet_Bewerbungen_Aktuell.getRange("B5:O35").sort(4);
      
      var Letze_Zeile = Sheet_Bewerbungen_Aktuell.getRange("B3").getValue();
    
      Logger.log(Array_Neu);

      Sheet_Bewerbungen_Aktuell.getRange("B" + Letze_Zeile + ":N" + Letze_Zeile).setValues(Array_Aktuell);
      Sheet_Bewerbungen_Aktuell.getRange("O" + Letze_Zeile).insertCheckboxes();

      Sheet_Bewerbungen_Aktuell.setActiveSelection("D" + Letze_Zeile);
    }

    Sheet_Bewerbungen_Neu.getRange("B5:L32").sort(2);
  }
}