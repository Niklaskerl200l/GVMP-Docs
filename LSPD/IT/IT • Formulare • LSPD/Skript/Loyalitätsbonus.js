ID_Loyal = LSPD.ID_Loyalitätsbonus;

function Loyal_Einstellung(e)
{
  var Sheet_Loyal = SpreadsheetApp.openById(ID_Loyal).getSheetByName("Auszahlungstabelle LSPD");

  var Werte = e.namedValues;
  var Letze_Zeile = Sheet_Loyal.getLastRow() + 1;

  Sheet_Loyal.insertRowAfter(Letze_Zeile - 1);

  Logger.log(Werte);
  
  var Beitritt = Werte.Beitritt.toString();

  var Jahr = Beitritt.substring(Beitritt.indexOf(".", Beitritt.indexOf(".")+1)+1,Beitritt.length)
  var Monat = Beitritt.substring(Beitritt.indexOf(".") + 1,Beitritt.indexOf(".", Beitritt.indexOf(".")+1))
  var Tag = Beitritt.substring(0,Beitritt.indexOf("."));

  var Datum = new Date(Jahr,Monat,Tag);

  Logger.log(Datum);

  Datum.setMonth(Datum.getMonth() + 2);

  Logger.log(Datum);

  Sheet_Loyal.getRange("B"+Letze_Zeile + ":F" + Letze_Zeile).setValues([[Werte.Name,Werte.Beitritt,Werte["Forum ID"],,Datum]]);

  Sheet_Loyal.getRange("E" + Letze_Zeile).setFormula("=DATEDIF(C"+Letze_Zeile+";TODAY();\"m\")");
  Sheet_Loyal.getRange("G" + Letze_Zeile).setFormula("=VLOOKUP(E"+Letze_Zeile+";'Übersicht'!$B$5:$C$25;2)");
}

function Loyal_Entlassung(e)
{
  var Sheet_Loyal = SpreadsheetApp.openById(ID_Loyal).getSheetByName("Auszahlungstabelle LSPD");

  var Werte = e.namedValues;

  var Array_Personen = Sheet_Loyal.getRange("B5:M").getValues();

  for(var y = 0; y < Array_Personen.length; y++)
  {
    if(Array_Personen[y][0] == Werte.Name)
    {
      var Zeile = y + 5;

      if(Array_Personen[y][6] == true)
      {
        var Letze_Zeile = Sheet_Loyal.getRange("Q27").getValue();

        Sheet_Loyal.getRange("Q" + Letze_Zeile + ":S" + Letze_Zeile).setValues([[Array_Personen[y][0],Array_Personen[y][4],Array_Personen[y][5]]]);
        Sheet_Loyal.getRange("T" + Letze_Zeile).insertCheckboxes();
      }

      Sheet_Loyal.getRange("H" + Zeile + ":L" + Zeile).removeCheckboxes();
      Sheet_Loyal.getRange("B" + Zeile + ":M" + Zeile).setValue("");

      break;
    }
  }

  Sheet_Loyal.getRange("B5:M").sort({column: Spalte_in_Index("F"),ascending: true});
}