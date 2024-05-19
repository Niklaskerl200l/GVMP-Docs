function Minus_Leitstelle(e)
{
  var Sheet_LST_Minus = SpreadsheetApp.getActive().getSheetByName("Minus Leitstelle");
  var Spalte = e.range.getColumn();
  var Zeile = e.range.getRow();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("I") && Zeile == 3 && Value == "TRUE")
  {
    Sheet_LST_Minus.getRange(Zeile, Spalte).setValue(false);

    var Array_Eintrag = Sheet_LST_Minus.getRange("B" + Zeile + ":H" + Zeile).getValues();
    Array_Eintrag = Array_Eintrag[0];

    Sheet_LST_Minus.insertRowAfter(6);
    Sheet_LST_Minus.getRange("B7:H7").setValues([Array_Eintrag]);
    Sheet_LST_Minus.getRange("I7").setValue(false);

    Sheet_LST_Minus.getRange("B" + Zeile + ":H" + Zeile).clearContent();
  }
}

function LST_Minus_Zeitsystem()
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Minus Leitstelle");
  var Daten = Sheet.getRange("B7:I").getValues();

  if(Daten.filter(function(e){return e[0] != "" && e[7] == false}).length == 0) return;

  var Zeitsystem = SpreadsheetApp.openById(LSPD.ID_Zeitsystem).getSheetByName("Zeitsystem");
  var Zeitsystem_Daten = Zeitsystem.getRange("B4:B").getValues();
  
  for(var i = 0; i < Daten.length; i++)
  {
    if(Daten[i][0] != "" && Daten[i][7] == false)
    {
      var Stunden = Daten[i][2];
      var Minuten = Daten[i][3];

      if(Stunden == "" || Stunden == undefined) Stunden = 0;
      if(Minuten == "" || Minuten == undefined) Minuten = 0;

      var Gefunden = false;
      for(var u = 0; u < Zeitsystem_Daten.length; u++)
      {
        if(Zeitsystem_Daten[u][0] == Daten[i][0])
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden)
      {
        var Zeitsystem_Zeit = new Date(Zeitsystem.getRange("T" + (u + 4)).getValue());
        Zeitsystem_Zeit.setHours(Zeitsystem_Zeit.getHours() - Stunden);
        Zeitsystem_Zeit.setMinutes(Zeitsystem_Zeit.getMinutes() - Minuten);

        Zeitsystem.getRange("T" + (u + 4)).setValue(Zeitsystem_Zeit);
      }

      Sheet.getRange("I" + (i + 7)).setValue(true);
    }
  }
}