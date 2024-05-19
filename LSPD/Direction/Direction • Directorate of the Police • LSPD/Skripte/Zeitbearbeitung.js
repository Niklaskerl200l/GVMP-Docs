function Zeitbearbeitung(e)
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Zeitbearbeitung");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("F") && Zeile >= 5 && Zeile <= 9 && Value == "TRUE")
  {
    var Daten = Sheet.getRange("B" + Zeile + ":E" + Zeile).getValues();
    Daten = Daten[0];

    if(Daten[1] == "" || Daten[1] == undefined) Daten[1] = 0;
    if(Daten[2] == "" || Daten[2] == undefined) Daten[2] = 0;

    var Array_Daten = 
    [[
      Daten[0],
      Daten[1],
      Daten[2],
      Daten[3],
      new Date(),
      LSPD.Umwandeln(),
      false
    ]];

    Sheet.insertRowAfter(12);
    Sheet.getRange("B13:H13").setValues(Array_Daten);

    Sheet.getRange("B" + Zeile + ":E" + Zeile).clearContent();
    Sheet.getRange(Zeile, Spalte).setValue(false);
  }
}

function Zeitbearbeitung_Zeitsystem()
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Zeitbearbeitung");
  var Daten = Sheet.getRange("B13:H").getValues();

  if(Daten.filter(function(e){return e[0] != "" && e[6] == false}).length == 0) return;

  var Zeitsystem = SpreadsheetApp.openById(LSPD.ID_Zeitsystem).getSheetByName("Zeitsystem");
  var Zeitsystem_Daten = Zeitsystem.getRange("B4:B").getValues();
  
  for(var i = 0; i < Daten.length; i++)
  {
    if(Daten[i][0] != "" && Daten[i][6] == false)
    {
      var Stunden = Daten[i][1];
      var Minuten = Daten[i][2];

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
        var Zeitsystem_Zeit = new Date(Zeitsystem.getRange("F" + (u + 4)).getValue());
        Zeitsystem_Zeit.setHours(Zeitsystem_Zeit.getHours() + Stunden);
        Zeitsystem_Zeit.setMinutes(Zeitsystem_Zeit.getMinutes() + Minuten);

        Zeitsystem.getRange("F" + (u + 4)).setValue(Zeitsystem_Zeit);
      }

      Sheet.getRange("H" + (i + 13)).setValue(true);
    }
  }
}