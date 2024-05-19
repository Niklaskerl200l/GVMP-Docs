function LST_Plus(e)
{
  var Sheet_LST_Plus = SpreadsheetApp.getActive().getSheetByName("Plus Leitstelle");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("C") && Zeile >= 3 && Zeile <= 4 && Value != undefined && Number(Value) > 1)
  {
    var UI = SpreadsheetApp.getUi();
    return UI.alert("Achtung!", "Sollten Sie Leitstellenzeit wegen eines Überstandes durch die Suche eintragen, so beachten Sie, dass das maximum an anrechbaren Stunden 1 Stunde beträgt.", UI.ButtonSet.OK);
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 3 && Zeile <= 4 && Value == "TRUE")
  {
    Sheet_LST_Plus.getRange(Zeile, Spalte).setValue(false);

    var Array_Eintrag = Sheet_LST_Plus.getRange("B" + Zeile + ":F" + Zeile).getValues();
    Array_Eintrag = Array_Eintrag[0];

    if(Array_Eintrag[1] == "" || Array_Eintrag[1] == undefined) Array_Eintrag[1] = 0;
    if(Array_Eintrag[2] == "" || Array_Eintrag[2] == undefined) Array_Eintrag[2] = 0;

    if(Array_Eintrag[4].toString().toUpperCase() == "TRUE")
    {
      var Ausfuehrung = new Date();
      Ausfuehrung.setDate(1);
      Ausfuehrung.setMonth(Ausfuehrung.getMonth() + 1);

      if(Ausfuehrung.getMonth() == 12)
      {
        Ausfuehrung.setFullYear(Ausfuehrung.getFullYear() + 1);
      }

      Ausfuehrung.setHours(00);
      Ausfuehrung.setMinutes(05);
      Ausfuehrung.setSeconds(00);

      Logger.log(Ausfuehrung);
    }
    else
    {
      Ausfuehrung = new Date();
    }

    var Array_Insert = 
    [[
      Array_Eintrag[0],
      Array_Eintrag[1],
      Array_Eintrag[2],
      Array_Eintrag[3],
      Ausfuehrung,
      new Date(),
      LSPD.Umwandeln(),
      false
    ]];

    Sheet_LST_Plus.insertRowAfter(7);
    Sheet_LST_Plus.getRange("B8:I8").setValues(Array_Insert);

    Sheet_LST_Plus.getRange("B" + Zeile + ":E" + Zeile).clearContent();
    Sheet_LST_Plus.getRange("F" + Zeile).setValue(false).insertCheckboxes();
  }
}

function LST_Plus_Zeitsystem()
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Plus Leitstelle");
  var Daten = Sheet.getRange("B8:I").getValues();

  if(Daten.filter(function(e){return e[0] != "" && e[7] == false}).length == 0) return;

  var Zeitsystem = SpreadsheetApp.openById(LSPD.ID_Zeitsystem).getSheetByName("Zeitsystem");
  var Zeitsystem_Daten = Zeitsystem.getRange("B4:B").getValues();
  
  for(var i = 0; i < Daten.length; i++)
  {
    if(Daten[i][0] != "" && Daten[i][7] == false && Daten[i][4] <= new Date())
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
        var Zeitsystem_Zeit = new Date(Zeitsystem.getRange("T" + (u + 4)).getValue());
        Zeitsystem_Zeit.setHours(Zeitsystem_Zeit.getHours() + Stunden);
        Zeitsystem_Zeit.setMinutes(Zeitsystem_Zeit.getMinutes() + Minuten);

        Zeitsystem.getRange("T" + (u + 4)).setValue(Zeitsystem_Zeit);
      }

      Sheet.getRange("I" + (i + 8)).setValue(true);
    }
  }
}