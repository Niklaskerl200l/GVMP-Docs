function Detective_Meldung()
{
  var Sheet_Meldungen = SpreadsheetApp.getActive().getSheetByName("Detective Meldungen");
  var Array_Meldungen = Sheet_Meldungen.getRange("B3:F").getValues();

  var SS_Detective = SpreadsheetApp.openById(LSPD.ID_Detective);
  var Sheet_Detective = SS_Detective.getSheetByName("Fuhrparkmeldungen");
  var Array_Detective = [];

  for(var i = 0; i < Array_Meldungen.length; i++)
  {
    if(Array_Meldungen[i][0] != "" && Array_Meldungen[i][4] == false)
    {
      Array_Detective.push(["In Bearbeitung", "Unbekannt", Array_Meldungen[i][0], Array_Meldungen[i][2], Array_Meldungen[i][1], "", false, false]);

      Sheet_Meldungen.getRange("F" + (i + 3)).setValue(new Date());
    }
  }

  if(Array_Detective.length > 0)
  {
    var Zeile = Sheet_Detective.getLastRow() + 1;

    Sheet_Detective.getRange(Zeile, Spalte_in_Index("B"), Array_Detective.length, Array_Detective[0].length).setValues(Array_Detective);
    Sheet_Detective.getRange(Zeile, Spalte_in_Index("H"), Array_Detective.length, 2).insertCheckboxes();

    Sheet_Detective.getRange("B4:I2003").setBackground(null);
  }
}
