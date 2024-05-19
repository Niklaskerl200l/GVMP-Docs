function Beschwerden_Abgeschlossen(e) 
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("AA") && Zeile >= 4 && Value == "TRUE")
  {
    var Sheet_Abgeschlossen = SpreadsheetApp.getActive().getSheetByName("Beschwerden Abgeschlossen");
    var Sheet_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden In Bearbeitung");
    var Letzte_Zeile = Sheet_Bearbeitung.getLastRow() + 1;

    var Array_Abgeschlossen = Sheet_Abgeschlossen.getRange("C" + Zeile + ":Z" + Zeile).getValues();

    //Sheet_Abgeschlossen.getRange("AA" + Zeile).setValue("");

    var Array_Revesion = 
    [[
      "R " + Array_Abgeschlossen[0][0],
      "",
      Array_Abgeschlossen[0][1],
      Array_Abgeschlossen[0][2],
      LSPD.Umwandeln(),
      "Revesion",
      "Revesion",
      "",
      "",
      "",
      Array_Abgeschlossen[0][9],
      Array_Abgeschlossen[0][10],
      Array_Abgeschlossen[0][11],
      Array_Abgeschlossen[0][12],
      Array_Abgeschlossen[0][13],
      Array_Abgeschlossen[0][14],
      Array_Abgeschlossen[0][15],
      Array_Abgeschlossen[0][16],
      Array_Abgeschlossen[0][17],
      Array_Abgeschlossen[0][18],
      Array_Abgeschlossen[0][19],
      Array_Abgeschlossen[0][20],
      Array_Abgeschlossen[0][21],
      Array_Abgeschlossen[0][22],
      Array_Abgeschlossen[0][23],
    ]];

    Sheet_Bearbeitung.getRange("B" + Letzte_Zeile + ":Z" + Letzte_Zeile).setValues(Array_Revesion);
    Sheet_Bearbeitung.getRange("AA" + Letzte_Zeile).insertCheckboxes();
    Sheet_Bearbeitung.setActiveSelection("C" + Letzte_Zeile);
  }
}
