
function Personal(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile >= 6 && Zeile <= 24 && Spalte == Spalte_in_Index("C"))   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
  if(Zeile >= 6 && Zeile <= 24 && Spalte >= Spalte_in_Index("E") && Spalte <= Spalte_in_Index("G"))   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
  if(Zeile >= 6 && Zeile <= 24 && Spalte >= Spalte_in_Index("I") && Value == "TRUE")
  {
    var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");
    var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Personal Archiv");

    var Array_Personal = Sheet_Personal.getRange("B" + Zeile + ":H" + Zeile).getValues();

    Sheet_Personal.getRange("C" + Zeile).setValue("");
    Sheet_Personal.getRange("F" + Zeile).setValue("");
    Sheet_Personal.getRange("G" + Zeile).setValue("");
    Sheet_Personal.getRange("H" + Zeile).setValue("");
    Sheet_Personal.getRange("I" + Zeile).setValue("");

    Array_Personal[0][Array_Personal[0].length] = new Date();

    Logger.log(Array_Personal);
    Sheet_Archiv.insertRowBefore(4);

    Sheet_Archiv.getRange("B4:I4").setValues(Array_Personal);

    Sheet_Archiv.setActiveSelection("J4");

    SpreadsheetApp.flush();

    Personal_Sortieren();
  }
}



function Personal_Sortieren()
{
  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");
  var Array_Personal = Sheet_Personal.getRange("B6:H24").getValues();
  
  Array_Personal = Array_Personal.sort(Abteilung_Sort);

  for(var y = 0; y < Array_Personal.length; y++)
  {
    Array_Personal[y][0] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Personaltabelle\'!$A$4:$F;2;FALSE))'
    Array_Personal[y][2] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Personaltabelle\'!$A$4:$F;6;FALSE))'
    Array_Personal[y][3] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Personaltabelle\'!$A$4:$H;8;FALSE))'
  }

  Sheet_Personal.getRange("B6:H24").setValues(Array_Personal);
  SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
}


