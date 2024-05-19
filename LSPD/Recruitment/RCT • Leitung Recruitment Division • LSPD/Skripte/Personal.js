function Personal(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile >= 6 && Zeile <= 21 && Spalte == Spalte_in_Index("C"))   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
  if(Zeile >= 6 && Zeile <= 21 && Spalte == Spalte_in_Index("G"))   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
}

function Personal_Sortieren()
{
  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");
  var Array_Personal = Sheet_Personal.getRange("B6:J19").getValues();
  
  Array_Personal = Array_Personal.sort(Abteilung_Sort);

  for(var y = 0; y < Array_Personal.length; y++)
  {
    Array_Personal[y][0] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Personaltabelle\'!$A$4:$F;2;FALSE))'
    Array_Personal[y][2] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Personaltabelle\'!$A$4:$F;6;FALSE))'
    Array_Personal[y][4] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Forum ID\'!$B$4:$C;2;FALSE))'
  }

  Sheet_Personal.getRange("B6:J19").setValues(Array_Personal);
  SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
}