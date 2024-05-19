function Personal(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte >= Spalte_in_Index("C") && Spalte <= Spalte_in_Index("D") && Zeile >= 6 && Zeile <= 35)
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
}

function Personal_Sortieren()
{
  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");
  var Array_Personal = Sheet_Personal.getRange("B6:F20").getValues();
  
  Array_Personal = Array_Personal.sort(Abteilung_Sort);

  for(var y = 0; y < Array_Personal.length; y++)
  {
    Array_Personal[y][0] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Personaltabelle\'!$A$4:$F;2;FALSE))'
    Array_Personal[y][3] = '=IF(C' + (y+6) + '="";"";VLOOKUP(C' + (y+6) + ';\'Import Personaltabelle\'!$A$4:$F;6;FALSE))'
  }

  Sheet_Personal.getRange("B6:F20").setValues(Array_Personal);
  SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
}