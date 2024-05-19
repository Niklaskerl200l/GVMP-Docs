function Personal(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile >= 6 && Zeile <= 24 && Spalte >= Spalte_in_Index("B") && Spalte <= Spalte_in_Index("Q"))   
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
}

function Personal_Sortieren()
{
  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");
  var Array_Personal = Sheet_Personal.getRange("B5:H22").getValues();
  
  Array_Personal = Array_Personal.sort(Abteilung_Sort);

  for(var y = 0; y < Array_Personal.length; y++)
  {
    Array_Personal[y][0] = '=IF(C' + (y+5) + '="";"";VLOOKUP(C' + (y+5) + ';\'Import Personaltabelle\'!$A$4:$F;2;FALSE))'
    Array_Personal[y][2] = '=IF(C' + (y+5) + '="";"";VLOOKUP(C' + (y+5) + ';\'Import Personaltabelle\'!$A$4:$F;6;FALSE))'
  }

  Sheet_Personal.getRange("B5:H22").setValues(Array_Personal);
  SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
}