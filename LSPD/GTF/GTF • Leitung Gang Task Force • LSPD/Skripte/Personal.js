function Personal(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");

  if(Spalte >= Spalte_in_Index("C") && Spalte <= Spalte_in_Index("F") && Zeile >= 3 && Zeile <= 15)
  {
    SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
  }
  
  if(Spalte == Spalte_in_Index("C") && Zeile >= 3 && Zeile <= 15 && Value != undefined && Value != "" && (OldValue == undefined || OldValue == ""))
  {
    SpreadsheetApp.getActive().getSheetByName("Personal").getRange(Zeile,5).setValue(new Date());
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile >= 3 && Zeile <= 15 && OldValue != undefined && OldValue != "" && (Value == undefined || Value == ""))
  {
    SpreadsheetApp.getActive().getSheetByName("Personal").getRange(Zeile,5,1,3).setValue("");
  }
}

function Personal_Sortieren()
{
  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personal");
  var Array_Personal = Sheet_Personal.getRange("B3:G15").getValues();
  
  Array_Personal = Array_Personal.sort(Abteilung_Sort);

  for(var y = 0; y < Array_Personal.length; y++)
  {
    Array_Personal[y][0] = '=IF(C' + (y+3) + '="";"";VLOOKUP(C' + (y+3) + ';\'Import Personaltabelle\'!$A$4:$F;2;FALSE))'
    Array_Personal[y][2] = '=IF(C' + (y+3) + '="";"";VLOOKUP(C' + (y+3) + ';\'Import Personaltabelle\'!$A$4:$F;6;FALSE))'
  }

  Sheet_Personal.getRange("B3:G15").setValues(Array_Personal);
  SpreadsheetApp.getActive().getSheetByName("Export Abteilungen").getRange("B11").setValue("WAHR");
}