function Rankdown(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Sheet_Master = SpreadsheetApp.getActive().getSheetByName("Personal Master");

  var Rang = Sheet_Master.getRange("D" + Zeile).getValue();
  var Name = Sheet_Master.getRange("C" + Zeile).getValue();

  Sheet_Master.getRange("R" + Zeile).setValue(false);
  
  if(Rang >= 1 && Rang <= 12 && Value == "TRUE")
  {
    Sheet_Master.getRange("D" + Zeile).setValue(Rang - 1);
    Sheet_Master.getRange("J" + Zeile).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy"));
    
    Sheet_Master.getRange("B6:O" + Sheet_Master.getLastRow()).sort([{column: 4, ascending: false}, {column: 3, ascending: true}]);

    Selection_Master_Name(Name);
  }
}
