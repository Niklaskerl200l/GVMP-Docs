function Ausbildungsblatt(e)
{
  var Sheet_Ausbildungsblatt = SpreadsheetApp.getActive().getSheetByName("Ausbildungsblatt");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile == 2 && Value != undefined || Spalte == Spalte_in_Index("D") && Zeile == 2 && Value != undefined)
  {
    var Name = Sheet_Ausbildungsblatt.getRange(Zeile,Spalte).getValue();
    var Array_Namen = Sheet_Ausbildungsblatt.getRange("B4:B" + Sheet_Ausbildungsblatt.getLastRow()).getValues();

    for(var y = 0; y < Array_Namen.length; y++)
    {
      if(Array_Namen[y][0] == Name)
      {
        Sheet_Ausbildungsblatt.setActiveSelection("B" + (y + 4));
      }
    }

  }
}

function Ausbildungsblatt_Sortieren()
{
  SpreadsheetApp.getActive().getSheetByName("Ausbildungsblatt").getRange("B4:EY").sort([{column: 3, ascending: false},{column: 2, ascending: true}]);
}