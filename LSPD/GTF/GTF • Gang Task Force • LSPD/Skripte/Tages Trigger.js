function Tages_Trigger()
{
  var Sheet_FraktTalk = SpreadsheetApp.getActive().getSheetByName("Fraktionsgespräche");

  Sheet_FraktTalk.getRange("B5:F" + Sheet_FraktTalk.getLastRow()).sort([{column: Spalte_in_Index("C"), ascending: true}, {column: Spalte_in_Index("B"), ascending: false}]);
}