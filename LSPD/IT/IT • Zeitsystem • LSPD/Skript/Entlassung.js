function Entlassung()
{ 
  var Sheet_Entlassung = SpreadsheetApp.getActive().getSheetByName("Entlassungen");

  if(Sheet_Entlassung.getRange("B1").getValue() == true)
  {
    var Sheet_Zeitsystem = SpreadsheetApp.getActive().getSheetByName("Zeitsystem");
    var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Archiv Zeitsystem");

    var Name = Sheet_Entlassung.getRange("B2").getValue();

    var Array_Zeitsystem = Sheet_Zeitsystem.getRange("B4:B" + Sheet_Zeitsystem.getRange("B1").getValue()).getValues();

    var Gefunden = false;

    for(var y = 0; y < Array_Zeitsystem.length; y++)
    {
      if(Array_Zeitsystem[y][0] == Name)
      {
        var Array_Archiv = Sheet_Zeitsystem.getRange("B" + (y + 4) + ":Z" + (y + 4)).getValues();

        Logger.log("Entlassung:\n" + Array_Archiv);

        Sheet_Archiv.insertRowAfter(4);

        Sheet_Archiv.getRange("B5:Z5").setValues(Array_Archiv);

        Sheet_Zeitsystem.deleteRow(y + 4);

        Sheet_Entlassung.deleteRow(2);

        Sheet_Zeitsystem.getRange("B4:Z" + Sheet_Zeitsystem.getRange("B1").getValue()).sort([{column: 3, ascending: false},{column: 2, ascending: true}]);

        Gefunden = true;
      }
    }

    if(Gefunden == false)
    {
      Logger.log("Entlassung nicht gefunden " + Name);

      Sheet_Entlassung.deleteRow(2);
    }
  }

  


}
