function Personalvermerke_Aktualisieren()
{
  var Sheet_Personalvermerke = SpreadsheetApp.getActive().getSheetByName("Personalvermerke");
  var Array_Personalvermerke = Sheet_Personalvermerke.getRange("B3:B202").getValues();

  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Import Personaltabelle");
  var Array_Personal = Sheet_Personal.getRange("D4:D199").getValues();

  var Array_Add = [];
  for(var i = 0; i < Array_Personal.length; i++)
  {
    var Gefunden = false;
    for(var u = 0; u < Array_Personalvermerke.length; u++)
    {
      if(Array_Personalvermerke[u][0] == Array_Personal[i][0])
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      Array_Add.push([Array_Personal[i][0]]);
    }
  }

  if(Array_Add.length > 0)
  {
    Sheet_Personalvermerke.getRange(Sheet_Personalvermerke.getRange("B1").getValue(), Spalte_in_Index("B"), Array_Add.length, 1).setValues(Array_Add);
  }

  for(var i = 0; i < Array_Personalvermerke.length; i++)
  {
    var Gefunden = false;
    for(var u = 0; u < Array_Personal.length; u++)
    {
      if(Array_Personal[u][0] == Array_Personalvermerke[i][0])
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      Sheet_Personalvermerke.getRange("B" + (i + 3) + ":G" + (i + 3)).clearContent();
    }
  }

  Sheet_Personalvermerke.getRange("B3:G202").sort([{column: Spalte_in_Index("C"), ascending: false}, {column: Spalte_in_Index("B"), ascending: true}]);
}