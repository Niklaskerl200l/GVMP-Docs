function Besprechung(e)
{
  try
  {
    var Zeile = e.range.getRow();
    var Spalte = e.range.getColumn();
    var Value = e.value;

    if(Zeile >= 7 && Zeile <= 59 && Spalte == 6 && Value == "TRUE")
    {
      var Sheet_Besprechung = SpreadsheetApp.getActive().getSheetByName("Besprechung");
      var Sheet_Personal_Master = SpreadsheetApp.getActive().getSheetByName("Personal Master");

      var Array_Personal_Master = Sheet_Personal_Master.getRange("B6:O").getValues();
      var Array_RV = Sheet_Besprechung.getRange("W7:W14").getValues();

      var Name = Sheet_Besprechung.getRange("B" + Zeile).getValue();
      var Neuer_Rang = Sheet_Besprechung.getRange("E" + Zeile).getValue();

      Sheet_Besprechung.getRange("F" + Zeile).setValue(false);
      
      for(var i = 0; i < Array_Personal_Master.length; i++)
      {
        if(Array_Personal_Master[i][1] == Name)
        {
          
          var Zeile_Personal = i + 6;

          Sheet_Personal_Master.getRange("D" + Zeile_Personal).setValue(Neuer_Rang);
          Sheet_Personal_Master.getRange("J" + Zeile_Personal).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy"));
          break;
        }
      }

      for(var i = 0; i < Array_RV.length; i++)
      {
        if(Array_RV[i][0] == Name)
        {
          Sheet_Besprechung.getRange("W" + (7+i) + ":Y" + (7+i)).setValue("");
          break;
        }
      }
    }
    else if(Zeile >= 7 && Zeile <= 14 && Spalte == Spalte_in_Index("U") && Value == "TRUE")
    {
      var Sheet_Besprechung = SpreadsheetApp.getActive().getSheetByName("Besprechung");
      var Sheet_Personal_Master = SpreadsheetApp.getActive().getSheetByName("Personal Master");

      var Array_Personal_Master = Sheet_Personal_Master.getRange("B6:O").getValues();
      var Array_RV = Sheet_Besprechung.getRange("W7:W14").getValues();

      var Name = Sheet_Besprechung.getRange("Q" + Zeile).getValue();
      var Neuer_Rang = Sheet_Besprechung.getRange("S" + Zeile).getValue();

      Sheet_Besprechung.getRange("U" + Zeile).setValue(false);

      for(var i = 0; i < Array_Personal_Master.length; i++)
      {
        if(Array_Personal_Master[i][1] == Name)
        {
          var Zeile_Personal = i + 6;

          Sheet_Personal_Master.getRange("D" + Zeile_Personal).setValue(Neuer_Rang);
          Sheet_Personal_Master.getRange("J" + Zeile_Personal).setValue(Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy"));
          break;
        }
      }

      for(var i = 0; i < Array_RV.length; i++)
      {
        if(Array_RV[i][0] == Name)
        {
          Sheet_Besprechung.getRange("W" + (7+i) + ":Y" + (7+i)).setValue("");
          break;
        }
      }

      Sheet_Besprechung.getRange("Q" + Zeile).setValue("");
      Sheet_Besprechung.getRange("S" + Zeile + ":T" + Zeile).setValue("");

      Sheet_Master.getRange("B6:S" + Sheet_Master.getLastRow()).sort([{column: 4, ascending: false}, {column: 3, ascending: true}]);
    }
  }
  catch(err)
  {
    SpreadsheetApp.getUi().alert(err);
  }
}

function Delete_Nicht_Rankup()
{
  SpreadsheetApp.getActive().getSheetByName("Besprechung").getRange("Q17:U18").setValue("");
}
