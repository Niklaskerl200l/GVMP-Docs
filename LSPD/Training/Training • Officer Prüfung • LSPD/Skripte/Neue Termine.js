function Neue_Termine(e)
{
  var Sheet = e.source.getActiveSheet();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile >= 3 && Zeile <= 1000 && Spalte == 12 && Value == "TRUE")
  {
    var Sheet_Import_Termine = SpreadsheetApp.getActive().getSheetByName("Import Termine");
    var Array_Namen = Sheet_Import_Termine.getRange("B2:C").getValues();

    var Array_Eingabe = Sheet.getRange("B" + Zeile + ":K" + Zeile).getValues();

    Sheet.getRange("L"+Zeile).setValue(false);

    for(var i = Array_Namen.length - 1; i => 0; i--)
    {
      if(Array_Namen[i][1] == Array_Eingabe[0][1] && Array_Eingabe[0][0] == Array_Namen[i][0])
      {
        Sheet.getRange("H"+Zeile+":K"+Zeile).setValue("");
        Sheet_Import_Termine.getRange("H" + (i + 2)).setValue(true);
        Sheet_Import_Termine.getRange("J" + (i + 2) + ":M" + (i + 2)).setValues([[Array_Eingabe[0][6],Array_Eingabe[0][7],Array_Eingabe[0][8],Array_Eingabe[0][9]]]);
        break;
      }
    }
  }
}