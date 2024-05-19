function Bearbeiten(e)
{
  var Sheet_Bearbeiten = SpreadsheetApp.getActive().getSheetByName("Bearbeiten");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("M") && Zeile >= 17 && Value == "TRUE")  // Eintrag Ausgewählt
  {
    var Array_Auswahl = Sheet_Bearbeiten.getRange("B" + Zeile + ":L" + Zeile).getValues();

    Sheet_Bearbeiten.getRange("B6:L6").setValues(Array_Auswahl);
    Sheet_Bearbeiten.getRange("B10:L10").setValues(Array_Auswahl);

    Sheet_Bearbeiten.getRange("M17:M3016").setValue("");
    Sheet_Bearbeiten.getRange("M" + Zeile).setValue(true);
  }
  else if(Spalte == Spalte_in_Index("M") && Zeile >= 17 && Value == "FALSE")  // Eintrag Ausgewählt
  {
    Sheet_Bearbeiten.getRange("B6:L6").setValue("");
    Sheet_Bearbeiten.getRange("B10:L10").setValue("");
  }
}
