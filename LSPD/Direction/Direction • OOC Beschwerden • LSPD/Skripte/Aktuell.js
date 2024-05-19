function Aktuell(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;
  var UI = SpreadsheetApp.getUi();

  var Sheet_Aktuell = SpreadsheetApp.getActive().getSheetByName("Aktuell");
  var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Archiv");


  if(Spalte == Spalte_in_Index("Q") && Zeile >= 3 && Value == "TRUE" && SpreadsheetApp.getUi().alert("Willst du das wirklich Archivieren?",SpreadsheetApp.getUi().ButtonSet.YES_NO) == SpreadsheetApp.getUi().Button.YES)
  {

    
    var Array_Eintrag = Sheet_Aktuell.getRange("B" + Zeile + ":Q" + Zeile).getValues();

    Sheet_Aktuell.getRange("B" + Zeile + ":Q" + Zeile).setValue("");

    Sheet_Archiv.insertRowAfter(3);

    Sheet_Archiv.getRange("B4:Q4").setValues(Array_Eintrag);
    Sheet_Archiv.getRange("Q4").setValue(new Date());
  }
  }