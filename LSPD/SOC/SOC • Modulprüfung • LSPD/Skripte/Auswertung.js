function Auswertung(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("F") && Zeile == 15 && Value == "TRUE")
  {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertung");
    var Sheet_Vorlage = SpreadsheetApp.getActive().getSheetByName("Modulprüfung (Template)");

    var Name = Sheet_Auswertung.getRange("F14").getValue();

    Sheet_Vorlage.copyTo(SS).setName("Modulprüfung (" + Name + ")");

    var Sheet_Pruefung = SpreadsheetApp.getActive().getSheetByName("Modulprüfung (" + Name + ")");

    Sheet_Pruefung.getRange("B5").setValue(new Date());
    Sheet_Pruefung.getRange("F5").setValue(Name);
    Sheet_Pruefung.getRange("J5").setValue(LSPD.Umwandeln());

    
    Sheet_Pruefung.setActiveSelection("O10");
    SS.moveActiveSheet(2);

    Sheet_Auswertung.getRange("F14:F15").setValue(""); 
  }
}