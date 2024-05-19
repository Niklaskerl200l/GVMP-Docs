function LST_Rechner(e)
{
  var Sheet_LST_Rechner = SpreadsheetApp.getActive().getSheetByName("LST Rechner");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("L") && Zeile >= 6 && Value == "TRUE")
  {
    var Sheet_Beschwerden_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
    var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

    //Sheet_LST_Rechner.getRange(Zeile,Spalte).setValue("")

    Logger.log("Test")

    var Name = Sheet_LST_Rechner.getRange("B" + Zeile).getValue();
    var Rang = Sheet_LST_Rechner.getRange("C" + Zeile).getValue();
    var Beschwerde = Sheet_LST_Rechner.getRange("K" + Zeile).getValue();
    var Fallnummer = Sheet_Auswertung.getRange("C3").getValue();
    var Letzte_Zeile = Sheet_Beschwerden_Neu.getRange("B1").getValue()

    Sheet_Beschwerden_Neu.getRange("B" + Letzte_Zeile + ":Q" + Letzte_Zeile).setValues([[Fallnummer,Name,"","Offen",new Date(),"Detective","LSPD",new Date(),"LST Zeit",Beschwerde,"","","","","",LSPD.Umwandeln()]])

    //Sheet_Beschwerden_Neu.setActiveSelection("B" + Letzte_Zeile);
    SpreadsheetApp.getActive().toast("Fall erstellt...");
  }
}
