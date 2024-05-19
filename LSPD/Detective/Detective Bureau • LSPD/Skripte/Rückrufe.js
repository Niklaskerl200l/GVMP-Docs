function Rueckrufe(e)
{
  var Sheet_Rueckrufe = SpreadsheetApp.getActive().getSheetByName("Rückrufanfragen");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("F") && Zeile >= 4)
  {
    if(Value == "TRUE")
    {
      Sheet_Rueckrufe.getRange(Zeile, Spalte).clearDataValidations();
      Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(LSPD.Umwandeln()); 
    }
    else if(Value == undefined)
    {
      Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
    }
  }
  else if(Spalte == Spalte_in_Index("H") && Zeile >= 4)
  {
    if(Value == "TRUE")
    {
      Sheet_Rueckrufe.getRange(Zeile, Spalte).clearDataValidations();
      Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM HH:mm"));

      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert("Guten Tag, ich melde mich aufgrund ihres angefragten Rückrufes an das Detective Bureau des LSPD. Ich warte auf ihre Rückmeldung. Mit freundlichen Grüßen, " + LSPD.Umwandeln());
    }
    else if(Value == undefined)
    {
      Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
    }
  }
  else if(Spalte == Spalte_in_Index("I") && Zeile >= 4)
  {
    if(Value == "TRUE")
    {
      Sheet_Rueckrufe.getRange(Zeile, Spalte).clearDataValidations();
      Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM HH:mm"));
    }
    else if(Value == undefined)
    {
      Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
    }
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 4 && Value == "TRUE")
  {
    Sheet_Rueckrufe.getRange(Zeile, Spalte).clearDataValidations();
    Sheet_Rueckrufe.getRange(Zeile, Spalte).setValue(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM HH:mm"));

    var Sheet_Beschwerde = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
    var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

    var Zeile_Beschwerde = Sheet_Beschwerde.getRange("B1").getValue() + 1;
    var Fallnummer = Sheet_Auswertung.getRange("C3").getValue();

    var Insert_Array = 
    [[
      Fallnummer,
      "" /* TVO */,
      "",
      "Offen",
      new Date(),
      Sheet_Rueckrufe.getRange(Zeile, Spalte_in_Index("C")).getValue() + "\n" + Sheet_Rueckrufe.getRange(Zeile, Spalte_in_Index("D")).getValue()
    ]];

    Sheet_Beschwerde.getRange(Zeile_Beschwerde, Spalte_in_Index("B"), 1, Insert_Array[0].length).setValues(Insert_Array);
    Sheet_Beschwerde.getRange(Zeile_Beschwerde, Spalte_in_Index("Q")).setValue(LSPD.Umwandeln());

    Sheet_Beschwerde.setActiveSelection("B" + Zeile_Beschwerde);
  }
}