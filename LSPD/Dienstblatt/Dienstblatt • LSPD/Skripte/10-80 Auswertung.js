function Auswertung_1080(e)
{
  var Sheet_1080 = SpreadsheetApp.getActive().getSheetByName("10-80 Auswertung");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("F") && Zeile >= 5 && Zeile <= 6 && Value == "TRUE")
  {
    Sheet_1080.getRange(Zeile, Spalte).setValue(false);

    var Array_1080 = Sheet_1080.getRange("B" + Zeile + ":E" + Zeile).getValues();
    Sheet_1080.insertRowAfter(9);
    Sheet_1080.getRange("B10:E10").setValues(Array_1080);
    Sheet_1080.getRange("F10:G10").setValues([[new Date(), LSPD.Umwandeln()]]);

    Sheet_1080.getRange("B" + Zeile + ":E" + Zeile).clearContent();
    Log_Zaehler("10-80 archiviert", "Gestellt: " + Array_1080[0][1] + "\nGeflohen: " + Array_1080[0][2]);
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile >= 5 && Zeile <= 6 && Value == "TRUE")
  {
    if(Sheet_1080.getRange(Zeile, Spalte + 1).getValue() == true)
    {
      Sheet_1080.getRange(Zeile, Spalte).setValue(false);
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert("Fehler!\nSie können einen 10-80 nicht als 'Gestellt' eintragen, währenddessen dieser als 'Geflohen' vermerkt ist.");
    }
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile >= 5 && Zeile <= 6 && Value == "TRUE")
  {
    if(Sheet_1080.getRange(Zeile, Spalte - 1).getValue() == true)
    {
      Sheet_1080.getRange(Zeile, Spalte).setValue(false);
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert("Fehler!\nSie können einen 10-80 nicht als 'Geflohen' eintragen, währenddessen dieser als 'Gestellt' vermerkt ist.");
    }
  }
}