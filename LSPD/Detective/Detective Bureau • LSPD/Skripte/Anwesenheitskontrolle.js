function Anwesenheitskontrolle(e)
{
  var Sheet_Anwesenheitskontrolle = SpreadsheetApp.getActive().getSheetByName("Anwesenheitskontrolle");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 204 && Value != undefined)
  {
    Sheet_Anwesenheitskontrolle.getRange("B3").setValue(new Date());
  }
  else if(Spalte == Spalte_in_Index("K") && Zeile >= 5 && Zeile <= 204 && Value == "TRUE")
  {
    var Array_Fehlend = Sheet_Anwesenheitskontrolle.getRange("I" + Zeile + ":J" + Zeile).getValues();
    Array_Fehlend = Array_Fehlend[0];

    Sheet_Anwesenheitskontrolle.getRange(Zeile, Spalte).setValue(false);

    var Sheet_Personalvermerke = SpreadsheetApp.getActive().getSheetByName("Personalvermerke");
    var Array_Personalvermerke = Sheet_Personalvermerke.getRange("B3:G202").getValues();

    var Gefunden = false;

    for(var i = 0; i < Array_Personalvermerke.length; i++)
    {
      if(Array_Personalvermerke[i][0] != "" && Array_Personalvermerke[i][0] == Array_Fehlend[1])
      {
        for(var o = 0; o < Array_Personalvermerke[i].length; o++)
        {
          if(Array_Personalvermerke[i][o].toString().includes("Fehlende Eintragung während Anwesenheitskontrolle") == true)
          {
            Gefunden = true;
            break;
          }

          if(Array_Personalvermerke[i][o] == "")
          {
            Sheet_Personalvermerke.getRange(i + 3, o + 2).setValue("Fehlende Eintragung während Anwesenheitskontrolle am: " + Utilities.formatDate(new Date(), "CET", "dd.MM.yy HH:mm"));
            break;
          }
        }

        break;
      }
    }

    if(Gefunden == true)
    {
      var Sheet_Beschwerden_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
      var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

      var Fallnummer = Sheet_Auswertung.getRange("C3").getValue();

      var Zeile_Beschwerde = Sheet_Beschwerden_Neu.getRange("B1").getValue();

      Sheet_Beschwerden_Neu.getRange("B" + Zeile_Beschwerde + ":Q" + Zeile_Beschwerde).setValues([[Fallnummer, Array_Fehlend[1], "", "Offen", new Date(), ("Detective " + (LSPD.Umwandeln().toString().split(" ")[1])), "LSPD", new Date(), "Missachtung der Dokumentenpflicht", "TVO war während Anwesenheitskontrolle nicht im Dienstblatt eingestempelt, jedoch im Dienst.", "", "", "", "", "Automatisierte Beschwerde, resultierend aus Personalvermerken", "LSPD IT"]]);
    }

    SpreadsheetApp.getActive().toast("Vermerkt!");
  }
}