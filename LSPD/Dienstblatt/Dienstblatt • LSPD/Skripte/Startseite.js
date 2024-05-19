function Startseite_NEU(e)
{
  var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  try
  {
    if(Spalte == Spalte_in_Index("L") && Zeile == 21 && Value == "TRUE") // Login: Streifendienst
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Stempeluhr(1, Umwandeln()[2]);

      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("P") && Zeile == 21 && Value == "TRUE") // Login: Büro
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Stempeluhr(2, Umwandeln()[2]);

      Sheet_Startseite.setActiveSelection("A1");
    }
    else if((Spalte == Spalte_in_Index("N") || Spalte == Spalte_in_Index("R")) && Zeile == 21 && Value == "TRUE") // Logout
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Stempeluhr(0, Umwandeln()[2]);

      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("J") && Zeile >= 14 && Zeile <= 16 && Value == "TRUE") // Einsatz: Eintragen (Selbst)
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Einsatz_Eintragung(LSPD.Umwandeln(), Sheet_Startseite.getRange("C" + Zeile).getValue());

      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("J") && Zeile >= 18 && Zeile <= 20 && Value == "TRUE") // Einsatz: Eintragen (Andere)
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Einsatz_Eintragung(Sheet_Startseite.getRange("F" + Zeile).getValue(), Sheet_Startseite.getRange("C" + Zeile).getValue());
      Sheet_Startseite.getRange("F" + Zeile).setValue("")

      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("S") && Zeile == 28 && Value == "TRUE") // Fremde Aktionen: Streifendienst austragen
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Stempeluhr(0, Sheet_Startseite.getRange("T" + Zeile).getValue());

      Log_Zaehler("Anderen Beamten\nAusgetragen Streife", Sheet_Startseite.getRange("T" + Zeile).getValue());
      Sheet_Startseite.getRange("T" + Zeile).clearContent();

      Sheet_Startseite.getRange("T" + Zeile).clearContent();
      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("V") && Zeile == 28 && Value == "TRUE") // Fremde Aktionen: Büro austragen
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Stempeluhr(0, Sheet_Startseite.getRange("W" + Zeile).getValue());

      Log_Zaehler("Anderen Beamten\nAusgetragen Büro", Sheet_Startseite.getRange("W" + Zeile).getValue());
      Sheet_Startseite.getRange("W" + Zeile).clearContent();

      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("S") && Zeile == 30 && Value == "TRUE") // Fremde Aktionen: Personen aus Einsatz austragen
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Einsatz_Eintragung(Sheet_Startseite.getRange("T" + Zeile).getValue(), "");

      Sheet_Startseite.getRange("T" + Zeile).clearContent();
      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("V") && Zeile == 30 && Value == "TRUE") // Fremde Aktionen: Alle aus Einsatz austragen
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      Einsatz_Beenden(Sheet_Startseite.getRange("W" + Zeile).getValue());

      Sheet_Startseite.getRange("W" + Zeile).clearContent();
      Sheet_Startseite.setActiveSelection("A1");
    }
    else if(Spalte == Spalte_in_Index("D") && Zeile >= 4 && Zeile <= 5 && Value != undefined) // Leitstelle
    {
      Sheet_Startseite.setActiveSelection("A1");

      if(Value == "Suche")
      {
        Sheet_Startseite.getRange(Zeile, Spalte).clearContent();
        LST_Suche_Neu(LSPD.Umwandeln());
      }

      if(Sheet_Startseite.getRange("D4:D5").getValues().filter(function(e){return e[0] == Value}).length > 1)
      {
        Sheet_Startseite.getRange(Zeile, Spalte).clearContent();
        SpreadsheetApp.flush();

        return SpreadsheetApp.getUi().alert("Fehler!\nSie sind bereits als Leitstelle eingetragen!");
      }

      Leitstelle(e);
    }
    else if(Spalte == Spalte_in_Index("I") && Zeile == 4 && Value != undefined) // DEFCON angepasst
    {
      Log_Zaehler("Defcon\nAktualisiert", Value);
    }
    else if(Spalte == Spalte_in_Index("E") && Zeile == 6 && Value != undefined) // FUNK angepasst
    {
      Log_Zaehler("Funkkanal\nAktualisiert", Value);
    }
    else if(Spalte == Spalte_in_Index("U") && Zeile >= 10 && Zeile <= 12 && Value != undefined) // WK angepasst
    {
      Sheet_Startseite.getRange("V" + Zeile + ":W" + Zeile).setValues([[new Date(), LSPD.Umwandeln()]]);
      Log_Zaehler("Waffenkisten\nAktualisiert", Value);
    }
    else if(Spalte == Spalte_in_Index("C") && Zeile == 22 && Value == "TRUE") // Mobile LST Zeit auslesen
    {
      Sheet_Startseite.getRange(Zeile, Spalte).setValue(false);
      SpreadsheetApp.flush();
      
      Mobile_Leitstelle_Ausgabe(LSPD.Umwandeln());
    }
    else if(Spalte == Spalte_in_Index("S") && Zeile >= 108 && Zeile <= 119 && Value != undefined) // Platzverweise
    {
      Sheet_Startseite.getRange("W" + Zeile).setValue(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy HH:mm"));
    }
    else if(Spalte == Spalte_in_Index("C") && Zeile >= 97 && Zeile <= 106) // Rückrufe
    {
      if(Value != undefined)
      {
        Sheet_Startseite.getRange("I" + Zeile).setValue(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM. HH:mm"));
      }
      else if(Value == undefined)
      {
        Sheet_Startseite.getRange("C" + Zeile + ":J" + Zeile).clearContent();
      }
    }
    else if(Spalte == Spalte_in_Index("C") && Zeile >= 109 && Zeile <= 119) // Fahrzeugschaden melden
    {
      if(Value != undefined)
      {
        Sheet_Startseite.getRange("H" + Zeile).setValue(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM. HH:mm"));
      }
      else if(Value == undefined)
      {
        Sheet_Startseite.getRange("C" + Zeile + ":H" + Zeile).clearContent();
      }
    }
  }
  catch(err)
  {
    Logger.log(err.stack);
    Fehler++;
  }
}