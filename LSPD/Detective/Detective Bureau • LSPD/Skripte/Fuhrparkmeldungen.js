function Fuhrparkmeldungen(e)
{
  var Sheet_Fuhrparkmeldungen = SpreadsheetApp.getActive().getSheetByName("Fuhrparkmeldungen");
  var Spalte = e.range.getColumn();
  var Zeile = e.range.getRow();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("D") && Zeile == 2 && Value == "TRUE")
  {
    Sheet_Fuhrparkmeldungen.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.getActive().toast("Suche lädt...");

    var Suchbegriff = Sheet_Fuhrparkmeldungen.getRange("C" + Zeile).getValue();
    if(Suchbegriff == "")
    {
      Suchbegriff = "In Bearbeitung";
    }

    var Kriterium = SpreadsheetApp.newFilterCriteria();

    switch(Suchbegriff)
    {
      case "In Bearbeitung": Kriterium.setHiddenValues(["Geschlossen", "Gemeldet"]); break;
      case "Geschlossen": Kriterium.setHiddenValues(["In Bearbeitung", "Gemeldet"]); break;
      case "Gemeldet": Kriterium.setHiddenValues(["In Bearbeitung", "Geschlossen"]); break;
    }

    Kriterium.build();

    Sheet_Fuhrparkmeldungen.getFilter().setColumnFilterCriteria(Spalte_in_Index("B"), Kriterium);
    SpreadsheetApp.getActive().toast("Suche abgeschlossen!");
  }
  else if(Spalte == Spalte_in_Index("H") && Zeile >= 4 && Value == "TRUE")
  {
    Sheet_Fuhrparkmeldungen.getRange(Zeile, Spalte).removeCheckboxes().setValue(LSPD.Umwandeln());
    Sheet_Fuhrparkmeldungen.getRange("B" + Zeile).setValue("Geschlossen");
  }
  else if(Spalte == Spalte_in_Index("I") && Zeile >= 4 && Value == "TRUE")
  {
    Sheet_Fuhrparkmeldungen.getRange(Zeile, Spalte).removeCheckboxes().setValue(LSPD.Umwandeln());
    Sheet_Fuhrparkmeldungen.getRange("B" + Zeile).setValue("Gemeldet");

    var SS = SpreadsheetApp.getActive();

    var Sheet_Auswertung = SS.getSheetByName("Auswertungsgedöns");
    var Fallnummer = Sheet_Auswertung.getRange("C3").getValue();

    var Sheet_Beschwerde_In_Bearbeitung = SS.getSheetByName("Beschwerden In Bearbeitung");
    var Zeile_Beschwerde = Sheet_Beschwerde_In_Bearbeitung.getRange("B1").getValue();

    var Array_Meldung = Sheet_Fuhrparkmeldungen.getRange("B" + Zeile + ":G" + Zeile).getValues();
    Array_Meldung = Array_Meldung[0];

    Sheet_Beschwerde_In_Bearbeitung.getRange("B" + Zeile_Beschwerde + ":AD" + Zeile_Beschwerde).setValues([[
      Fallnummer,
      "Team",
      Array_Meldung[1],
      "",
      LSPD.Umwandeln(),
      "In Bearbeitung",
      "In Bearbeitung",
      "",
      "",
      "",
      "",
      new Date(),
      "Fahrzeugwartung",
      "LSPD",
      Array_Meldung[3],
      "Missachtung der Fahrzeugordnung (" + Array_Meldung[2] + ")",
      Array_Meldung[4],
      Array_Meldung[5],
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      LSPD.Umwandeln(),
      false,
      new Date()
    ]]);

    Sheet_Beschwerde_In_Bearbeitung.getRange("AC" + Zeile_Beschwerde).insertCheckboxes();
    Sheet_Beschwerde_In_Bearbeitung.setActiveSelection("B" + Zeile_Beschwerde);

    SpreadsheetApp.flush();
  }

  Sheet_Fuhrparkmeldungen.getRange("B4:I2003").setBackground(null);
}