function onOpen(e) 
{
  var UI = SpreadsheetApp.getUi();
  var Benutzer = LSPD.Umwandeln(false, false);
  
  Logger.log("onOpen: " + Benutzer);

  if(LSPD.Propertie_Lesen() != null)
  {
    Umwandeln_Abteilung(LSPD.Umwandeln());
  }

  if(LSPD.Propertie_Lesen("LSPD_Abteilung", "User") != null)
  {
    var Abteilungen_JSON = JSON.parse(LSPD.Propertie_Lesen("LSPD_Abteilung", "User"));
    var Abteilungen;
    for(var i = 0; i < Abteilungen_JSON.length; i++)
    {
      if(Abteilungen_JSON[i][0] == "IT" || Abteilungen_JSON[i][2] >= 8)
      {
        Abteilungen = Abteilungen_JSON[i];
        break;
      }
    }

    Logger.log("Array der Abteilungen für Funktions-UI: " + Abteilungen);

    if(Abteilungen != undefined && (Abteilungen[2] >= 8 || Abteilungen[0] == "IT"))
    {
      UI.createMenu("Funktionen")
      .addItem("Beamtenliste erstellen...", "BeamtenlisteForum")
      .addItem("Person einstempeln...", "Stempeluhr_Dritte")
      .addSeparator()
      .addItem("Alle Beamte ausstempeln...", "Logout_Alle")
      .addToUi();
    }

    var Benutzer_Vorhanden = true;
  }

  LSPD.onOpen();

  if(LSPD.Z.toString().includes(Benutzer) == true) // SWAT IT
  {
    SWAT.onOpen();
  }

  if(Benutzer == "Niklas Kerl") // ARMY IT
  {
    //ARMY.onOpen();
  }

  Aktenklaerung_onOpen();
  Rueckrufe_onOpen();

  if(Benutzer_Vorhanden == true)
  {
    var Gefunden = false;
    for(var i = 0; i < Abteilungen_JSON.length; i++)
    {
      if(Abteilungen_JSON[i][0] == "Fahrzeugwartung")
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == true)
    {
      var Anzahl_Kontrollen = SpreadsheetApp.getActive().getSheetByName("Import Beschädigte Fahrzeuge").getRange("D10").getValue();
      if(Anzahl_Kontrollen > 0)
      {
        Logger.log("Fuhrparkwarnung gesendet!");
        SpreadsheetApp.getUi().alert("Fahrzeugwartung", `Achtung!\nEs sind insgesamt ${Anzahl_Kontrollen} Kontrollen für Fahrzeuge offen!${Anzahl_Kontrollen >= 50 ? "\nBITTE DRINGEND DARUM KÜMMERN!" : ""}`, SpreadsheetApp.getUi().ButtonSet.OK);
      }
    }
  }

  SpreadsheetApp.getActive().toast("Dienstblatt ist geladen...", "Dienstblatt");
}