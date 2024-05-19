function Kontrolle(e)
{
  var Sheet_Kontrolle = SpreadsheetApp.getActive().getSheetByName("Kontrolle");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 3 && Zeile <= 7)
  {
    if(Value != undefined)
    {
      var Sheet_Fuhrpark = SpreadsheetApp.getActive().getSheetByName("Fuhrpark");
      var Array_Fuhrpark = Sheet_Fuhrpark.getRange("B4:C253").getValues();

      var Gefunden = false;
      var Modell;

      for(var i = 0; i < Array_Fuhrpark.length; i++)
      {
        if(Array_Fuhrpark[i][0] != "" && Array_Fuhrpark[i][0].toString() == Value)
        {
          Gefunden = true;
          Modell = Array_Fuhrpark[i][1];

          break;
        }
      }

      if(Gefunden == true)
      {
        Array_Fuhrpark = Sheet_Fuhrpark.getRange("I4:M253").getValues();
        Gefunden = false;

        for(var i = 0; i < Array_Fuhrpark.length; i++)
        {
          if(Array_Fuhrpark[i][0] != "" && Array_Fuhrpark[i][0] == Modell)
          {
            Gefunden = true;
            break;
          }
        }

        if(Gefunden == true)
        {
          Sheet_Kontrolle.getRange("E" + Zeile).setNote(`mind. ${Array_Fuhrpark[i][2].toString()} Verbandskästen`);
          Sheet_Kontrolle.getRange("F" + Zeile).setNote(`mind. ${Array_Fuhrpark[i][3].toString()} Reparaturkästen`);
          Sheet_Kontrolle.getRange("G" + Zeile).setNote(`mind. ${Math.floor(Number(Array_Fuhrpark[i][4]) * 0.75).toString()}l`);
        }
        else
        {
          return SpreadsheetApp.getUi().alert("Fehler!", "Dieses Fahrzeug ist nicht in den Kontrolledaten hinterlegt!", SpreadsheetApp.getUi().ButtonSet.OK);
        }
      }
    }
    else
    {
      Sheet_Kontrolle.getRange(`B${Zeile}:K${Zeile}`).clearContent();
      Sheet_Kontrolle.getRange(`B${Zeile}:K${Zeile}`).clearNote();
    }
  }
  else if(Spalte >= Spalte_in_Index("E") && Spalte <= Spalte_in_Index("G") && Zeile >= 3 && Zeile <= 7 && Value != undefined)
  {
    Sheet_Kontrolle.getRange("I" + Zeile).setValue(true);
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 3 && Zeile <= 7 && Value == "TRUE")
  {
    Sheet_Kontrolle.getRange(Zeile, Spalte).uncheck();

    var Array_Eintrag = Sheet_Kontrolle.getRange(`B${Zeile}:I${Zeile}`).getValues();
    Array_Eintrag = Array_Eintrag[0];

    Array_Eintrag[8] = new Date();
    Array_Eintrag[9] = LSPD.Umwandeln();

    Sheet_Kontrolle.insertRowAfter(10);
    Sheet_Kontrolle.getRange("B11:K11").setValues([Array_Eintrag]);

    Sheet_Kontrolle.getRange(`B${Zeile}:K${Zeile}`).clearContent();
    Sheet_Kontrolle.getRange(`B${Zeile}:K${Zeile}`).clearNote();

    if(PropertiesService.getUserProperties().getProperty("rueckleitung").toString().toUpperCase() == "TRUE")
    {
      SpreadsheetApp.getActive().getSheetByName("Offene Kontrollen").setActiveSelection("B2");
    }

    if(Number(Array_Eintrag[1]) >= 65 && Array_Eintrag[1] != "") // Reifenschaden
    {
      SpreadsheetApp.getActive().getSheetByName("Schadensliste").appendRow(["", Array_Eintrag[0], "Reifenschaden", new Date(), LSPD.Umwandeln(), "", ""]);
    }

    if(Array_Eintrag[2] == "Beschädigt" && Array_Eintrag[2] != "") // Motorschaden
    {
      SpreadsheetApp.getActive().getSheetByName("Schadensliste").appendRow(["", Array_Eintrag[0], "Motorschaden", new Date(), LSPD.Umwandeln(), "", ""]);
    }

    if(Array_Eintrag[6] != "")
    {
      if(Array_Eintrag[6].toString().includes("Fehlbetankung mit Motorschaden") == true)
      {
        SpreadsheetApp.getActive().getSheetByName("Schadensliste").appendRow(["", Array_Eintrag[0], "Motorschaden", new Date(), LSPD.Umwandeln(), "", ""]);
        SpreadsheetApp.getActive().getSheetByName("Schadensliste").appendRow(["", Array_Eintrag[0], "Fehlbetankung", new Date(), LSPD.Umwandeln(), "", ""]);
      }
      else if(Array_Eintrag[6].toString().includes("Fehlbetankung") == true)
      {
        SpreadsheetApp.getActive().getSheetByName("Schadensliste").appendRow(["", Array_Eintrag[0], "Fehlbetankung", new Date(), LSPD.Umwandeln(), "", ""]);
      }
    }

    if(Array_Eintrag[7] == true) // Detective Meldung
    {
      var Sheet_Fuhrpark = SpreadsheetApp.getActive().getSheetByName("Fuhrpark");
      var Array_Fuhrpark = Sheet_Fuhrpark.getRange("B4:C253").getValues();

      var Gefunden = false;
      var Modell;

      for(var i = 0; i < Array_Fuhrpark.length; i++)
      {
        if(Array_Fuhrpark[i][0] != "" && Array_Fuhrpark[i][0].toString() == Array_Eintrag[0].toString())
        {
          Gefunden = true;
          Modell = Array_Fuhrpark[i][1];

          break;
        }
      }

      if(Gefunden == true)
      {
        Array_Fuhrpark = Sheet_Fuhrpark.getRange("I4:M253").getValues();
        Gefunden = false;

        for(var i = 0; i < Array_Fuhrpark.length; i++)
        {
          if(Array_Fuhrpark[i][0] != "" && Array_Fuhrpark[i][0] == Modell)
          {
            Gefunden = true;
            break;
          }
        }

        if(Gefunden == true)
        {
          var Detective_Inhalt = [];

          if((Array_Eintrag[3].toString() == "0") || Number(Array_Fuhrpark[i][2]) > Number(Array_Eintrag[3]) && Array_Eintrag[3] != "") // Verbandskästen
          {
            Detective_Inhalt.push(`Fehlerhafte Anzahl an Verbandskästen mit ${Array_Eintrag[3]} von eigentlichen ${Array_Fuhrpark[i][2]}`);
          }

          if((Array_Eintrag[4].toString() == "0") || Number(Array_Fuhrpark[i][3]) > Number(Array_Eintrag[4]) && Array_Eintrag[4] != "") // Reparaturkästen
          {
            Detective_Inhalt.push(`Fehlerhafte Anzahl an Reparaturkästen mit ${Array_Eintrag[4]} von eigentlichen ${Array_Fuhrpark[i][3]}`);
          }

          if(Math.floor(Number(Array_Fuhrpark[i][4]) * 0.75) > Number(Array_Eintrag[5]) && Array_Eintrag[5] != "") // Tankstand
          {
            Detective_Inhalt.push(`Fehlerhafter Tankstand mit ${Array_Eintrag[5]}l von eigentlichen ${Math.floor(Number(Array_Fuhrpark[i][4]) * 0.75)}l`);
          }

          if(Array_Eintrag[6].toString().includes("Unbeaufsichtigt abgestellt (Vespucci)") && Array_Eintrag[6] != "")
          {
            Detective_Inhalt.push(`Fahrzeug in der Vespucci Garage eingeparkt gelassen.`);
          }
          else if(Array_Eintrag[6].toString().includes("Unbeaufsichtigt abgestellt") && Array_Eintrag[6] != "")
          {
            Detective_Inhalt.push(`Fahrzeug unbeaufsichtigt an externem Standort stehen gelassen.`);
          }

          if(Detective_Inhalt.length > 0)
          {
            SpreadsheetApp.getActive().getSheetByName("Detective Meldungen").appendRow(["", Array_Eintrag[0], Detective_Inhalt.toString().replace(",", ", "), new Date, LSPD.Umwandeln(), false, ""]);
          }
        }
        else
        {
          return;
        }
      }
    }
  }
}