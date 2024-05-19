function Rueckrufe_onOpen()
{
  var Array_Abteilung = JSON.parse(LSPD.Propertie_Lesen("LSPD_Abteilung"));
  if(Array_Abteilung != null)
  {
    var Count = 0;

    var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
    var Array_Startseite = Sheet_Startseite.getRange("G97:G106").getValues();

    if(Array_Startseite.length <= 0)
    {
      return;
    }

    for(var i = 0; i < Array_Startseite.length; i++)
    { 
      if(Array_Startseite[i][0] != "")
      {
        var Division;
        switch(Array_Startseite[i][0])
        {
          case "Direction":     Division = "Directorate of the Police"; break;
          case "Recruitment":   Division = "Recruitment Division"; break;
          case "Training":      Division = "Training Division"; break;
          case "WLD":           Division = "Weapons License Division"; break;
          case "SOC":           Division = "Special Operation Command"; break;
          case "IT":            Division = "IT"; break;
        }

        var Abteilung;
        Array_Abteilung.forEach(i => {if(i[0] != "LSPD" && i[0] == Division){Abteilung = i}});

        try
        {
          if(Abteilung[0] != null)
          {
            if(Abteilung[0] == Division)
            {
              Count++;
            }
          }
        }
        catch(err)
        {
          // Nothing...
        }
      }
    }

    if(Count > 0)
    {
      Logger.log(LSPD.Umwandeln() + " über Rückruf informiert.");
      SpreadsheetApp.flush();

      SpreadsheetApp.getUi().alert("Mitteilung", "Es " + (Count > 1 ? ("liegen " + Count + " Rückrufe") : "liegt ein Rückruf") + " für Sie vor.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}