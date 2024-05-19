var Array_Meldungen = [
  ["Die", "Staatsbank (+ 150 m)", "staatsbank"],
  ["Der", "Juwelier (+ 150m)", "juwe"],
  ["Die", "Mazebank (+ 150m)", "mazebank"],
  ["Das", "Krankenhaus 1 (+ 75m)", "kh1"],
  ["Das", "Krankenhaus 2 (+ 75m)", "kh2"],
  ["Das", "Krankenhaus 3 (+ 75m)", "kh3"],
  ["Das", "Krankenhaus Sandy Shores (+ 75m)", "khs"],
  ["Das", "Krankenhaus Paleto Bay (+ 75m)", "khp"],
  ["Der", "Würfelpark (+ 100m)", "wuerfelpark"],
  ["Das", "Mission Row PD (+ 150m)", "mrpd"],
  ["Die", "Regierung (+ 75m)", "regierung"],
  ["Der", "Muscle Beach (+ 150m)", "musclebeach"]
];

try
{
  for(var x = 0; x < Array_Meldungen.length; x++)
  {
    eval(`function Meldung_${Array_Meldungen[x][2]}(){SpreadsheetApp.getUi().alert("${Array_Meldungen[x][0] + " " + Array_Meldungen[x][1].toString()} ist ab sofort eine Sperrzone. Unbefugtes betreten oder befahren wird mit einer Festnahme geahndet!")}`);
  }
}
catch(err)
{
  Logger.log(err.stack);
}

function Meldung_Manuell()
{
  var UI = SpreadsheetApp.getUi();
  var Radius = UI.prompt("Sperrzone ausrufen...", "Geben Sie eine Distanz ein! (50-500)", UI.ButtonSet.OK).getResponseText();

  if(Radius != undefined && Radius != null && Radius != "")
  {
    Radius = Radius.toString().match("^[0-9]*$");
    if(Radius != null)
    {
      if(Radius <= 500 && Radius >= 50)
      {
        var Meldung_Inhalt = UI.prompt("Sperrzone ausrufen...", "Geben Sie einen Standort an...\n\nBeispiel: [ANGABE] ist ab sofort eine Sperrzone...", UI.ButtonSet.OK).getResponseText();
        if(Meldung_Inhalt != undefined && Meldung_Inhalt != null && Meldung_Inhalt != "")
        {
          UI.alert(`${Meldung_Inhalt} (+ ${Radius}m) ist ab sofort eine Sperrzone. Unbefugtes betreten oder befahren wird mit einer Festnahme geahndet.`);
        }
      }
      else
      {
        UI.alert("Dies ist ein zu großer/kleiner Radius!");
      }
    }
    else
    {
      UI.alert("Dies ist keine Zahl!");
    }
  }
}

function Schlichtung_Manuell()
{
  var UI = SpreadsheetApp.getUi();
  var Standort = UI.prompt("Eingriffsmeldung ausrufen...", "Geben Sie einen Standort an.\n\nBeispiel: Die Parteien [ADJEKTIV] [STANDORT] (z.b. 'in Davis') werden...", UI.ButtonSet.OK).getResponseText();

  if(Standort != undefined && Standort != null && Standort != "")
  {
    UI.alert(`/gov Das LSPD ruft alle Parteien ${Standort} dazu auf gesetzeswidrige Handlungen einzustellen und mit den staatlichen Kräften zu kooperieren.`);
  }
}