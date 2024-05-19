function Mobile_Leitstelle_Minutentrigger()
{
  var SS_Dienstblatt = SpreadsheetApp.getActive();
  var Sheet_Auswertung = SS_Dienstblatt.getSheetByName("Auswertungsgedöns");
  var State_Leitstelle = Sheet_Auswertung.getRange("E20").getValue();

  if(State_Leitstelle == true)
  {
    var Sheet_Startseite = SS_Dienstblatt.getSheetByName("Startseite");
    var Leitstelle = Sheet_Startseite.getRange("D4").getValue();

    if(Leitstelle == "") return;

    var Sheet_Mobile = SS_Dienstblatt.getSheetByName("Log Mobile Leitstelle");
    var Array_Mobile_Heute = Sheet_Mobile.getRange("B6:E205").getValues();

    var Gefunden = false;
    for(var i = 0; i < Array_Mobile_Heute.length; i++)
    {
      if(Array_Mobile_Heute[i][0] != "" && Array_Mobile_Heute[i][0] == Leitstelle)
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      var Array_Mobile_Monat = Sheet_Mobile.getRange("G6:H205").getValues();
      var Gefunden_Monat_Voll = false;
      for(var i = 0; i < Array_Mobile_Monat.length; i++)
      {
        if(Array_Mobile_Monat[i][0] != "" && Array_Mobile_Monat[i][0] == Leitstelle && Array_Mobile_Monat[i][1] >= 60)
        {
          Gefunden_Monat_Voll = true;
          break;
        }
      }

      if(Gefunden_Monat_Voll == false)
      {
        var Zeile = Sheet_Mobile.getRange("E2").getValue();
        Sheet_Mobile.getRange("B" + Zeile + ":C" + Zeile).setValues([[Leitstelle, 1]]);

        SpreadsheetApp.flush();
      }
    }
    else if(Gefunden == true)
    {
      if(Array_Mobile_Heute[i][3] < 60)
      {
        var Zeile = (i + 6);
        Sheet_Mobile.getRange("C" + Zeile).setValue((Array_Mobile_Heute[i][1] + 1));
      }
    }
  }
}

function Mobile_Leitstelle_Tagestrigger()
{
  var Sheet_Mobile = SpreadsheetApp.getActive().getSheetByName("Log Mobile Leitstelle");
  var Array_Heute = Sheet_Mobile.getRange("B6:E205").getValues();
  var Array_Monat = Sheet_Mobile.getRange("G6:H205").getValues();

  var Zeile = Sheet_Mobile.getRange("H2").getValue();
  var Array_Speicher = [];

  for(var i = 0; i < Array_Heute.length; i++)
  {
    if(Array_Heute[i][0] != "")
    {
      var Gefunden = false;
      for(var o = 0; o < Array_Monat.length; o++)
      {
        if(Array_Monat[o][0] == Array_Heute[i][0])
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == true)
      {
        var Zeile_Monat = (o + 6);
        Sheet_Mobile.getRange("H" + Zeile_Monat).setValue(Array_Heute[i][3]);
      }
      else if(Gefunden == false)
      {
        Array_Speicher.push([Array_Heute[i][0], Array_Heute[i][3]]);
      }
    }
  }

  if(Array_Speicher.length > 0)
  {
    Sheet_Mobile.getRange(Zeile, Spalte_in_Index("G"), Array_Speicher.length, Array_Speicher[0].length).setValues(Array_Speicher);
  }

  Sheet_Mobile.getRange("B6:E205").clearContent();
  Sheet_Mobile.getRange("G6:H205").sort({column: Spalte_in_Index("H"), ascending: false});
}

function Mobile_Leitstelle_Monatstrigger()
{
  var Sheet_Mobile = SpreadsheetApp.getActive().getSheetByName("Log Mobile Leitstelle");
  var Array_Monat = Sheet_Mobile.getRange("G6:H205").getValues();

  if(Array_Monat.length > 0)
  {
    var SS_Zeitsystem = SpreadsheetApp.openById(LSPD.ID_Zeitsystem);
    var Sheet_Zeitsystem = SS_Zeitsystem.getSheetByName("Zeitsystem");
    var Array_Zeitsystem = Sheet_Zeitsystem.getRange("B4:B").getValues();

    for(var i = 0; i < Array_Monat.length; i++)
    {
      if(Array_Monat[i][0] != "")
      {
        var Gefunden = false;
        for(var o = 0; o < Array_Zeitsystem.length; o++)
        {
          if(Array_Zeitsystem[o][0] == Array_Monat[i][0])
          {
            Gefunden = true;
            break;
          }
        }

        if(Gefunden == true)
        {
          var Zeit_Zeitsystem = new Date(Sheet_Zeitsystem.getRange("T" + (o + 4)).getValue());
          Zeit_Zeitsystem.setMinutes(Zeit_Zeitsystem.getMinutes() + Array_Monat[i][1]);

          Sheet_Zeitsystem.getRange("T" + (o + 4)).setValue(Zeit_Zeitsystem);
        }
      }
    }

    Sheet_Mobile.getRange("G6:H205").clearContent();
  }
}

function Mobile_Leitstelle_Ausgabe(Benutzer = LSPD.Umwandeln())
{
  var Sheet_Mobile = SpreadsheetApp.getActive().getSheetByName("Log Mobile Leitstelle");
  var Array_Mobile = Sheet_Mobile.getRange("G6:H205").getValues();

  var Gefunden = false;
  for(var i = 0; i < Array_Mobile.length; i++)
  {
    if(Array_Mobile[i][0] != "" && Array_Mobile[i][0] == Benutzer)
    {
      Gefunden = true;
      break;
    }
  }

  if(Gefunden == true)
  {
    var UI = SpreadsheetApp.getUi();
    UI.alert("📞 Mobile Leitstelle", "Du hast für kommenden Monat " + Array_Mobile[i][1].toString() + " Minuten gesammelt!\n\nHinweis: Die Zeit, die du möglicherweise heute gesammelt hast, ist in dieser Zahl noch nicht einberechnet!", UI.ButtonSet.OK);
  }
  else
  {
    SpreadsheetApp.getActive().toast("Du hast noch keine Zeit diesen Monat gesammelt!\n(Die Zeit von heute noch nicht einberechnet...)");
  }
}