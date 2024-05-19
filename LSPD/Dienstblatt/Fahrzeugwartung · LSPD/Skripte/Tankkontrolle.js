function Tankkontrolle(e)
{
  var Sheet_Tankkontrolle = SpreadsheetApp.getActive().getSheetByName("Tankkontrolle");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("F") && Zeile >= 3 && Zeile <= 252 && Value == "TRUE")
  {
    var Array_Eintrag = Sheet_Tankkontrolle.getRange("B" + Zeile + ":E" + Zeile).getValues();
    Array_Eintrag = Array_Eintrag[0];

    Sheet_Tankkontrolle.getRange(Zeile, Spalte).setValue(false);

    if(Number(Array_Eintrag[2]) > Number(Array_Eintrag[3]))
    {
      var Sheet_Meldungen = SpreadsheetApp.getActive().getSheetByName("Detective Meldungen");
      Sheet_Meldungen.appendRow(["", Array_Eintrag[0], `Fehlerhafter Tankstand mit ${Array_Eintrag[3]}l von eigentlichen ${Array_Eintrag[2]}l`, new Date(), LSPD.Umwandeln(), false, ""]);
    }

    Sheet_Tankkontrolle.getRange("B" + Zeile).clearContent();
    Sheet_Tankkontrolle.getRange("E" + Zeile).clearContent();

    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName("Log Tankkontrollen");
    Sheet_Log.appendRow(["", Array_Eintrag[0], new Date(), LSPD.Umwandeln(), ""]);

    SpreadsheetApp.flush();
    Sheet_Tankkontrolle.getRange("B3:C252").sort({column: Spalte_in_Index("C"), ascending: true});
  }
}

function Tankkontrolle_Zufall()
{
  var SS_Fahrzeugwartung = SpreadsheetApp.getActive();

  var Sheet_Tankkontrolle = SS_Fahrzeugwartung.getSheetByName("Tankkontrolle");
  var Sheet_Log = SS_Fahrzeugwartung.getSheetByName("Log Tankkontrollen");
  var Sheet_Fuhrpark = SS_Fahrzeugwartung.getSheetByName("Fuhrpark");

  var Array_Fuhrpark = Sheet_Fuhrpark.getRange("B4:B253").getValues();
  var Array_Log = Sheet_Log.getRange("B3:C").getValues();
  var Array_Ausgabe = [];

  var Anzahl = 50;
  
  var Tagesabstand = new Date();
  Tagesabstand.setDate(Tagesabstand.getDate() - 2);

  Array_Log = Array_Log.filter(function(e){return e[0] != "" && e[1] >= Tagesabstand});

  Logger.log(Array_Log);

  for(var i = 0; i < Anzahl; i++)
  {
    var Zufall = Math.floor(Math.random() * (Array_Fuhrpark.filter(function(e){return e[0] != ""}).length)) + 0;
    var Gefunden = false;

    for(var o = 0; o < Array_Ausgabe.length; o++)
    {
      if(Array_Ausgabe[o][0] == Array_Fuhrpark[Zufall][0])
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      var Gefunden2 = false;

      for(var o = 0; o < Array_Log.length; o++)
      {
        if(Array_Log[o][0] == Array_Fuhrpark[Zufall][0])
        {
          Gefunden2 = true;
          break;
        }
      }

      if(Gefunden2 == false)
      {
        Array_Ausgabe.push([Array_Fuhrpark[Zufall][0]]);
      }
    }
    else
    {
      i--;
    }
  }

  Logger.log(Array_Ausgabe);

  Sheet_Tankkontrolle.getRange("B3:B252").clearContent();
  Sheet_Tankkontrolle.getRange(3, Spalte_in_Index("B"), Array_Ausgabe.length, Array_Ausgabe[0].length).setValues(Array_Ausgabe);

  SpreadsheetApp.flush();

  Sheet_Tankkontrolle.getRange("B3:C252").sort({column: Spalte_in_Index("C"), ascending: true});
}