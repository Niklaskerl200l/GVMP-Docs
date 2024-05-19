function Zufallsort()
{
  var Sheet_Fragen = SpreadsheetApp.getActive().getSheetByName("Fragen");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Anzahl_Orte = Sheet_Auswertung.getRange("C12").getValue();

  var Zeile_Orte = Sheet_Fragen.getRange("K3").getValue();

  var Array_Orte = Sheet_Fragen.getRange("K6:K" + Zeile_Orte).getValues();

  var Array_Zufallszahlen = [];
  var x = 0;
  var Gefunden = false;

  for(var y = 1; y <= Anzahl_Orte; )
  {
    Gefunden = false;
    var Zufallszahl = Math.floor(Math.random() * (Array_Orte.length - 0) + 0);

    for(var x = 0; x < Array_Zufallszahlen.length; x++)
    {
      if(Zufallszahl == Array_Zufallszahlen[x])
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      Array_Zufallszahlen[Array_Zufallszahlen.length] = Zufallszahl;
      y++;
    }
  }

  var Array_Ausgabe = [];

  Logger.log(Array_Zufallszahlen)

  for(var x = 0; x < Array_Zufallszahlen.length; x++)
  {
    Array_Ausgabe[x] = Array_Orte[Array_Zufallszahlen[x]];
  }

  Logger.log(Array_Ausgabe);
  
  return Array_Ausgabe;
}
