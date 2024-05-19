function Zufallsfunk()
{
  var Sheet_Fragen = SpreadsheetApp.getActive().getSheetByName("Fragen");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Zeile_Funk = Sheet_Fragen.getRange("M3").getValue();

  var Array_Funk = Sheet_Fragen.getRange("M6:N" + Zeile_Funk).getValues();

  var Anzahl_Funk = Sheet_Auswertung.getRange("C11").getValue();

  var Array_Zufallszahlen = [];
  var Gefunden = false;

  for(var y = 1; y <= Anzahl_Funk; )
  {
    Gefunden = false;
    var Zufallszahl = Math.floor(Math.random() * (Array_Funk.length - 0) + 0);

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

  Logger.log(Array_Zufallszahlen);

  var Array_Ausgabe = [];

  for(var x = 0; x < Array_Zufallszahlen.length; x++)
  {
    Array_Ausgabe[x] = new Array();
    Array_Ausgabe[x][0] = Array_Funk[Array_Zufallszahlen[x]][0];
    Array_Ausgabe[x][1] = Array_Funk[Array_Zufallszahlen[x]][1];
  }

  Logger.log(Array_Ausgabe);
  
  return Array_Ausgabe;
}
