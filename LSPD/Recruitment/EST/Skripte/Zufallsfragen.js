function Zufallsfragen()
{
  var Sheet_Fragen = SpreadsheetApp.getActive().getSheetByName("Fragen");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Zeile_Leicht = Sheet_Fragen.getRange("B3").getValue();
  var Zeile_Mittel = Sheet_Fragen.getRange("E3").getValue();
  var Zeile_Schwer = Sheet_Fragen.getRange("H3").getValue();

  var Array_Leicht = Sheet_Fragen.getRange("B6:C" + Zeile_Leicht).getValues();
  var Array_Mittel = Sheet_Fragen.getRange("E6:F" + Zeile_Mittel).getValues();
  var Array_Schwer = Sheet_Fragen.getRange("H6:I" + Zeile_Schwer).getValues();

  var Anzahl_Leicht = Sheet_Auswertung.getRange("C3").getValue();
  var Anzahl_Mittel = Sheet_Auswertung.getRange("C4").getValue();
  var Anzahl_Schwer = Sheet_Auswertung.getRange("C5").getValue();

  var Array_Var = [["Array_Leicht","Anzahl_Leicht"],["Array_Mittel","Anzahl_Mittel"],["Array_Schwer","Anzahl_Schwer"]];

  var Array_Zufallszahlen = [];
  var Gefunden = false;
  for(var z = 0; z < Array_Var.length; z++)
  {
    Array_Zufallszahlen[z] = [];
    Gefunden = false;

    for(var y = 1; y <= eval(Array_Var[z][1]);)
    {
      Gefunden = false;
      var Zufallszahl = Math.floor(Math.random() * ((eval(Array_Var[z][0]).length) - 0 ) + 0);
      
      for(var x = 0; x < Array_Zufallszahlen[z].length; x++)
      {
        if(Array_Zufallszahlen[z][x] == Zufallszahl)
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == false)
      {
        Array_Zufallszahlen[z][Array_Zufallszahlen[z].length] = Zufallszahl;
        y++
      }
    }
  }

  Logger.log(Array_Zufallszahlen);

  var Array_Ausgabe = [];
  var x = 0;

  for(var z = 0; z < Array_Var.length; z++)
  {
    for(var y = 0; y < Array_Zufallszahlen[z].length; y++)
    {
      Array_Ausgabe[x] = new Array();
      Array_Ausgabe[x][0] = eval(Array_Var[z][0])[Array_Zufallszahlen[z][y]][0];
      Array_Ausgabe[x][1] = eval(Array_Var[z][0])[Array_Zufallszahlen[z][y]][1];
      x++;
    }
  }

  Logger.log(Array_Ausgabe)
  return Array_Ausgabe;
}
