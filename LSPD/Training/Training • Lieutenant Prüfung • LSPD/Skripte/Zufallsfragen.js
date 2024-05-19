function Zufallsfragen()
{
  var Sheet_Fragen = SpreadsheetApp.getActive().getSheetByName("Fragen");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Zeile_STPO = Sheet_Fragen.getRange("B3").getValue();
  var Zeile_BDG = Sheet_Fragen.getRange("F3").getValue();
  var Zeile_DV = Sheet_Fragen.getRange("J3").getValue();

  var Array_STPO = Sheet_Fragen.getRange("B6:D" + Zeile_STPO).getValues();
  var Array_BDG = Sheet_Fragen.getRange("F6:H" + Zeile_BDG).getValues();
  var Array_DV = Sheet_Fragen.getRange("J6:L" + Zeile_DV).getValues();

  var Anzahl_STPO = Sheet_Auswertung.getRange("C3").getValue();
  var Anzahl_BDG = Sheet_Auswertung.getRange("C4").getValue();
  var Anzahl_DV = Sheet_Auswertung.getRange("C5").getValue();

  var Array_Var = [["Array_STPO","Anzahl_STPO"],["Array_BDG","Anzahl_BDG"],["Array_DV","Anzahl_DV"]];

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
      Array_Ausgabe[x][2] = eval(Array_Var[z][0])[Array_Zufallszahlen[z][y]][2];
      x++;
    }
  }

  Logger.log(Array_Ausgabe)
  return Array_Ausgabe;
}
