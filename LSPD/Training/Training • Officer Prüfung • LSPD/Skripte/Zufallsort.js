function Zufallsort() 
{ 
  var Sheet_Auswertungszeug = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
  var Anzahl_Fragen = Sheet_Auswertungszeug.getRange("C14").getValue();

  var Import_Fragen_Anzahl = Sheet_Auswertungszeug.getRange("K1").getValue();
  var Import_Fragen_Array = Sheet_Auswertungszeug.getRange("K3:K" + Import_Fragen_Anzahl).getValues();

  var Array_Zufallszahlen = [];
  var x = 0;
  var gefunden = false; 

  while(x <= Anzahl_Fragen)
  {
    gefunden = false;
    var Zufallszahl = Math.floor(Math.random() * (Import_Fragen_Array.length)) + 1;

    for(var i = 0;i <= x; i++)
    {
      if(Zufallszahl == Array_Zufallszahlen[i])
      {
        gefunden = true;
        break;
      }
    }

    if(gefunden == false)
    {
      Array_Zufallszahlen[x] = Zufallszahl;
      x++;
    }
  }

  var Array_Ausgabe = [];

  for(var i = 0; i < x; i++)
  {
    Array_Ausgabe[i] = Import_Fragen_Array[Array_Zufallszahlen[i]-1];
  }

  Logger.log(Array_Ausgabe)

  return Array_Ausgabe;

}

