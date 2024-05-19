function Zufallsfragen(Name) 
{ 
  var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Officer_Prüfung);
  var Sheet_Export_Uebersicht = Sheet_Export.getSheetByName("Übersicht");

  var Array_Uebersicht = Sheet_Export_Uebersicht.getRange("B3:B" + (Sheet_Export_Uebersicht.getRange("B1").getValue() - 1)).getValues();
  var ID = [];

  for(var y = 0; y < Array_Uebersicht.length; y++)
  {
    if(Array_Uebersicht[y][0] == Name)
    {
      var Formel = Sheet_Export_Uebersicht.getRange("H" + (y+3)).getFormula();

      Formel = Formel.substring(Formel.search("#gid=") + 5,Formel.indexOf("\";",Formel.search("#gid=") + 5 ));
      ID.push(Formel);
    }
  }

  var Array_Export_Fragen = new Array();

  for(var i = 0; i < ID.length; i++)
  {
    try
    {
      function GetSheetById(id) 
      {
        return Sheet_Export.getSheets().filter(function(s) {return s.getSheetId() === id;})[0];
      }

      var Sheet_Archiv = GetSheetById(Number(ID[i]));

      var Letze_Zeile_Links = "B39";
      var Letze_Zeile_Rechts = "L36";
      var Zeile = 9;
      var Spalte = ["B","L"];
      var Spalten_Counter = 0;

      while(true)
      {
        Array_Export_Fragen.push(Sheet_Archiv.getRange(Spalte[Spalten_Counter] + Zeile).getValue());

        if(Spalte[Spalten_Counter] + Zeile == Letze_Zeile_Links)
        {
          Zeile = 6;
          Spalten_Counter = 1;
        }
        else if (Spalte[Spalten_Counter] + Zeile == Letze_Zeile_Rechts)
        {
          break;
        }

        Zeile = Zeile + 3;  
      }
    }
    catch(err)
    {
      Logger.log(err.stack);
    }
  }

  var Sheet_Leicht = SpreadsheetApp.getActive().getSheetByName("Leichte Theoriefragen");
  var Sheet_Mittel = SpreadsheetApp.getActive().getSheetByName("Mittlere Theoriefragen");
  var Sheet_Schwer = SpreadsheetApp.getActive().getSheetByName("Schwere Theoriefragen");
  var Sheet_Auswertungszeug = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Fragen_Leicht_Situationsfragen = Sheet_Auswertungszeug.getRange("C3").getValue();
  var Fragen_Leicht_Rechtsfragen = Sheet_Auswertungszeug.getRange("C4").getValue();
  var Fragen_Leicht_Allgemeinefragen = Sheet_Auswertungszeug.getRange("C5").getValue();
  var Fragen_Mittel_Situationsfragen = Sheet_Auswertungszeug.getRange("C6").getValue();
  var Fragen_Mittel_Rechtsfragen = Sheet_Auswertungszeug.getRange("C7").getValue();
  var Fragen_Mittel_Allgemeinefragen = Sheet_Auswertungszeug.getRange("C8").getValue();
  var Fragen_Schwer_Situationsfragen = Sheet_Auswertungszeug.getRange("C9").getValue();
  var Fragen_Schwer_Rechtsfragen = Sheet_Auswertungszeug.getRange("C10").getValue();
  var Fragen_Schwer_Allgemeinefragen = Sheet_Auswertungszeug.getRange("C11").getValue();

  var Anzahl_Leicht_Situationsfragen = Sheet_Leicht.getRange("B4").getValue();
  var Anzahl_Leicht_Rechtsfragen = Sheet_Leicht.getRange("B63").getValue();
  var Anzahl_Leicht_Allgemeinefragen = Sheet_Leicht.getRange("B122").getValue();
  var Anzahl_Mittel_Situationsfragen = Sheet_Mittel.getRange("B4").getValue();
  var Anzahl_Mittel_Rechtsfragen = Sheet_Mittel.getRange("B63").getValue();
  var Anzahl_Mittel_Allgemeinefragen = Sheet_Mittel.getRange("B122").getValue();
  var Anzahl_Schwer_Situationsfragen = Sheet_Schwer.getRange("B4").getValue();
  var Anzahl_Schwer_Rechtsfragen = Sheet_Schwer.getRange("B63").getValue();
  var Anzahl_Schwer_Allgemeinefragen = Sheet_Schwer.getRange("B122").getValue();

  var Beginn_Leicht_Situationsfragen = 6;
  var Beginn_Leicht_Rechtsfragen = 65;
  var Beginn_Leicht_Allgemeinefragen = 124;
  var Beginn_Mittel_Situationsfragen = 6;
  var Beginn_Mittel_Rechtsfragen = 65;
  var Beginn_Mittel_Allgemeinefragen = 124;
  var Beginn_Schwer_Situationsfragen = 6;
  var Beginn_Schwer_Rechtsfragen = 65;
  var Beginn_Schwer_Allgemeinefragen = 124;

  var Array_Leicht_Situationsfragen = Sheet_Leicht.getRange("C6:D" + Anzahl_Leicht_Situationsfragen).getValues();
  var Array_Leicht_Rechtsfragen = Sheet_Leicht.getRange("C65:D" + Anzahl_Leicht_Rechtsfragen).getValues();
  var Array_Leicht_Allgemeinefragen = Sheet_Leicht.getRange("C124:D" + Anzahl_Leicht_Allgemeinefragen).getValues();
  var Array_Mittel_Situationsfragen = Sheet_Mittel.getRange("C6:D" + Anzahl_Mittel_Situationsfragen).getValues();
  var Array_Mittel_Rechtsfragen = Sheet_Mittel.getRange("C65:D" + Anzahl_Mittel_Rechtsfragen).getValues();
  var Array_Mittel_Allgemeinefragen = Sheet_Mittel.getRange("C124:D" + Anzahl_Mittel_Allgemeinefragen).getValues();
  var Array_Schwer_Situationsfragen = Sheet_Schwer.getRange("C6:D" + Anzahl_Schwer_Situationsfragen).getValues();
  var Array_Schwer_Rechtsfragen = Sheet_Schwer.getRange("C65:D" + Anzahl_Schwer_Rechtsfragen).getValues();
  var Array_Schwer_Allgemeinefragen = Sheet_Schwer.getRange("C124:D" + Anzahl_Schwer_Allgemeinefragen).getValues();

  var Import_Fragen_Anzahl = Sheet_Auswertungszeug.getRange("AB1").getValue();
  var Import_Fragen_Array = Sheet_Auswertungszeug.getRange("AB3:AC" + Import_Fragen_Anzahl).getValues();

  var Array_Var = [["Fragen_Leicht_Situationsfragen","Anzahl_Leicht_Situationsfragen","Beginn_Leicht_Situationsfragen","Array_Leicht_Situationsfragen","5"],["Fragen_Leicht_Rechtsfragen","Anzahl_Leicht_Rechtsfragen","Beginn_Leicht_Rechtsfragen","Array_Leicht_Rechtsfragen","64"],["Fragen_Leicht_Allgemeinefragen","Anzahl_Leicht_Allgemeinefragen","Beginn_Leicht_Allgemeinefragen","Array_Leicht_Allgemeinefragen","123"],["Fragen_Mittel_Situationsfragen","Anzahl_Mittel_Situationsfragen","Beginn_Mittel_Situationsfragen","Array_Mittel_Situationsfragen","5"],["Fragen_Mittel_Rechtsfragen","Anzahl_Mittel_Rechtsfragen","Beginn_Mittel_Rechtsfragen","Array_Mittel_Rechtsfragen","64"],["Fragen_Mittel_Allgemeinefragen","Anzahl_Mittel_Allgemeinefragen","Beginn_Mittel_Allgemeinefragen","Array_Mittel_Allgemeinefragen","123"],["Fragen_Schwer_Situationsfragen","Anzahl_Schwer_Situationsfragen","Beginn_Schwer_Situationsfragen","Array_Schwer_Situationsfragen","5"],["Fragen_Schwer_Rechtsfragen","Anzahl_Schwer_Rechtsfragen","Beginn_Schwer_Rechtsfragen","Array_Schwer_Rechtsfragen","64"],["Fragen_Schwer_Allgemeinefragen","Anzahl_Schwer_Allgemeinefragen","Beginn_Schwer_Allgemeinefragen","Array_Schwer_Allgemeinefragen","123"]];


  var Array_Zufallszahlen = [];
  var gefunden = false;
  var Array_Ausgabe = new Array();

  for(var z = 0; z < Array_Var.length; z++)
  {
    Array_Zufallszahlen[z] = [];

    for(var y = 1; y <= eval(Array_Var[z][0]);)
    {
      gefunden = false;
      var Zufallszahl = Math.floor(Math.random() * (eval(Array_Var[z][1]) -  eval(Array_Var[z][2]))) + eval(Array_Var[z][2]);
    
      var Zufallsfrage = eval(Array_Var[z][3])[Zufallszahl - Number(eval(Array_Var[z][4]))][0];
      var Zufallsantwort = eval(Array_Var[z][3])[Zufallszahl - Number(eval(Array_Var[z][4]))][1];


      for(var x = 0; x < Array_Ausgabe.length; x++)
      {
        if(Array_Ausgabe[x][0] == Zufallsfrage)
        {
          gefunden = true;
          break;
        }
      }

      for(var x = 0; x < Array_Export_Fragen.length; x++)
      {
        if(Array_Export_Fragen[x] == Zufallsfrage)
        {
          gefunden = true;
          break;
        }
      }

      if(gefunden == false)
      {
        Array_Ausgabe.push([Zufallsfrage, Zufallsantwort]);
        y++;        
      }
    }
  }

  Logger.log(Array_Ausgabe)
  return Array_Ausgabe
}


