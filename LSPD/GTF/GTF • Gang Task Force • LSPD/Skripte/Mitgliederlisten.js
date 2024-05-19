function Mitgliederliste(e)
{
  var Sheet_Mitgliederliste = SpreadsheetApp.getActive().getSheetByName("Mitgliederlisten");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile == 5 && Value != "" && Value != undefined)   // Fraktion ausgewählt
  {
    var Array_Formeln = new Array();

    var Datum = new Date();
    Datum.setMinutes(Datum.getMinutes() - 5,0,0);

    for(var i = 10; i < 110; i++)
    {
      var Rang     = '=IF($B' + i + ' = "";"";IFERROR(VLOOKUP($B' + i + ';\'Import Aktuell\'!$B$3:$D;3;false);))';
      var HN       = '=IF($B' + i + ' = "";"";IFERROR(VLOOKUP($B' + i + ';\'Import Aktuell\'!$B$3:$E;4;false);))';
      var Tel      = '=IF($B' + i + ' = "";"";IFERROR(VLOOKUP($B' + i + ';\'Import Aktuell\'!$B$3:$F;5;false);))';

      Array_Formeln.push([Rang,HN,Tel]);
    }

    Sheet_Mitgliederliste.getRange("C5:F5").setValues([["Mitgliederliste",,new Date(),LSPD.Umwandeln()]]);

    Sheet_Mitgliederliste.getRange("H5:L5").setValues([["Zivilist","Mitgliederliste",,Datum,LSPD.Umwandeln()]]);

    Sheet_Mitgliederliste.getRange("C10:E109").setFormulas(Array_Formeln);

    Sheet_Mitgliederliste.getRange("F10:F109").setValue(false);

    Sheet_Mitgliederliste.getRange("L10:L109").setValue(false);

    Sheet_Mitgliederliste.setActiveSelection("D5");
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile == 5 && (Value == "" || Value == undefined))  // Fraktion leeren
  {
    Sheet_Mitgliederliste.getRange("C5:F5").setValue("");

    Sheet_Mitgliederliste.getRange("H5:L5").setValue("");
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile == 5 && Value != "" && Value != undefined)  // Setzte Bemerkung
  {
    Sheet_Mitgliederliste.getRange("I5").setValue(Value);
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile == 5 && (Value == "" || Value == undefined))  // Bemerkung leeren
  {
    Sheet_Mitgliederliste.getRange("I5").setValue("");
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 5 && Value != "" && Value != undefined)  // Setzte Beweis
  {
    Sheet_Mitgliederliste.getRange("J5").setValue(Value);
  }
  else if(Spalte == Spalte_in_Index("D") && Zeile == 5 && (Value == "" || Value == undefined))  // Beweis leeren
  {
    Sheet_Mitgliederliste.getRange("J5").setValue("");
  }
  else if(Spalte == Spalte_in_Index("F") && Zeile >= 10 && Value == "TRUE")     // Mitglied Übertragen
  {
    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName('Logs');

    var Array_Header = Sheet_Mitgliederliste.getRange("B5:F5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("B" + Zeile + ":E" + Zeile).getValues();
    var Letzte_Zeile = Sheet_Log.getRange("B1").getValue();

    var Array_Eingabe = [[Array_Zeile[0][0],Array_Header[0][0],Array_Zeile[0][1],Array_Zeile[0][2],Array_Zeile[0][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]]];

    Logger.log(Array_Eingabe);
    var Array_Ausgabe = Eintrag_Check(Array_Eingabe,0,1,7,8,true);
    
    if(Array_Ausgabe == 1) return 0;

    Sheet_Log.getRange("B" + Letzte_Zeile + ":F" + Letzte_Zeile).setValues([[ Array_Ausgabe[0][0],Array_Ausgabe[0][1], Array_Ausgabe[0][5],new Date(),Array_Ausgabe[0][8]]]);
  }
  else if(Spalte == Spalte_in_Index("L") && Zeile >= 10 && Value == "TRUE")     // Auf Zivi umtragen
  {
    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName('Logs');

    var Array_Header = Sheet_Mitgliederliste.getRange("H5:L5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("H" + Zeile + ":K" + Zeile).getValues();
    var Letzte_Zeile = Sheet_Log.getRange("B1").getValue();

    var Array_Eingabe = [[Array_Zeile[0][0],Array_Header[0][0],Array_Zeile[0][1],Array_Zeile[0][2],Array_Zeile[0][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]]];

    Logger.log(Array_Eingabe);
    var Array_Ausgabe = Eintrag_Check(Array_Eingabe,0,1,7,8,true);
    
    if(Array_Ausgabe == 1) return 0;
    
    Sheet_Log.getRange("B" + Letzte_Zeile + ":F" + Letzte_Zeile).setValues([[ Array_Ausgabe[0][0],Array_Ausgabe[0][1], Array_Ausgabe[0][5],new Date(),Array_Ausgabe[0][8]]]);
  }
  else if(Spalte == Spalte_in_Index("F") && Zeile == 7 && Value == "TRUE")     // Auf Zivi umtragen
  {
    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName('Logs');

    var Array_Header = Sheet_Mitgliederliste.getRange("B5:F5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("B10:E109").getValues();
    var Letzte_Zeile = Sheet_Log.getRange("B1").getValue();
    var Array_Logs = new Array();

    for(var i = 0; i < Array_Zeile.length; i++)
    {
      if(Array_Zeile[i][0] != "")
      {
        var Array_Eingabe = [[Array_Zeile[i][0],Array_Header[0][0],Array_Zeile[i][1],Array_Zeile[i][2],Array_Zeile[i][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]]]
        
        var Array_Ausgabe = Eintrag_Check(Array_Eingabe,0,1,7,8,true);
    
        Array_Logs.push([ Array_Ausgabe[0][0],Array_Ausgabe[0][1], Array_Ausgabe[0][5],new Date(),Array_Ausgabe[0][8]]);
      }
    }

    if(Array_Logs.length > 0)
    {
      Sheet_Log.getRange("B" + Letzte_Zeile + ":F" + (Letzte_Zeile + Array_Logs.length - 1)).setValues(Array_Logs);
    }
  }
  else if(Spalte == Spalte_in_Index("L") && Zeile == 7 && Value == "TRUE")     // Auf Zivi umtragen
  {
    var Sheet_Log = SpreadsheetApp.getActive().getSheetByName('Logs');

    var Array_Header = Sheet_Mitgliederliste.getRange("H5:L5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("H10:M109").getValues();
    var Letzte_Zeile = Sheet_Log.getRange("B1").getValue();
    var Array_Logs = new Array();

    for(var i = 0; i < Array_Zeile.length; i++)
    {
      if(Array_Zeile[i][0] != "" && Array_Zeile[i][5] == false)
      {
        var Array_Eingabe = [[Array_Zeile[i][0],Array_Header[0][0],Array_Zeile[i][1],Array_Zeile[i][2],Array_Zeile[i][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]]]
        
        var Array_Ausgabe = Eintrag_Check(Array_Eingabe,0,1,7,8,true);
    
        Array_Logs.push([ Array_Ausgabe[0][0],Array_Ausgabe[0][1], Array_Ausgabe[0][5],new Date(),Array_Ausgabe[0][8]]);
      }
    }

    if(Array_Logs.length > 0)
    {
      Sheet_Log.getRange("B" + Letzte_Zeile + ":F" + (Letzte_Zeile + Array_Logs.length - 1)).setValues(Array_Logs);
    }
  }
}