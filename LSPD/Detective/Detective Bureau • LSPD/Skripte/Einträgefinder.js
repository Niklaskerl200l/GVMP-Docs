function Eintraegefinder(e)
{
  var Sheet_Eintraegefinder = SpreadsheetApp.getActive().getSheetByName("Einträgefinder (NEU)");
  var Spalte = e.range.getColumn();
  var Zeile = e.range.getRow();
  var Value = e.value;

  var Limit = 100;

  if(Spalte == Spalte_in_Index("E") && Zeile == 2 && Value == "TRUE") // Stempeluhr
  {
    Sheet_Eintraegefinder.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.getActive().toast("Bitte warten Sie!\nIhr Anfrage läuft...");

    Sheet_Eintraegefinder.getRange("B6:D105").clearContent();

    var Array_Suche = Sheet_Eintraegefinder.getRange("B4:D4").getValues();
    Array_Suche = Array_Suche[0];

    var Offset_Name = 0;
    var Offset_Beginn = 1;
    var Offset_Ende = 2;

    var SS_Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_Stempeluhr);
    var Sheet_Archiv = SS_Archiv.getSheetByName("Export Archiv Stempeluhr");
    var Array_Archiv = Sheet_Archiv.getRange("B3:D").getValues();

    Array_Archiv = Array_Archiv.filter(function(e){return e[0] != ""});

    if(Array_Suche[Offset_Name] != "") // Suche nach: Name
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[0].toString().toUpperCase() == Array_Suche[Offset_Name].toString().toUpperCase()});
    }

    if(Array_Suche[Offset_Beginn] != "") // Suche nach: Beginn
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[1] >= Array_Suche[Offset_Beginn]});
    }

    if(Array_Suche[Offset_Ende] != "") // Suche nach: Ende
    {
      var Zeitstempel = Array_Suche[Offset_Ende]
      Zeitstempel.setHours(23);
      Zeitstempel.setMinutes(59);
      Zeitstempel.setSeconds(59);
      Array_Archiv = Array_Archiv.filter(function(e){return e[2] <= Zeitstempel});
    }

    var Array_Ausgabe = [];

    for(var i = Array_Archiv.length - 1; i >= 0; i--)
    {
      if(Array_Ausgabe.length < Limit)
      {
        Array_Ausgabe.push([Array_Archiv[i][0], Array_Archiv[i][1], Array_Archiv[i][2]]);
      }
    }

    if(Array_Ausgabe.length > 0)
    {
      Sheet_Eintraegefinder.getRange(6, Spalte_in_Index("B"), Array_Ausgabe.length, Array_Ausgabe[0].length).setValues(Array_Ausgabe);
    }
    else
    {
      SpreadsheetApp.getUi().alert("Kein Fund!");
    }
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile == 2 && Value == "TRUE") // Leitstelle
  {
    Sheet_Eintraegefinder.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.getActive().toast("Bitte warten Sie!\nIhr Anfrage läuft...");

    Sheet_Eintraegefinder.getRange("G6:I105").clearContent();

    var Array_Suche = Sheet_Eintraegefinder.getRange("G4:I4").getValues();
    Array_Suche = Array_Suche[0];

    var Offset_Name = 0;
    var Offset_Beginn = 1;
    var Offset_Ende = 2;

    var SS_Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_Leitstelle);
    var Sheet_Archiv = SS_Archiv.getSheetByName("Archiv Leitstelle");
    var Array_Archiv = Sheet_Archiv.getRange("B3:D").getValues();

    Array_Archiv = Array_Archiv.filter(function(e){return e[0] != ""});

    if(Array_Suche[Offset_Name] != "") // Suche nach: Name
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[0].toString().toUpperCase() == Array_Suche[Offset_Name].toString().toUpperCase()});
    }

    if(Array_Suche[Offset_Beginn] != "") // Suche nach: Beginn
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[1] >= Array_Suche[Offset_Beginn]});
    }

    if(Array_Suche[Offset_Ende] != "") // Suche nach: Ende
    {
      var Zeitstempel = Array_Suche[Offset_Ende]
      Zeitstempel.setHours(23);
      Zeitstempel.setMinutes(59);
      Zeitstempel.setSeconds(59);
      Array_Archiv = Array_Archiv.filter(function(e){return e[2] <= Zeitstempel});
    }

    var Array_Ausgabe = [];

    for(var i = Array_Archiv.length - 1; i >= 0; i--)
    {
      if(Array_Ausgabe.length < Limit)
      {
        Array_Ausgabe.push([Array_Archiv[i][0], Array_Archiv[i][1], Array_Archiv[i][2]]);
      }
    }

    if(Array_Ausgabe.length > 0)
    {
      Sheet_Eintraegefinder.getRange(6, Spalte_in_Index("G"), Array_Ausgabe.length, Array_Ausgabe[0].length).setValues(Array_Ausgabe);
    }
    else
    {
      SpreadsheetApp.getUi().alert("Kein Fund!");
    }
  }
  else if(Spalte == Spalte_in_Index("P") && Zeile == 2 && Value == "TRUE") // Aktionen im Dienstblatt
  {
    Sheet_Eintraegefinder.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.getActive().toast("Bitte warten Sie!\nIhr Anfrage läuft...");

    Sheet_Eintraegefinder.getRange("L6:P105").clearContent();

    var Array_Suche = Sheet_Eintraegefinder.getRange("L4:P4").getValues();
    Array_Suche = Array_Suche[0];

    var Offset_Name = 0;
    var Offset_Aktion = 1;
    var Offset_Wert = 2;
    var Offset_Beginn = 3;
    var Offset_Ende = 4;

    var SS_Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_Dienstblatt_Logs);
    var Sheet_Archiv = SS_Archiv.getSheetByName("Archiv Log");
    var Array_Archiv = Sheet_Archiv.getRange("B3:G").getValues();

    Array_Archiv = Array_Archiv.filter(function(e){return e[0] != ""});

    if(Array_Suche[Offset_Name] != "") // Suche nach: Name
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[0].toString().toUpperCase() == Array_Suche[Offset_Name].toString().toUpperCase()});
    }

    if(Array_Suche[Offset_Aktion] != "") // Suche nach: Aktion
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[2].toString().toUpperCase() == Array_Suche[Offset_Aktion].toString().toUpperCase()});
    }

    if(Array_Suche[Offset_Wert] != "") // Suche nach: Wert
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[3].toString().toUpperCase().includes(Array_Suche[Offset_Wert].toString().toUpperCase())});
    }

    if(Array_Suche[Offset_Beginn] != "") // Suche nach: Beginn
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[1] >= Array_Suche[Offset_Beginn]});
    }

    if(Array_Suche[Offset_Ende] != "") // Suche nach: Ende
    {
      var Zeitstempel = Array_Suche[Offset_Ende]
      Zeitstempel.setHours(23);
      Zeitstempel.setMinutes(59);
      Zeitstempel.setSeconds(59);
      Array_Archiv = Array_Archiv.filter(function(e){return e[1] <= Zeitstempel});
    }

    var Array_Ausgabe = [];

    for(var i = Array_Archiv.length - 1; i >= 0; i--)
    {
      if(Array_Ausgabe.length < Limit)
      {
        Array_Ausgabe.push([Array_Archiv[i][0], Array_Archiv[i][2], Array_Archiv[i][3], Array_Archiv[i][1]]);
      }
    }

    if(Array_Ausgabe.length > 0)
    {
      Sheet_Eintraegefinder.getRange(6, Spalte_in_Index("L"), Array_Ausgabe.length, Array_Ausgabe[0].length).setValues(Array_Ausgabe);
    }
    else
    {
      SpreadsheetApp.getUi().alert("Kein Fund!");
    }
  }
  else if(Spalte == Spalte_in_Index("Y") && Zeile == 2 && Value == "TRUE") // Geschwindigkeitsüberschreitungen
  {
    Sheet_Eintraegefinder.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.getActive().toast("Bitte warten Sie!\nIhr Anfrage läuft...");

    Sheet_Eintraegefinder.getRange("R6:Y105").clearContent();

    var Array_Suche = Sheet_Eintraegefinder.getRange("R4:Y4").getValues();
    Array_Suche = Array_Suche[0];

    var Offset_Name = 0;
    var Offset_20 = 1;
    var Offset_50 = 2;
    var Offset_100 = 3;
    var Offset_101 = 4;
    var Offset_Art = 5;
    var Offset_Beginn = 6;
    var Offset_Ende = 7;

    var SS_Archiv = SpreadsheetApp.openById(LSPD.ID_Geschwindigkeits_Tickets);
    var Sheet_Archiv = SS_Archiv.getSheetByName("Tickets");
    var Array_Archiv = Sheet_Archiv.getRange("B13:J").getValues();

    Array_Archiv = Array_Archiv.filter(function(e){return e[0] != ""});

    if(Array_Suche[Offset_Name] != "") // Suche nach: Name
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[0] == Array_Suche[Offset_Name]});
    }

    if(Array_Suche[Offset_Art] != "") // Suche nach: Art
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[5] == Array_Suche[Offset_Art]});
    }

    if(Array_Suche[Offset_Beginn] != "") // Suche nach: Beginn
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[8] >= Array_Suche[Offset_Beginn]});
    }

    if(Array_Suche[Offset_Ende] != "") // Suche nach: Ende
    {
      var Zeitstempel = Array_Suche[Offset_Ende]
      Zeitstempel.setHours(23);
      Zeitstempel.setMinutes(59);
      Zeitstempel.setSeconds(59);
      Array_Archiv = Array_Archiv.filter(function(e){return e[8] <= Zeitstempel});
    }

    var Array_Ausgabe = [];

    for(var i = 0; i < Array_Archiv.length; i++)
    {
      if(Array_Ausgabe.length < Limit)
      {
        Array_Ausgabe.push([Array_Archiv[i][0], Array_Archiv[i][1], Array_Archiv[i][2], Array_Archiv[i][3], Array_Archiv[i][4], Array_Archiv[i][5], Array_Archiv[i][8], Array_Archiv[i][7]]);
      }
    }

    if(Array_Ausgabe.length > 0)
    {
      Sheet_Eintraegefinder.getRange(6, Spalte_in_Index("R"), Array_Ausgabe.length, Array_Ausgabe[0].length).setValues(Array_Ausgabe);
    }
    else
    {
      SpreadsheetApp.getUi().alert("Kein Fund!");
    }
  }
  else if(Spalte == Spalte_in_Index("AD") && Zeile == 2 && Value == "TRUE") // Einsatzteilnahmen
  {
    Sheet_Eintraegefinder.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.getActive().toast("Bitte warten Sie!\nIhr Anfrage läuft...");

    Sheet_Eintraegefinder.getRange("AA6:AD105").clearContent();

    var Array_Suche = Sheet_Eintraegefinder.getRange("AA4:AD4").getValues();
    Array_Suche = Array_Suche[0];

    var Offset_Name = 0;
    var Offset_Einsatz = 1;
    var Offset_Beginn = 2;
    var Offset_Ende = 3;

    var SS_Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_Einsatz_Logs);
    var Sheet_Archiv = SS_Archiv.getSheetByName("Archiv Einsatz");
    var Array_Archiv = Sheet_Archiv.getRange("B3:D").getValues();

    Array_Archiv = Array_Archiv.filter(function(e){return e[0] != ""});

    if(Array_Suche[Offset_Name] != "") // Suche nach: Name
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[0] == Array_Suche[Offset_Name]});
    }

    if(Array_Suche[Offset_Einsatz] != "") // Suche nach: Einsatz
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[1] == Array_Suche[Offset_Einsatz]});
    }

    if(Array_Suche[Offset_Beginn] != "") // Suche nach: Beginn
    {
      Array_Archiv = Array_Archiv.filter(function(e){return e[2] >= Array_Suche[Offset_Beginn]});
    }

    if(Array_Suche[Offset_Ende] != "") // Suche nach: Ende
    {
      var Zeitstempel = Array_Suche[Offset_Ende]
      Zeitstempel.setHours(23);
      Zeitstempel.setMinutes(59);
      Zeitstempel.setSeconds(59);
      Array_Archiv = Array_Archiv.filter(function(e){return e[2] <= Zeitstempel});
    }

    var Array_Ausgabe = [];

    for(var i = Array_Archiv.length - 1; i >= 0; i--)
    {
      if(Array_Ausgabe.length < Limit)
      {
        Array_Ausgabe.push([Array_Archiv[i][0], Array_Archiv[i][1], Array_Archiv[i][2]]);
      }
    }

    if(Array_Ausgabe.length > 0)
    {
      Sheet_Eintraegefinder.getRange(6, Spalte_in_Index("AA"), Array_Ausgabe.length, Array_Ausgabe[0].length).setValues(Array_Ausgabe);
    }
    else
    {
      SpreadsheetApp.getUi().alert("Kein Fund!");
    }
  }
}