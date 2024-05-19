function Install_onEdit(e)
{
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  
  
  
  

  var Off_Name = 0, Off_Fraktion = 1, Off_Rang = 2, Off_Haus = 3, Off_Tel = 4, Off_Aktivitaet = 7, Off_Beamter = 8, Off_Akteneintrag = 9, Off_Datenbank = 10;


  //-------------------------------------- Dokumentation -----------------------------------//

  if(SheetName == "Dokumentation" && Spalte == Spalte_in_Index("K") && Zeile >= 5 && Zeile <= 8 && Value == "TRUE")   // Eintrag in Datenbank
  {
    var Sheet_Dokumentation = SpreadsheetApp.getActive().getSheetByName("Dokumentation");
    
    var Array_Eingabe = Sheet_Dokumentation.getRange("B" + Zeile + ":J" + Zeile).getValues();

    var Error = Datenbank_Eintrag(Array_Eingabe,Off_Name,Off_Fraktion,Off_Rang,Off_Haus,Off_Tel,Off_Aktivitaet,Off_Beamter,Off_Akteneintrag);
  
    if(Error == 1)    // Wenn falsche Eingaben Stop
    {
      Sheet_Dokumentation.getRange("K" + Zeile).setValue("");
      Stop++;
      return 0;
    }
    else if(Error == 2)   // Wenn unerwarteter Fehler im Script
    {
      Stop++;
      return 0;
    }

    var Fraktion = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$C;2;false);))';
    var Rang     = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$D;3;false);))';
    var HN       = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$E;4;false);))';
    var Tel      = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$F;5;false);))';

    Sheet_Dokumentation.getRange("B" + Zeile + ":K" + Zeile).setValues([["",Fraktion,Rang,HN,Tel,"","","","",""]]);

    return 0;
  }

  //-------------------------------------------- ENDE ------------------------------------------//



  //-------------------------------------- Bearbeiten -----------------------------------//

  else if(SheetName == "Bearbeiten" && Spalte == Spalte_in_Index("M") && Zeile == 6 && Value == "TRUE")   // Eintag aus Datenbank Löschen
  {
    var Sheet_Bearbeiten = SpreadsheetApp.getActive().getSheetByName("Bearbeiten");
    var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);

    var Array_Auswahl = Sheet_Bearbeiten.getRange("B6:L6").getValues();

    Logger.log("Auswahl: " + Array_Auswahl);

    if(Array_Auswahl[0][Off_Datenbank] == "Aktuell")
    {
      var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");
      var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");

      var Array_DB_Aktuell = Sheet_Aktuell.getRange("B3:B").getValues();
      var Array_DB_Archiv = Sheet_Archiv.getRange("B3:K").getValues();
      var Such_Zeile = 0;
      var Letzter_Zeile = 0;

      var Array_Temp = Array_Auswahl[0];
      var Array_Letzer = new Array();
      Array_Temp.pop()

      for(var i = 0; i < Array_DB_Archiv.length; i++)
      {
        if(Array_Letzer.length == 0 && Array_DB_Archiv[i][Off_Name] == Array_Auswahl[0][Off_Name])
        {
          Array_Letzer = Array_DB_Archiv[i];
          Letzter_Zeile = i+3;
        }

        if(Array_Letzer[Off_Aktivitaet] < Array_DB_Archiv[i][Off_Aktivitaet] && Array_DB_Archiv[i][Off_Name] == Array_Auswahl[0][Off_Name])
        {
          Array_Letzer = Array_DB_Archiv[i];
          Letzter_Zeile = i+3;
        }
      }

      for(var y = 0; y < Array_DB_Aktuell.length; y++)  // Suche nach Eintrag in Datenbank
      {
        if(Array_DB_Aktuell[y][Off_Name] == Array_Auswahl[0][Off_Name])
        {
          Such_Zeile = y + 3;

          Logger.log("Suche: " + Array_Temp);

          var Array_Gefunden = Sheet_Aktuell.getRange("B" + Such_Zeile + ":K" + Such_Zeile).getValues();

          if(Array_Temp.toString() == Array_Gefunden.toString())
          {
            Logger.log("Lösche Zeile " + Such_Zeile + " aus DB Aktuell mit dem Eintrag:\n" + Array_Gefunden);

            Sheet_Aktuell.deleteRow(Such_Zeile);

            if(Array_Letzer.length > 0 && Letzter_Zeile != 0)
            {
              Array_Letzer.unshift("");
              Logger.log("Füge Letzten Eintrag aus Archiv in Aktuell ein\n" + Array_Letzer);
              Sheet_Aktuell.appendRow(Array_Letzer);
              Logger.log("Delet Letzer Eintrag aus Archiv in Zeile " + Letzter_Zeile);
              Sheet_Archiv.deleteRow(Letzter_Zeile);
            }

            Sheet_Bearbeiten.getRange("B6:M6").setValue("");
            Sheet_Bearbeiten.getRange("B10:M10").setValue("");
            Sheet_Bearbeiten.getRange("M17:M3016").setValue("");

            return 0;
          }
          else
          {
            Logger.log("Gefundener Eintrag in Zeile " + Such_Zeile + " unterscheidet sich von dem Ausgewählten. Suche weiter.\nGefunden: " + Array_Gefunden)
          }
        }
      }

      Logger.log("Datenbank Eintrag nicht gefunden:\n" + Array_Auswahl);
    }
    else if(Array_Auswahl[0][Off_Datenbank] == "Archiv")
    {
      var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");

      var Array_DB_Archiv = Sheet_Archiv.getRange("B3:B").getValues();
      var Such_Zeile = 0;

      var Array_Temp = Array_Auswahl[0];
      Array_Temp.pop()

      for(var y = 0; y < Array_DB_Archiv.length; y++)  // Suche nach Eintrag in Datenbank
      {
        if(Array_DB_Archiv[y][Off_Name] == Array_Auswahl[0][Off_Name])
        {
          Such_Zeile = y + 3;

          Logger.log("Suche: " + Array_Temp);

          var Array_Gefunden = Sheet_Archiv.getRange("B" + Such_Zeile + ":K" + Such_Zeile).getValues();

          if(Array_Temp.toString() == Array_Gefunden.toString())
          {
            Logger.log("Lösche Zeile " + Such_Zeile + " aus DB Archiv mit dem Eintrag:\n" + Array_Temp);

            Sheet_Archiv.deleteRow(Such_Zeile);

            Sheet_Bearbeiten.getRange("B6:M6").setValue("");
            Sheet_Bearbeiten.getRange("B10:M10").setValue("");
            Sheet_Bearbeiten.getRange("M17:M3016").setValue("");

            return 0;
          }
          else
          {
            Logger.log("Gefundener Eintrag in Zeile " + Such_Zeile + " unterscheidet sich von dem Ausgewählten. Suche weiter.\nGefunden: " + Array_Gefunden)
          }
        }
      }

      Logger.log("Datenbank Eintrag nicht gefunden:\n" + Array_Auswahl);
    }
  }



  //------------------------------- Speichern --------------------------------------//


  else if(SheetName == "Bearbeiten" && Spalte == Spalte_in_Index("M") && Zeile == 10 && Value == "TRUE")   // Eintag aus Datenbank Löschen
  {
    var Sheet_Bearbeiten = SpreadsheetApp.getActive().getSheetByName("Bearbeiten");
    var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);

    var Array_Auswahl = Sheet_Bearbeiten.getRange("B6:L6").getValues();
    var Array_Eingabe = Sheet_Bearbeiten.getRange("B10:K10").getValues();

    Logger.log("Auswahl: " + Array_Auswahl);

    if(Array_Auswahl[0][Off_Datenbank] == "Aktuell")
    {
      var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");

      var Array_DB_Aktuell = Sheet_Aktuell.getRange("B3:B").getValues();
      var Such_Zeile = 0;

      var Array_Temp = Array_Auswahl[0];
      Array_Temp.pop()

      for(var y = 0; y < Array_DB_Aktuell.length; y++)  // Suche nach Eintrag in Datenbank
      {
        if(Array_DB_Aktuell[y][Off_Name] == Array_Auswahl[0][Off_Name])
        {
          Such_Zeile = y + 3;

          Logger.log("Suche: " + Array_Temp);

          var Array_Gefunden = Sheet_Aktuell.getRange("B" + Such_Zeile + ":K" + Such_Zeile).getValues();

          if(Array_Temp.toString() == Array_Gefunden.toString())
          {
            Logger.log("Bearbeite Zeile " + Such_Zeile + " aus DB Aktuell zu:\n" + Array_Eingabe);

            Sheet_Aktuell.getRange("B" + Such_Zeile + ":K" + Such_Zeile).setValues(Array_Eingabe);

            Sheet_Bearbeiten.getRange("B6:M6").setValue("");
            Sheet_Bearbeiten.getRange("B10:M10").setValue("");
            Sheet_Bearbeiten.getRange("M17:M3016").setValue("");

            return 0;
          }
          else
          {
            Logger.log("Gefundener Eintrag in Zeile " + Such_Zeile + " unterscheidet sich von dem Ausgewählten. Suche weiter.\nGefunden: " + Array_Gefunden)
          }
        }
      }

      Logger.log("Datenbank Eintrag nicht gefunden:\n" + Array_Auswahl);
    }
    else if(Array_Auswahl[0][Off_Datenbank] == "Archiv")
    {
      var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");

      var Array_DB_Archiv = Sheet_Archiv.getRange("B3:B").getValues();
      var Such_Zeile = 0;

      var Array_Temp = Array_Auswahl[0];
      Array_Temp.pop();

      for(var y = 0; y < Array_DB_Archiv.length; y++)  // Suche nach Eintrag in Datenbank
      {
        if(Array_DB_Archiv[y][Off_Name] == Array_Auswahl[0][Off_Name])
        {
          Such_Zeile = y + 3;

          Logger.log("Suche: " + Array_Temp);

          var Array_Gefunden = Sheet_Archiv.getRange("B" + Such_Zeile + ":K" + Such_Zeile).getValues();

          if(Array_Temp.toString() == Array_Gefunden.toString())
          {
            Logger.log("Bearbeite Zeile " + Such_Zeile + " aus DB Aktuell zu:\n" + Array_Eingabe);

            Sheet_Archiv.getRange("B" + Such_Zeile + ":K" + Such_Zeile).setValues(Array_Eingabe);

            Sheet_Bearbeiten.getRange("B6:M6").setValue("");
            Sheet_Bearbeiten.getRange("B10:M10").setValue("");
            Sheet_Bearbeiten.getRange("M17:M3016").setValue("");

            return 0;
          }
          else
          {
            Logger.log("Gefundener Eintrag in Zeile " + Such_Zeile + " unterscheidet sich von dem Ausgewählten. Suche weiter.\nGefunden: " + Array_Gefunden)
          }
        }
      }

      Logger.log("Datenbank Eintrag nicht gefunden:\n" + Array_Auswahl);
    }
  }
  
  //-------------------------------------------- ENDE ------------------------------------------//




  //------------------------------------- Akteneinträge ----------------------------------------//

  else if(SheetName == "Akteneinträge" && Spalte == Spalte_in_Index("K") && Zeile >= 7 && Value == "TRUE")
  {
    var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);

    var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Akteneinträge");
    var Sheet_Akten = SpreadsheetApp.openById(LSPD.ID_Akteneinträge).getSheetByName("Akteneinträge");
    var Sheet_Eintrag = SpreadsheetApp.getActive().getSheetByName("Akteneinträge");
    var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");

    var Array_DB_Aktuell = Sheet_Aktuell.getRange("B3:B" + Sheet_Aktuell.getLastRow()).getValues();
    var Array_Eintrag = Sheet_Eintrag.getRange("B" + Zeile + ":J" + Zeile).getValues();
    
    for(var i = 0; i < Array_DB_Aktuell.length; i++)
    {
      if(Array_DB_Aktuell[i][Off_Name] == Array_Eintrag[0][0])   // Suche nach Namen in DB Aktuell
      {
        Logger.log("Setzt Akteneintrag von: " + Array_DB_Aktuell[i][Off_Name]);
        Sheet_Aktuell.getRange("K" + (i+3)).setValue(new Date());   // Setze Datum

        if(Array_Eintrag[0][2] != Array_Eintrag[0][6])
        {
          Logger.log("Setzt Rang von: " + Array_DB_Aktuell[i][Off_Name] + " auf: " + Array_Eintrag[0][6]);
          Sheet_Aktuell.getRange("D" + (i+3)).setValue(Array_Eintrag[0][6]);   // Setze Rang
        }

        if(Array_Eintrag[0][3] != Array_Eintrag[0][7])
        {
          Logger.log("Setzt Haus Nummer von: " + Array_DB_Aktuell[i][Off_Name] + " auf: " + Array_Eintrag[0][7]);
          Sheet_Aktuell.getRange("E" + (i+3)).setValue(Array_Eintrag[0][7]);   // Setze HN
        }

        if(Array_Eintrag[0][4] != Array_Eintrag[0][8]) 
        {
          Logger.log("Setzt Tel von: " + Array_DB_Aktuell[i][Off_Name] + " auf: " + Array_Eintrag[0][8]);
          Sheet_Aktuell.getRange("F" + (i+3)).setValue(Array_Eintrag[0][8]);   // Setze Tel
        }

        Sheet_Eintrag.getRange("B" + Zeile +":K" + Zeile).setValue("");
        Sheet_Eintrag.getRange("H" + Zeile +":J" + Zeile).setValues([["=D"+Zeile,"=E"+Zeile,"=F"+Zeile]]);

        if(Sheet_FIB.getRange("B" + Zeile).getValue() == Array_Eintrag[0][0])
        {
          Sheet_FIB.getRange("B" + Zeile +":K" + Zeile).setValue("");
          Sheet_FIB.getRange("H" + Zeile +":J" + Zeile).setValues([["=D"+Zeile,"=E"+Zeile,"=F"+Zeile]]);
        }
        else
        {
          var Array_FIB =  Sheet_FIB.getRange("B7:B").getValues();

          for(var y = 0; y < Array_FIB.length; y++)
          {
            if(Array_FIB[y][0] == Array_Eintrag[0][0])
            {
              Sheet_FIB.getRange("B" + (y+7) +":K" + (y+7)).setValue("");
              Sheet_FIB.getRange("H" + (y+7) +":J" + (y+7)).setValues([["=D"+(y+7),"=E"+(y+7),"=F"+(y+7)]]);
              break;
            }
          }
        }

        if(Sheet_Akten.getRange("B" + Zeile).getValue() == Array_Eintrag[0][0])
        {
          Sheet_Akten.getRange("B" + Zeile +":K" + Zeile).setValue("");
          Sheet_Akten.getRange("H" + Zeile +":J" + Zeile).setValues([["=D"+Zeile,"=E"+Zeile,"=F"+Zeile]]);
        }
        else
        {
          var Array_Akten =  Sheet_Akten.getRange("B7:B").getValues();

          for(var y = 0; y < Array_Akten.length; y++)
          {
            if(Array_Akten[y][0] == Array_Eintrag[0][0])
            {
              Sheet_Akten.getRange("B" + (y+7) +":K" + (y+7)).setValue("");
              Sheet_Akten.getRange("H" + (y+7) +":J" + (y+7)).setValues([["=D"+(y+7),"=E"+(y+7),"=F"+(y+7)]]);
              break;
            }
          }
        }

        break;
      }
    }
  }

  else if(SheetName == "Akteneinträge" && Spalte == Spalte_in_Index("J") && Zeile == 3)
  {
    var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Akteneinträge");
    var Sheet_Akten = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle).getSheetByName("Akteneinträge");

    Sheet_FIB.getRange("J3").setValue(Value);
    Sheet_Akten.getRange("J3").setValue(Value);
  }

  else if(SheetName == "Akteneinträge" && Spalte >= Spalte_in_Index("H") && Spalte <= Spalte_in_Index("J") && Zeile >= 7)
  {
    var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Akteneinträge");
    var Sheet_Akten = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle).getSheetByName("Akteneinträge");
    var Sheet_Eintrag = SpreadsheetApp.getActive().getSheetByName("Akteneinträge");

    if(Value == "" || Value == undefined)
    {
      Sheet_FIB.getRange(Zeile,Spalte).setFormula(Sheet_Eintrag.getRange(Zeile,Spalte).getFormula());
      Sheet_Akten.getRange(Zeile,Spalte).setFormula(Sheet_Eintrag.getRange(Zeile,Spalte).getFormula());
    }
    else
    {
      Sheet_FIB.getRange(Zeile,Spalte).setValue(Value);
      Sheet_Akten.getRange(Zeile,Spalte).setValue(Value);
    }
  }

  //-------------------------------------------- ENDE ------------------------------------------//


  //------------------------------------- Mitgliederlisten -------------------------------------//

  else if(SheetName == "Mitgliederlisten" && Spalte == Spalte_in_Index("F") && Zeile >= 10 && Value == "TRUE")   // Eintrag in Datenbank
  {
    var Sheet_Mitgliederliste = SpreadsheetApp.getActive().getSheetByName("Mitgliederlisten");
    
    var Array_Header = Sheet_Mitgliederliste.getRange("B5:F5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("B" + Zeile + ":E" + Zeile).getValues();

    var Array_Eingabe = [[Array_Zeile[0][0],Array_Header[0][0],Array_Zeile[0][1],Array_Zeile[0][2],Array_Zeile[0][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]]];

    var Error = Datenbank_Eintrag(Array_Eingabe,Off_Name,Off_Fraktion,Off_Rang,Off_Haus,Off_Tel,Off_Aktivitaet,Off_Beamter,Off_Akteneintrag);
  
    if(Error == 1)    // Wenn falsche Eingaben Stop
    {
      Sheet_Mitgliederliste.getRange("F" + Zeile).setValue("");
      Stop++;
      return 0;
    }
    else if(Error == 2)   // Wenn unerwarteter Fehler im Script
    {
      Stop++;
      return 0;
    }

    var Rang     = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$D;3;false);))';
    var HN       = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$E;4;false);))';
    var Tel      = '=IF($B' + Zeile + ' = "";"";IFERROR(VLOOKUP($B' + Zeile + ';\'Import Aktuell\'!$B$3:$F;5;false);))';

    Sheet_Mitgliederliste.getRange("B" + Zeile + ":F" + Zeile).setValues([["",Rang,HN,Tel,false]]);

    return 0;
  }
  else if(SheetName == "Mitgliederlisten" && Spalte == Spalte_in_Index("L") && Zeile >= 10 && Value == "TRUE")   // Umtragen auf Zivi
  {
    var Sheet_Mitgliederliste = SpreadsheetApp.getActive().getSheetByName("Mitgliederlisten");
    
    var Array_Header = Sheet_Mitgliederliste.getRange("H5:L5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("H" + Zeile + ":K" + Zeile).getValues();

    var Array_Eingabe = [[Array_Zeile[0][0],Array_Header[0][0],Array_Zeile[0][1],Array_Zeile[0][2],Array_Zeile[0][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]]];

    var Error = Datenbank_Eintrag(Array_Eingabe,Off_Name,Off_Fraktion,Off_Rang,Off_Haus,Off_Tel,Off_Aktivitaet,Off_Beamter,Off_Akteneintrag);
  
    if(Error == 1)    // Wenn falsche Eingaben Stop
    {
      Sheet_Mitgliederliste.getRange("L" + Zeile).setValue("");
      Stop++;
      return 0;
    }
    else if(Error == 2)   // Wenn unerwarteter Fehler im Script
    {
      Stop++;
      return 0;
    }

    Sheet_Mitgliederliste.getRange("L" + Zeile).setValue("");

    return 0;
  }
  else if(SheetName == "Mitgliederlisten" && Spalte == Spalte_in_Index("F") && Zeile == 7 && Value == "TRUE")   // Alle aus Mitgliederliste in DB eintragen
  {
    var Sheet_Mitgliederliste = SpreadsheetApp.getActive().getSheetByName("Mitgliederlisten");
    
    var Letzte_Zeile = Sheet_Mitgliederliste.getRange("B8").getValue();

    Sheet_Mitgliederliste.getRange("B10:F" + Letzte_Zeile).sort(Spalte_in_Index("B"));
    SpreadsheetApp.flush();
    
    var Array_Header = Sheet_Mitgliederliste.getRange("B5:F5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("B10:E" + Letzte_Zeile).getValues();
    var Array_Eingabe = new Array();

    for(var i = 0; i < Array_Zeile.length; i++)
    {
      if(Array_Zeile[i][0] != "")
      {
        Array_Eingabe.push([Array_Zeile[i][0],Array_Header[0][0],Array_Zeile[i][1],Array_Zeile[i][2],Array_Zeile[i][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]])
      }
    }

    for(var i = 0; i < Array_Eingabe.length; i++)
    {
      var Error = Datenbank_Eintrag([Array_Eingabe[i]],Off_Name,Off_Fraktion,Off_Rang,Off_Haus,Off_Tel,Off_Aktivitaet,Off_Beamter,Off_Akteneintrag);
    
      if(Error != 1 && Error != 2)
      {
        var Rang     = '=IF($B' + (i+10) + ' = "";"";IFERROR(VLOOKUP($B' + (i+10) + ';\'Import Aktuell\'!$B$3:$D;3;false);))';
        var HN       = '=IF($B' + (i+10) + ' = "";"";IFERROR(VLOOKUP($B' + (i+10) + ';\'Import Aktuell\'!$B$3:$E;4;false);))';
        var Tel      = '=IF($B' + (i+10) + ' = "";"";IFERROR(VLOOKUP($B' + (i+10) + ';\'Import Aktuell\'!$B$3:$F;5;false);))';

        Sheet_Mitgliederliste.getRange("B" + (i+10) + ":F" + (i+10)).setValues([["",Rang,HN,Tel,false]]);
      }
    }

    Sheet_Mitgliederliste.getRange("F7").setValue(false);
    return 0;
  }
  else if(SheetName == "Mitgliederlisten" && Spalte == Spalte_in_Index("L") && Zeile == 7 && Value == "TRUE")   // Alle aus Mitgliederliste die Rot sind in DB auf Zivi eintragen
  {
    var Sheet_Mitgliederliste = SpreadsheetApp.getActive().getSheetByName("Mitgliederlisten");
    
    var Letzte_Zeile = Sheet_Mitgliederliste.getRange("H8").getValue();
    
    var Array_Header = Sheet_Mitgliederliste.getRange("H5:L5").getValues();
    var Array_Zeile = Sheet_Mitgliederliste.getRange("H10:M" + Letzte_Zeile).getValues();
    var Array_Eingabe = new Array();

    for(var i = 0; i < Array_Zeile.length; i++)
    {
      if(Array_Zeile[i][0] != "" && Array_Zeile[i][5] == false)
      {
        Array_Eingabe.push([Array_Zeile[i][0],Array_Header[0][0],Array_Zeile[i][1],Array_Zeile[i][2],Array_Zeile[i][3],Array_Header[0][1],Array_Header[0][2],Array_Header[0][3],Array_Header[0][4]])
      }
    }

    Logger.log(Array_Eingabe);

    for(var i = 0; i < Array_Eingabe.length; i++)
    {
      var Error = Datenbank_Eintrag([Array_Eingabe[i]],Off_Name,Off_Fraktion,Off_Rang,Off_Haus,Off_Tel,Off_Aktivitaet,Off_Beamter,Off_Akteneintrag);
    }

    Sheet_Mitgliederliste.getRange("L7").setValue(false);
    return 0;
  }

  //-------------------------------------------- ENDE ------------------------------------------//


  //------------------------------------- Namensänderungen -------------------------------------//

  else if(SheetName == "Namensänderungen" && Spalte == Spalte_in_Index("F") && Zeile == 5 && Value == "TRUE")   // Namensänderung
  {
    var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);
    var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");
    var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");
    var Sheet_Namensaenderung = SpreadsheetApp.getActive().getSheetByName("Namensänderungen");

    var Array_Eingabe = Sheet_Namensaenderung.getRange("B5:E5").getValues();

    var Array_Ausgabe = Eintrag_Check([[Array_Eingabe[0][1],Array_Eingabe[0][2]]],0,-1,-1,1,true);

    var Num1 = Sheet_Aktuell.createTextFinder(Array_Eingabe[0][0]).matchEntireCell(true).replaceAllWith(Array_Ausgabe[0][0]);
    var Num2 = Sheet_Archiv.createTextFinder(Array_Eingabe[0][0]).matchEntireCell(true).replaceAllWith(Array_Ausgabe[0][0]);
    
    SpreadsheetApp.flush();

    Archivieren_Name(Array_Ausgabe[0][0]);

    Sheet_Namensaenderung.insertRowBefore(10);

    Sheet_Namensaenderung.getRange("B10:F10").setValues([[Array_Eingabe[0][0],Array_Ausgabe[0][0],Array_Eingabe[0][2],Array_Eingabe[0][3],(Num1 + Num2)]]);

    Sheet_Namensaenderung.getRange("B5:F5").setValue("");
  }

  //-------------------------------------------- ENDE ------------------------------------------//



  //------------------------------------- Zugehörigkeit -------------------------------------//

  else if(SheetName == "Zugehörigkeit")
  {
    var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Zugehörigkeit");
    var Sheet_PD = SpreadsheetApp.getActive().getSheetByName("Zugehörigkeit");
    
    if(Spalte == Spalte_in_Index("J") && Zeile >= 4 && Value != "" && Value != undefined)
    {
      Sheet_FIB.getRange(Zeile,Spalte).setValue(Value);
      Sheet_FIB.getRange("K" + Zeile + ":L" + Zeile).setValues([[new Date(), "LSPD"]]);
    }
    else if(Spalte == Spalte_in_Index("J") && Zeile >= 4 && (Value == "" || Value == undefined))
    {
      Sheet_FIB.getRange(Zeile,Spalte).setValue(Value);
      Sheet_FIB.getRange("K" + Zeile + ":L" + Zeile).setValue("");
    }
    else if(Spalte == Spalte_in_Index("N") && Zeile >= 4 && Value == "TRUE")
    {
      Sheet_PD.getRange(Zeile,Spalte).setValue("");

      Sheet_PD.getFilter().setColumnFilterCriteria(Spalte_in_Index("J"),SpreadsheetApp.newFilterCriteria().setHiddenValues(["Abgelehnt","Angenommen"]).build())
      Sheet_FIB.getFilter().setColumnFilterCriteria(Spalte_in_Index("J"),SpreadsheetApp.newFilterCriteria().setHiddenValues(["Abgelehnt","Angenommen"]).build())
    }
    else
    {
      Sheet_FIB.getRange(Zeile,Spalte).setValue(Value);
    }
  }

  //-------------------------------------------- ENDE ------------------------------------------//

}

function Datenbank_Eintrag(Array_Eingabe, Off_Name = 0, Off_Fraktion = 1, Off_Rang = 2, Off_Haus = 3, Off_Tel = 4, Off_Aktivitaet = 7, Off_Beamter = 8, Off_Akteneintrag = 9, Off_Datenbank = 10)
{
  try
  {
    var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);
    var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");
    var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");

    var Array_Ausgabe = Eintrag_Check(Array_Eingabe,Off_Name,Off_Fraktion,Off_Aktivitaet,Off_Beamter,false);

    if(Array_Ausgabe == 1)    // Wenn falsche Eingaben Stop
    {
      Logger.log("Stop weil Eingabe Fehler");
      return 1;
    }

    var Array_DB_Aktuell = Sheet_Aktuell.getRange("B3:B").getValues();
    var Such_Zeile = 0;

    for(var y = 0; y < Array_DB_Aktuell.length; y++)  // Suche nach bestehenden Eintrag in Datenbank
    {
      if(Array_DB_Aktuell[y][Off_Name] == Array_Ausgabe[0][Off_Name])
      {
        Such_Zeile = y + 3;
        Logger.log("Person: " + Array_DB_Aktuell[y][Off_Name] + " in Datenbank Zeile " + Such_Zeile + " gefunden");
        break;
      }
    }

    if(Such_Zeile == 0)     // Neuer Datenbank Eintrag
    {
      Logger.log("Person: " + Array_Ausgabe[0][Off_Name] + " in Datenbank nicht gefunden. Erstelle neuen Eintrag");

      Array_Eingabe[0].unshift("");

      Sheet_Aktuell.appendRow(Array_Eingabe[0]);
    }
    else                  // Überschreibe und Archiviere Datenbank Eintrag
    {
      var Array_Gefunden = Sheet_Aktuell.getRange("B" + Such_Zeile + ":K" + Such_Zeile).getValues();
      var Aktivitaet_plus_3Stunden = new Date(Array_Gefunden[0][Off_Aktivitaet]);
      Aktivitaet_plus_3Stunden.setHours(Aktivitaet_plus_3Stunden.getHours() + 3);

      if(Array_Gefunden[0][Off_Fraktion] == Array_Ausgabe[0][Off_Fraktion] && Array_Gefunden[0][Off_Rang] == Array_Ausgabe[0][Off_Rang] && Array_Gefunden[0][Off_Haus] == Array_Ausgabe[0][Off_Haus] && Array_Gefunden[0][Off_Tel] == Array_Ausgabe[0][Off_Tel] && Aktivitaet_plus_3Stunden >= Array_Ausgabe[0][Off_Aktivitaet])   // Prüfe ob innerhalb von 3 Stunden der Eintrag schon eingetragen wurde
      {
        Logger.log("Eintrag würd nicht überschrieben da der selbe Eintrag innerhalb der letzten 3 Stunden schon eingetragen wurde\n" + Array_Gefunden);
      }
      else
      {
        if(Array_Gefunden[0][Off_Akteneintrag] != "" && Array_Gefunden[0][Off_Akteneintrag] != undefined)
        {
          var Akteneintrag_plus_30Tage = new Date(Array_Gefunden[0][Off_Akteneintrag]);
          Akteneintrag_plus_30Tage.setDate(Akteneintrag_plus_30Tage.getDate() + 14);

          if(Array_Gefunden[0][Off_Fraktion] == Array_Ausgabe[0][Off_Fraktion] && Akteneintrag_plus_30Tage <= Array_Ausgabe[0][Off_Aktivitaet] || Array_Gefunden[0][Off_Fraktion] != Array_Ausgabe[0][Off_Fraktion])  // Lösche Akteneintrag weil älter als 30 Tage her Oder Andere Fraktion
          {
            Logger.log("Lösche Akteneintrag weil älter als 30 Tage her Oder Andere Fraktion");
            Sheet_Aktuell.getRange("K" + Such_Zeile).setValue(""); 
          }
        }

        Logger.log("Überschreibe Eintrag und Archiviere ihn\n" + Array_Gefunden);

        if(Array_Ausgabe[0][Off_Fraktion] == "Zivilist")
        {
          Array_Ausgabe[0][Off_Rang] = "";
        }
        
        Sheet_Aktuell.getRange("B" + Such_Zeile + ":J" + Such_Zeile).setValues(Array_Ausgabe);  // Überschreibe Eintrag

        Array_Gefunden[0].unshift("");

        Sheet_Archiv.appendRow(Array_Gefunden[0])   // Archiviere alten Eintrag
      }
    }

    return 0;
  }
  catch(err)
  {
    Logger.log(err.stack);
    return 2;
  }
}

function Archivieren_Name(Name)
{
  var Off_Name = 0, Off_Fraktion = 1, Off_Rang = 2, Off_Haus = 3, Off_Tel = 4, Off_Aktivitaet = 7, Off_Beamter = 8, Off_Akteneintrag = 9, Off_Datenbank = 10;

  var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);
  var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");
  var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");

  var Array_DB_Aktuell = Sheet_Aktuell.getRange("B3:K" + Sheet_Aktuell.getLastRow()).getValues();

  var Gefundene_Array = new Array();
  var Gefundene_Zeilen = new Array();

  for(var i = 0; i < Array_DB_Aktuell.length; i++)
  {
    if(Array_DB_Aktuell[i][0] == Name)
    {
      Logger.log("Gefunden " + Array_DB_Aktuell[i]);

      Gefundene_Array.push(Array_DB_Aktuell[i]);
      Gefundene_Zeilen.push(i + 3);
    }
  }
    
  if(Gefundene_Array.length > 1)
  {
    var Array_Groß = Gefundene_Array[0];

    for(var i = 1; i < Gefundene_Array.length; i++)
    {
      if(Array_Groß[Off_Aktivitaet] < Gefundene_Array[i][Off_Aktivitaet])
      {
        Array_Groß = Gefundene_Array[i];
      }
    }

    for(var i = Gefundene_Array.length-1; i >= 0 ; i--)
    {
      if(Gefundene_Array[i][Off_Aktivitaet] != Array_Groß[Off_Aktivitaet])
      {
        Logger.log("Archiviere eintrag in Zeile: " + Gefundene_Zeilen[i] +"\n" + Gefundene_Array[i]);
        Gefundene_Array[i].unshift("");
        Sheet_Archiv.appendRow(Gefundene_Array[i]);
        Sheet_Aktuell.deleteRow(Gefundene_Zeilen[i]);
      }
    }
  }
}