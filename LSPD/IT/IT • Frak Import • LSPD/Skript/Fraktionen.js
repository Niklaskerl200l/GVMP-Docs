var Array_Eintragen = new Array();
var Array_Update = new Array();

function Fraktionen()
{
  if(Bezug_Check() == false)
  {
    try
    {
      var Fehler = false;

      var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
      var SS_Schnittstelle = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);
      var Sheet_Aktuell = SS_Schnittstelle.getSheetByName("Aktuell");
      var Sheet_Archiv = SS_Schnittstelle.getSheetByName("Archiv");

      var Array_Aktuell = Sheet_Aktuell.getRange("B3:F" + Sheet_Aktuell.getLastRow()).getValues();
      var Array_Auswertung = Sheet_Auswertung.getRange("B4:E" + Sheet_Auswertung.getLastRow()).getValues();  

      var Array_LSPD = new Array();
      var Array_LSMC = new Array();
      var Array_ARMY = new Array();
      var Array_GOV = new Array();
      var Array_DMV = new Array();
      var Array_DPOS = new Array();
      var Array_WN = new Array();
      var Array_FIB = new Array();

      var Array_DB_LSPD = new Array();
      var Array_DB_LSMC = new Array();
      var Array_DB_ARMY = new Array();
      var Array_DB_GOV = new Array();
      var Array_DB_DMV = new Array();
      var Array_DB_DPOS = new Array();
      var Array_DB_WN = new Array();
      var Array_DB_FIB = new Array();

      for(var y = 0; y < Array_Auswertung.length; y++)
      {
        switch(Array_Auswertung[y][1])
        {
          case "LSPD" : Array_LSPD.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning")); break;
          case "LSMC" : Array_LSMC.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning")); break;
          case "ARMY" : Array_ARMY.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning")); break;
          case "GOV"  : Array_GOV.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning"));  break;
          case "DMV"  : Array_DMV.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning"));  break;
          case "DPOS" : Array_DPOS.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning")); break;
          case "WN" : Array_WN.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning")); break;
          case "FIB" : Array_FIB.push(Eintrag_Check(Array_Auswertung[y],0,1,-1,-1,false,"Warning")); break;
        }
      }

      for(var y = 0; y < Array_Aktuell.length; y++)
      {
        switch(Array_Aktuell[y][1])
        {
          case "LSPD" : Array_DB_LSPD.push(Array_Aktuell[y]); break;
          case "LSMC" : Array_DB_LSMC.push(Array_Aktuell[y]); break;
          case "ARMY" : Array_DB_ARMY.push(Array_Aktuell[y]); break;
          case "GOV"  : Array_DB_GOV.push(Array_Aktuell[y]);  break;
          case "DMV"  : Array_DB_DMV.push(Array_Aktuell[y]);  break;
          case "DPOS" : Array_DB_DPOS.push(Array_Aktuell[y]); break;
          case "WN" : Array_DB_WN.push(Array_Aktuell[y]); break;
          case "FIB" : Array_DB_FIB.push(Array_Aktuell[y]); break;
        }
      }
      Logger.log("\n------------------------------ LSPD ------------------------------\n");

      Fraktion_Check(Array_LSPD,Array_DB_LSPD,Array_Aktuell);

      Logger.log("\n------------------------------ LSMC ------------------------------\n");

      Fraktion_Check(Array_LSMC,Array_DB_LSMC,Array_Aktuell);

      Logger.log("\n------------------------------ ARMY ------------------------------\n");
      
      Fraktion_Check(Array_ARMY,Array_DB_ARMY,Array_Aktuell);

      Logger.log("\n------------------------------ GOV ------------------------------\n");
      
      Fraktion_Check(Array_GOV,Array_DB_GOV,Array_Aktuell);

      Logger.log("\n------------------------------ DMV ------------------------------\n");
      
      Fraktion_Check(Array_DMV,Array_DB_DMV,Array_Aktuell);

      Logger.log("\n------------------------------ DPOS ------------------------------\n");
      
      Fraktion_Check(Array_DPOS,Array_DB_DPOS,Array_Aktuell)

      Logger.log("\n------------------------------ WN ------------------------------\n");
      
      Fraktion_Check(Array_WN,Array_DB_WN,Array_Aktuell);

      Logger.log("\n------------------------------ FIB ------------------------------\n");
      
      Fraktion_Check(Array_FIB,Array_DB_FIB,Array_Aktuell);

      Logger.log("\n------------------------------ Eintragung ------------------------------\n");

      //----------------- Eintragen ------------------------//

      SpreadsheetApp.flush();

      Array_Aktuell = Sheet_Aktuell.getRange("B3:F" + Sheet_Aktuell.getLastRow()).getValues();

      for(var i = 0; i < Array_Aktuell.length; i++)     // Updaten DB
      {
        for(var x = 0; x < Array_Update.length; x++)
        {
          if(Array_Aktuell[i][0] == Array_Update[x][0] && Array_Aktuell[i][1] == Array_Update[x][1])
          {
            if(Array_Update[x][2] != "")
            {
              Logger.log("Update Rang in Zeile " + (i+3) + " bei " + Array_Aktuell[i][0] + " von " + Array_Aktuell[i][2] + " zu " + Array_Update[x][2]);

              Sheet_Aktuell.getRange("D" + (i+3)).setValue(Array_Update[x][2]);
            }
            
            if(Array_Update[x][3] != "")
            {
              Logger.log("Update Tel in Zeile " + (i+3) + " bei " + Array_Aktuell[i][0] + " von " + Array_Aktuell[i][4] + " zu " + Array_Update[x][3]);

              Sheet_Aktuell.getRange("F" + (i+3)).setValue(Array_Update[x][3]);
            }
          }
        }
      }

      Array_Eintragen.sort(         // Sortieren
        function(a,b)
        {
          if(a[1].toLowerCase() > b[1].toLowerCase())
          {
            return -1;
          }
          else if(a[1].toLowerCase() < b[1].toLowerCase())
          {
            return 1;
          }
        }
      );

      Logger.log(Array_Eintragen);

      Logger.log("\n------------------------------ Eintragen ------------------------------\n");

      for(var i = 0; i < Array_Eintragen.length; i++)   // DB Eintragen
      {
        if(Datenbank_Eintrag(Array_Eintragen[i]) == 2) Fehler = true;
      }

      if(Fehler == true)
      {
        throw Error("Error");
      }
    }
    catch(err)
    {
      Logger.log(err.stack);

      MailApp.sendEmail("1@1.de","LSPD Frak Import Fehler","Fehler: " + err.stack);
      
      throw Error("Error");
    }
  }
}

function Fraktion_Check(Array_Frak,Array_DB_Frak,Array_Aktuell)
{
  var Off_Name = 0, Off_Fraktion = 1, Off_Rang = 2, Off_Tel = 4;

  if(Array_Frak != null && Array_Frak != undefined && Array_Frak.length >= 1 && Array_DB_Frak != null && Array_DB_Frak != undefined && Array_DB_Frak.length >= 0)
  {
    // Update bestehender und Eintrag Neuer

    for(var i = 0; i < Array_Frak.length; i++)
    {
      var Gefunden = false;

      for(var x = 0; x < Array_DB_Frak.length; x++)
      {
        if(Array_Frak[i][0] == Array_DB_Frak[x][Off_Name])    // Name Gefunden
        {
          if(Array_Frak[i][2] != Array_DB_Frak[x][Off_Rang])  // Anderer Rang
          {
            var temp = [Array_Frak[i][0],Array_Frak[i][1],Array_Frak[i][2],""]

            Logger.log("Neuer Rang: " + temp);

            Array_Update.push(temp);
          }

          if(Array_Frak[i][3] != Array_DB_Frak[x][Off_Tel] && Array_Frak[i][3] != "")  // Anderer Tel
          {
            var temp =[Array_Frak[i][0],Array_Frak[i][1],"",Array_Frak[i][3]]

            Logger.log("Neue Tel: " + temp);

            Array_Update.push(temp);
          }

          Gefunden = true;
          break;
        }
      }

      if(Gefunden == false)
      {
        if( !((Array_Frak[i][1] == "LSMC" && Array_Frak[i][2] == 0) || (Array_Frak[i][1] == "ARMY" && Array_Frak[i][2] == 0) || Array_Frak[i][1] == "WN" || Array_Frak[i][1] == "FIB") )   // Neu Einstellung
        {
          var temp = [Array_Frak[i][0],Array_Frak[i][1],Array_Frak[i][2],"",Array_Frak[i][3],"Automatischer Eintrag","Mitarbeiterliste",new Date(),"LSPD Bot"]

          Logger.log("Neu Einstellung: " + temp);

          Array_Eintragen.push(temp);
        }
        else
        {
          for(var z = 0; z < Array_Aktuell.length; z++)
          {
            if(Array_Aktuell[z][Off_Name] == Array_Frak[i][0] && Array_Aktuell[z][Off_Fraktion] != "Zivilist")
            {
              var Datum = new Date();
              Datum.setMinutes(Datum.getMinutes() - 2,0,0);

              var temp = [Array_Frak[i][0],"Zivilist","","",Array_Frak[i][3],"Automatischer Eintrag","Mitarbeiterliste",Datum,"LSPD Bot"]

              Logger.log("GWD / ZD / WN: " + temp);

              Array_Eintragen.push(temp);

              break;
            }
          }
        }
      }
    }


    // Austragen von Mitarbeitern


    for(var i = 0; i < Array_DB_Frak.length; i++)
    {
      var Gefunden = false;

      for(var x = 0; x < Array_Frak.length; x++)
      {
        if(Array_DB_Frak[i][0] == Array_Frak[x][Off_Name])    // Name Gefunden
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == false)   // Entlassen
      {
        var Datum = new Date();
        Datum.setMinutes(Datum.getMinutes() - 2,0,0);

        var temp = [Array_DB_Frak[i][0],"Zivilist","",Array_DB_Frak[i][3],Array_DB_Frak[i][4],"Automatischer Eintrag","Mitarbeiterliste",Datum,"LSPD Bot"]

        Logger.log("Entlassung: " + temp);

        Array_Eintragen.push(temp);
      }
    }
  }
}

function Datenbank_Eintrag(Array_Eingabe, Off_Name = 0, Off_Fraktion = 1, Off_Rang = 2, Off_Haus = 3, Off_Tel = 4, Off_Aktivitaet = 7, Off_Beamter = 8, Off_Akteneintrag = 9, Off_Datenbank = 10)
{
  try
  {
    var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);
    var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");
    var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");

    var Array_Ausgabe = Eintrag_Check(Array_Eingabe,Off_Name,Off_Fraktion,Off_Aktivitaet,Off_Beamter,false,"Error");

    if(Array_Ausgabe == 1)    // Wenn falsche Eingaben Stop
    {
      Logger.log("Stop weil Eingabe Fehler");
      return 1;
    }

    var Array_DB_Aktuell = Sheet_Aktuell.getRange("B3:B").getValues();
    var Such_Zeile = 0;

    for(var y = 0; y < Array_DB_Aktuell.length; y++)  // Suche nach bestehenden Eintrag in Datenbank
    {
      if(Array_DB_Aktuell[y][Off_Name] == Array_Ausgabe[Off_Name])
      {
        Such_Zeile = y + 3;
        Logger.log("Person: " + Array_DB_Aktuell[y][Off_Name] + " in Datenbank Zeile " + Such_Zeile + " gefunden");
        break;
      }
    }

    if(Such_Zeile == 0)     // Neuer Datenbank Eintrag
    {
      Logger.log("Person: " + Array_Ausgabe[Off_Name] + " in Datenbank nicht gefunden. Erstelle neuen Eintrag");

      Array_Eingabe.unshift("");

      Sheet_Aktuell.appendRow(Array_Eingabe);
    }
    else                  // Überschreibe und Archiviere Datenbank Eintrag
    {
      var Array_Gefunden = Sheet_Aktuell.getRange("B" + Such_Zeile + ":K" + Such_Zeile).getValues();
      var Aktivitaet_plus_3Stunden = new Date(Array_Gefunden[0][Off_Aktivitaet]);
      Aktivitaet_plus_3Stunden.setHours(Aktivitaet_plus_3Stunden.getHours() + 3);

      if(Array_Gefunden[0][Off_Fraktion] == Array_Ausgabe[Off_Fraktion] && Array_Gefunden[0][Off_Rang] == Array_Ausgabe[Off_Rang] && Array_Gefunden[0][Off_Haus] == Array_Ausgabe[Off_Haus] && Array_Gefunden[0][Off_Tel] == Array_Ausgabe[Off_Tel] && Aktivitaet_plus_3Stunden >= Array_Ausgabe[Off_Aktivitaet])   // Prüfe ob innerhalb von 3 Stunden der Eintrag schon eingetragen wurde
      {
        Logger.log("Eintrag würd nicht überschrieben da der selbe Eintrag innerhalb der letzten 3 Stunden schon eingetragen wurde\n" + Array_Gefunden);
      }
      else
      {
        if(Array_Gefunden[0][Off_Akteneintrag] != "" && Array_Gefunden[0][Off_Akteneintrag] != undefined)
        {
          var Akteneintrag_plus_30Tage = new Date(Array_Gefunden[0][Off_Akteneintrag]);
          Akteneintrag_plus_30Tage.setDate(Akteneintrag_plus_30Tage.getDate() + 30);

          if(Array_Gefunden[0][Off_Fraktion] == Array_Ausgabe[Off_Fraktion] && Akteneintrag_plus_30Tage <= Array_Ausgabe[Off_Aktivitaet] || Array_Gefunden[0][Off_Fraktion] != [Off_Fraktion])  // Lösche Akteneintrag weil älter als 30 Tage her Oder Andere Fraktion
          {
            Logger.log("Lösche Akteneintrag weil älter als 30 Tage her Oder Andere Fraktion");
            Sheet_Aktuell.getRange("K" + Such_Zeile).setValue(""); 
          }
        }

        Logger.log("Überschreibe Eintrag und Archiviere ihn\n" + Array_Gefunden);

        if(Array_Ausgabe[Off_Fraktion] == "Zivilist")
        {
          Array_Ausgabe[Off_Rang] = "";
        }

        if(Array_Ausgabe[Off_Haus] == "" && Array_Gefunden[0][Off_Haus] != "")
        {
          Array_Ausgabe[Off_Haus] = Array_Gefunden[0][Off_Haus];
        }

        if(Array_Ausgabe[Off_Tel] == "" && Array_Gefunden[0][Off_Tel] != "")
        {
          Array_Ausgabe[Off_Tel] = Array_Gefunden[0][Off_Tel];
        }
        
        Sheet_Aktuell.getRange("B" + Such_Zeile + ":J" + Such_Zeile).setValues([Array_Ausgabe]);  // Überschreibe Eintrag

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

function Bezug_Check()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();

  var LSPD = SS.getSheetByName("Import LSPD").getRange("B3").getValue();
  //var FIB = SS.getSheetByName("Import FIB").getRange("B3").getValue();
  var LSMC = SS.getSheetByName("Import LSMC").getRange("B3").getValue();
  //var Army = SS.getSheetByName("Import Army").getRange("B3").getValue();
  var Gov = SS.getSheetByName("Import GOV").getRange("C3").getValue();
  var DMV = SS.getSheetByName("Import DMV").getRange("B3").getValue();
  var DPOS = SS.getSheetByName("Import DPOS").getRange("B3").getValue();
  var WN = SS.getSheetByName("Import WN").getRange("B3").getValue();

  if
  (
    (LSPD == "#REF!" || LSPD == "#ERROR!" || LSPD == "#NV!") ||
    //(FIB == "#REF!" || FIB == "#ERROR!" || FIB == "#NV!") ||
    (LSMC == "#REF!" || LSMC == "#ERROR!" || LSMC == "#NV!") ||
    //(Army == "#REF!" || Army == "#ERROR!" || Army == "#NV!") ||
    (Gov == "#REF!" || Gov == "#ERROR!" || Gov == "#NV!") ||
    (DMV == "#REF!" || DMV == "#ERROR!" || DMV == "#NV!") ||
    (DPOS == "#REF!" || DPOS == "#ERROR!" || DPOS == "#NV!") ||
    (WN == "#REF!" || WN == "#ERROR!" || WN == "#NV!")
  )
  {
    Logger.log("Bezugs Fehler")
    MailApp.sendEmail("1@1.de","LSPD Frak Import BEZUG Fehler Import broke","Fehler");
    return true;
  }
  return false;
}