function Namen()
{
  Update_GOV_Namen();

  SpreadsheetApp.flush();

  
  var Sheet_Namen = SpreadsheetApp.getActive().getSheetByName("Namensänderungen");
  var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);
  var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");
  var Sheet_Archiv = SS_Datenbank.getSheetByName("Archiv");

  var Array_Namen = Sheet_Namen.getRange("B3:E" + Sheet_Namen.getLastRow()).getValues();

  for(var i = 0; i < Array_Namen.length; i++)
  {
    if(Array_Namen[i][3] != true && Array_Namen[i][0] != "")
    {
      var Alter_Name = Eintrag_Check([Array_Namen[i]],0,-1,-1,-1,false,"Warning");
      var Neuer_Name = Eintrag_Check([Array_Namen[i]],1,-1,-1,-1,false,"Warning");

      if(Alter_Name != 1 && Neuer_Name != 1)
      {
        var Anzahl = Sheet_Aktuell.createTextFinder(Alter_Name[0]).matchEntireCell(true).replaceAllWith(Neuer_Name[1]);

        Anzahl += Sheet_Archiv.createTextFinder(Alter_Name[0]).matchEntireCell(true).replaceAllWith(Neuer_Name[1]);

        if(Anzahl > 0)
        {
          SpreadsheetApp.flush();

          Archivieren_Name(Neuer_Name[1]);
        }

        Sheet_Namen.getRange("E" + (i+3) + ":F" + (i+3)).setValues([[true,Anzahl]]);
      }
      else
      {
        Sheet_Namen.getRange("E" + (i+3) + ":F" + (i+3)).setValues([[false,0]]);
      }
    }
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

function Update_GOV_Namen()
{
  var Sheet_LSPD = SpreadsheetApp.getActive().getSheetByName("Namensänderungen");
  var Sheet_GOV = SpreadsheetApp.openById("1PjPiOAj2k_ADq6G-qCTytmTKmifD59ArL3DRY02KicQ").getSheetByName("Namensänderungen");

  var Array_LSPD = Sheet_LSPD.getRange("B3:D" + Sheet_LSPD.getLastRow()).getValues();
  var Array_GOV = Sheet_GOV.getRange("B2:D" + Sheet_GOV.getLastRow()).getValues();
  
  var Array_Ausgabe = new Array();

  for(var i = 0; i < Array_GOV.length; i++)
  {
    var Gefunden = false;

    for(var y = 0; y < Array_LSPD.length; y++)
    {
      if(Array_GOV[i].toString() == Array_LSPD[y].toString())
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      Logger.log("Add: " + Array_GOV[i][0] + "  -->  " + Array_GOV[i][1]);

      Array_Ausgabe.push([Array_GOV[i][0],Array_GOV[i][1],Array_GOV[i][2],"","",new Date()]);
    }
  }

  if(Array_Ausgabe.length > 0)
  {
    Logger.log(Array_Ausgabe);
    Sheet_LSPD.getRange("B" + (Sheet_LSPD.getLastRow() + 1) + ":G" + (Sheet_LSPD.getLastRow() + Array_Ausgabe.length)).setValues(Array_Ausgabe);
  }
  else
  {
    Logger.log("Keine Neuen Namen vom GOV");
  }
}