function Selbstzerstörung()
{
  var Array_Sheets = ["Formulare","Lesen"]; //"LSPD","Direction","Detective","Recruitment","Training","SOC","PRU","GTF","WLD","IT","IT",

  var ID_Zeile = 3;

  for(var z = 0; z < Array_Sheets.length; z++)
  {
    var Sheet = SpreadsheetApp.getActive().getSheetByName(Array_Sheets[z]);

    var Anzahl_Sheets = Sheet.getRange("A2").getValue();
    var Spalte = 2;

    for(var z2 = 1; z2 <= Anzahl_Sheets; z2++)
    {
      var ID = Sheet.getRange(ID_Zeile,Spalte + 1).getValue();
      var Editoren_Mails = [];

      Logger.log("Bereich: " + Array_Sheets[z] + "\t\tName: " + Sheet.getRange(2,Spalte).getValue() + "\t\tID: " + ID);

      try
      {
        var Sheet_Export = SpreadsheetApp.openById(ID);
      }
      catch(err)
      {
        Logger.log("Keine Tabelle");
        try
        {
          var Sheet_Export = DocumentApp.openById(ID);
        }
        catch(err)
        {
          Logger.log("Kein Document");
          try
          {
            var Sheet_Export = FormApp.openById(ID);
            Logger.log("Done");
          }
          catch(err)
          {
            Logger.log("Kein Formular");
            try
            {
              var Sheet_Export = SlidesApp.openById(ID);
            }
            catch(err)
            {
              Logger.log("Keine Präsentation");
              try
              {
                var Sheet_Export = DriveApp.getFolderById(ID);
                Editoren_Mails.push(Sheet_Export.getOwner().getEmail());
                Logger.log("Owner: " + Sheet_Export.getOwner().getEmail());
              }
              catch(err)
              {
                MailApp.sendEmail("1@1.de","LSPD Master Fehler","Fehler: " + err);
                Logger.log("ERROR auch Kein Ordner\t" + ID + "\n" + err);
              }
            }
          }
        }
      }

      var Editoren = Sheet_Export.getEditors();

      for(var y = 0; y < Editoren.length; y++)
      {
        Editoren_Mails.push(Editoren[y].getEmail());
      }

      var Editoren_Remove = [];

      for(var i = 0; i < Editoren_Mails.length; i++)
      {
        if(Editoren_Mails[i] != "1@1.de" && Editoren_Mails[i] != "gvmp.lspd.bot@gmail.com")
        {
          Editoren_Remove.push(Editoren_Mails[i]);
        }
      }

      Logger.log("Remove: " + Editoren_Remove);

      try
      {
        for(var i = 0; i < Editoren_Remove.length; i++)
        {
          try
          {
            Sheet_Export.removeEditor(Editoren_Remove[i]);
          }
          catch(err)
          {
            Logger.log("Fehler bei entfernen von" + Editoren_Remove[i] + "\n" + err.stack)
          }
        }

        //DriveApp.getFileById(ID).setTrashed(true);
        Logger.log("Datei gelöscht");
      }
      catch(err)
      {
        Logger.log("Keine Berechtigung auf ID: " + ID + "\nError: " + err);
      }

      Spalte = Spalte + 3;
    }
  }

  Selbstzerstörung_Lesen();

  var SS = SpreadsheetApp.getActive();
  var Editoren = SS.getEditors();

  for(var i = 0; i < Editoren.length; i++)
  {
    if(Editoren[i].getEmail() != "1@1.de" && Editoren[i].getEmail() != "gvmp.lspd.bot@gmail.com")
    {
      SS.removeEditor(Editoren[i].getEmail());
    }
  }
}

function Selbstzerstörung_Lesen()
{
  var Array_Sheets = ["Lesen"];

  var ID_Zeile = 3;

  for(var z = 0; z < Array_Sheets.length; z++)
  {
    var Sheet = SpreadsheetApp.getActive().getSheetByName(Array_Sheets[z]);

    var Anzahl_Sheets = Sheet.getRange("A2").getValue();
    var Spalte = 2;

    for(var z2 = 1; z2 <= Anzahl_Sheets; z2++)
    {
      var ID = Sheet.getRange(ID_Zeile,Spalte + 1).getValue();
      var Editoren_Mails = [];

      Logger.log("Bereich: " + Array_Sheets[z] + "\t\tName: " + Sheet.getRange(2,Spalte).getValue() + "\t\tID: " + ID);

      try
      {
        var Sheet_Export = SpreadsheetApp.openById(ID);
      }
      catch(err)
      {
        Logger.log("Keine Tabelle");
        try
        {
          var Sheet_Export = DocumentApp.openById(ID);
        }
        catch(err)
        {
          Logger.log("Kein Document");
          try
          {
            var Sheet_Export = FormApp.openById(ID);
          }
          catch(err)
          {
            Logger.log("Kein Formular");
            try
            {
              var Sheet_Export = SlidesApp.openById(ID);
            }
            catch(err)
            {
              Logger.log("Keine Präsentation");
              try
              {
                var Sheet_Export = DriveApp.getFolderById(ID);
                Editoren_Mails.push(Sheet_Export.getOwner().getEmail());
                Logger.log("Owner: " + Sheet_Export.getOwner().getEmail());
              }
              catch(err)
              {
                MailApp.sendEmail("1@1.de","LSPD Master Fehler","Fehler: " + err);
                Logger.log("ERROR auch Kein Ordner\t" + ID + "\n" + err);
              }
            }
          }
        }
      }

      var Editoren = Sheet_Export.getViewers();

      for(var y = 0; y < Editoren.length; y++)
      {
        Editoren_Mails.push(Editoren[y].getEmail());
      }

      var Editoren_Remove = [];

      for(var i = 0; i < Editoren_Mails.length; i++)
      {
        if(Editoren_Mails[i] != "1@1.de" && Editoren_Mails[i] != "gvmp.lspd.bot@gmail.com")
        {
          Editoren_Remove.push(Editoren_Mails[i]);
        }
      }

      Logger.log("Remove: " + Editoren_Remove);

      try
      {
        for(var i = 0; i < Editoren_Remove.length; i++)
        {
          try
          {
            Sheet_Export.removeViewer(Editoren_Remove[i]);
          }
          catch(err)
          {
            Logger.log("Fehler bei entfernen von" + Editoren_Remove[i] + "\n" + err.stack)
          }
        }

        //DriveApp.getFileById(ID).setTrashed(true);
        Logger.log("Datei gelöscht");
      }
      catch(err)
      {
        Logger.log("Keine Berechtigung auf ID: " + ID + "\nError: " + err);
      }

      Spalte = Spalte + 3;
    }
  }
}