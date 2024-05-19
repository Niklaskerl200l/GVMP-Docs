function Lese_Rechte()
{
  var Import_Abteilung = SpreadsheetApp.getActive().getSheetByName("Import Abteilungen").getRange("B3").getValue();
  Logger.log("Import Abteilung: " + Import_Abteilung);
  
  if(Import_Abteilung != "#N/A" && Import_Abteilung != "#REF!"  && Import_Abteilung != "#ERROR!")
  {
    var Array_Sheets = ["Lesen"];

    var ID_Zeile = 3;

    for(var z = 0; z < Array_Sheets.length; z++)
    {
      var Sheet = SpreadsheetApp.getActive().getSheetByName(Array_Sheets[z]);

      var Anzahl_Sheets = Sheet.getRange("A2").getValue();
      var Zeile = 18, Spalte = 2, Spalten = 2;

      for(var z2 = 1; z2 <= Anzahl_Sheets; z2++)
      {
        var Array_Freigaben = Sheet.getRange(Zeile,Spalte,Sheet.getLastRow() - Zeile,Spalten).getValues();
        Array_Freigaben = Array_Freigaben.filter(function(e){return e[0] != "" && e[1] != "" && e[1] != "#NV"});

        var ID = Sheet.getRange(ID_Zeile,Spalte + 1).getValue();
        var Viewer_Mails = [];

        Logger.log("Bereich: " + Array_Sheets[z] + "\t\tName: " + Sheet.getRange(2,Spalte).getValue() + "\t\tID: " + ID);

        try
        {
          var Sheet_Export = SpreadsheetApp.openById(ID);
        }
        catch(err)
        {
          try
          {
            var Sheet_Export = DocumentApp.openById(ID);
            Logger.log("Ist Document");
          }
          catch(err)
          {
            try
            {
              var Sheet_Export = FormApp.openById(ID);
              Logger.log("Ist Formular");
            }
            catch(err)
            {
              try
              {
                var Sheet_Export = SlidesApp.openById(ID);
                Logger.log("Ist Präsentation");
              }
              catch(err)
              {
                try
                {
                  var Sheet_Export = DriveApp.getFolderById(ID);
                  Logger.log("Ist Ordner");
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

        if(DriveApp.getFileById(ID).isShareableByEditors())
        {
          Logger.log("Remove ShareableByEditors");
          DriveApp.getFileById(ID).setShareableByEditors(false);
        }

        var Viewers = Sheet_Export.getViewers();
        var Editors = Sheet_Export.getEditors();


        for(var y = 0; y < Viewers.length; y++)
        {
          var Gefunden = false;

          for(var x = 0; x < Editors.length; x++)
          {
            if(Viewers[y].getEmail() == Editors[x].getEmail())
            {
              Gefunden = true;
              break;
            }
          }

          if(Gefunden == false)
          {
            Viewer_Mails.push(Viewers[y].getEmail());
          }
        }
        
        var Viewers_Remove = [];
        var Viewers_Invite = [];
        var Gefunden = false;

        for(var y = 0; y < Viewer_Mails.length; y++)
        {
          Gefunden = false;

          for(var x = 0; x < Array_Freigaben.length; x++)
          {
            if(Viewer_Mails[y].toString().toUpperCase() == Array_Freigaben[x][1].toString().toUpperCase())
            {
              Gefunden = true;
              break;
            }
          }

          if(Gefunden == false)
          {
            Viewers_Remove.push(Viewer_Mails[y]);
          }
        }

        for(var y = 0; y < Array_Freigaben.length; y++)
        {
          Gefunden = false;

          for(var x = 0; x < Viewer_Mails.length; x++)
          {
            if(Array_Freigaben[y][1].toString().toUpperCase() == Viewer_Mails[x].toString().toUpperCase())
            {
              Gefunden = true;
              break;
            }
          }

          if(Gefunden == false)
          {
            Viewers_Invite.push(Array_Freigaben[y][1]);
          }
        }

        if(Viewers_Remove.length != 0 || Viewers_Invite.length != 0)
        {
          Logger.log("Remove: " + Viewers_Remove);
          Logger.log("Invite: " + Viewers_Invite);
        }

        //if(Viewers_Invite.indexOf("#N/A") != -1)
        //{
        //  Logger.log("#N/A wurde gefunden. Tabelle wird übersprungen")
        //}
        //else
        //{
          try
          {
            for(var i = 0; i < Viewers_Remove.length; i++)
            {
              try
              {
                Sheet_Export.removeViewer(Viewers_Remove[i]);
              }
              catch(err)
              {
                Logger.log("Fehler bei entfernen von" + Viewers_Remove[i] + "\n" + err.stack)
              }
            }

            for(var i = 0; i < Viewers_Invite.length; i++)
            {
              try
              {
                if(Viewers_Invite[i] != "#N/A")
                {
                  Sheet_Export.addViewer(Viewers_Invite[i]);
                }
              }
              catch(err)
              {
                Logger.log("Fehler bei hinzufügen von " + Viewers_Remove[i] + "\n" + err.stack)
              }
            }
          }
          catch(err)
          {
            Logger.log("Keine Berechtigung auf ID: " + ID + "\nError: " + err);
          }
        //}

        Spalte = Spalte + 3;
      }
    }
  }
}
