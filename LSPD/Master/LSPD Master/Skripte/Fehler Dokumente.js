function Freigaben_Fehler()
{
  var Import_Abteilung = SpreadsheetApp.getActive().getSheetByName("Import Abteilungen").getRange("B3").getValue();
  Logger.log("Import Abteilung: " + Import_Abteilung);
  
  if(Import_Abteilung != "#N/A" && Import_Abteilung != "#REF!"  && Import_Abteilung != "#ERROR!")
  {
    var Array_Sheets = ["Ordner","Formulare"];

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
        var Editoren_Mails = [];

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
                  Editoren_Mails.push(Sheet_Export.getOwner().getEmail());
                  //Logger.log("Owner: " + Sheet_Export.getOwner().getEmail());
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

        var Editoren = Sheet_Export.getEditors();


        for(var y = 0; y < Editoren.length; y++)
        {
          Editoren_Mails.push(Editoren[y].getEmail());
        }
        
        var Editoren_Remove = [];
        var Editoren_Invite = [];
        var Gefunden = false;

        for(var y = 0; y < Editoren_Mails.length; y++)
        {
          Gefunden = false;

          for(var x = 0; x < Array_Freigaben.length; x++)
          {
            if(Editoren_Mails[y].toString().toUpperCase() == Array_Freigaben[x][1].toString().toUpperCase())
            {
              Gefunden = true;
              break;
            }
          }

          if(Gefunden == false)
          {
            Editoren_Remove.push(Editoren_Mails[y]);
          }
        }

        for(var y = 0; y < Array_Freigaben.length; y++)
        {
          Gefunden = false;

          for(var x = 0; x < Editoren_Mails.length; x++)
          {
            if(Array_Freigaben[y][1].toString().toUpperCase() == Editoren_Mails[x].toString().toUpperCase())
            {
              Gefunden = true;
              break;
            }
          }

          if(Gefunden == false)
          {
            Editoren_Invite.push(Array_Freigaben[y][1]);
          }
        }

        if(Editoren_Remove.length != 0 || Editoren_Invite.length != 0)
        {
          Logger.log("Remove: " + Editoren_Remove);
          Logger.log("Invite: " + Editoren_Invite);
        }

        //if(Editoren_Invite.indexOf("#N/A") != -1)
        //{
        //  Logger.log("#N/A wurde gefunden. Tabelle wird übersprungen")
        //}
        //else
        //{
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

            for(var i = 0; i < Editoren_Invite.length; i++)
            {
              try
              {
                if(Editoren_Invite[i] != "#N/A")
                {
                  Sheet_Export.addEditor(Editoren_Invite[i]);
                }
              }
              catch(err)
              {
                Logger.log("Fehler bei hinzufügen von " + Editoren_Remove[i] + "\n" + err.stack)
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
