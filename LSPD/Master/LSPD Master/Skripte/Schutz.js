var Fehler = false;

function Schutz()
{
  var Sheet_Schutz = SpreadsheetApp.getActive().getSheetByName("Schutz");

  var Letzte_Zeile = Sheet_Schutz.getLastRow();
  var Letzte_Spalte = Sheet_Schutz.getLastColumn();

  var Array_Schutz = Sheet_Schutz.getRange(1,1,Letzte_Zeile,Letzte_Spalte).getValues();
  var Array_Zugriff = new Array();

  for(var y = 1; y < Array_Schutz.length; y=y+7)    // Erstelle Array für die Zugriffe
  {
    var Array_Freigaben = new Array();

    for(var y2 = y; y2 < y+6; y2++)
    {
      var Array_Email = new Array();

      if(Array_Schutz[y2][4] == "Sonderzugriff")
      {
        for(var x = 5; Array_Schutz[y2][x] != "" && Array_Schutz[y2][x] != undefined; x++)
        {
          var Array_Split = Array_Schutz[y2][x].toString().split(":",2);
          var Gefunden = false;


          for(var i = 0; i < Array_Freigaben.length; i++)
          {
            if(Array_Freigaben[i][0] == Array_Split[0])
            {
              Gefunden = true;
              Array_Freigaben[i][1].push(Array_Split[1])
              break;
            }
          }

          if(Gefunden == false)
          {
            Array_Freigaben.push([Array_Split[0],[Array_Split[1]]]);
          }
        }
      }
      else
      {
        for(var x = 5; Array_Schutz[y2][x] != "" && Array_Schutz[y2][x] != undefined; x++)
        {
          Array_Email.push(Array_Schutz[y2][x]);
        }

        Array_Freigaben.push([Array_Schutz[y2][4],Array_Email]);
      }
      
    }

    Array_Zugriff.push([Array_Schutz[y][1],Array_Schutz[y + 2][2],Array_Schutz[y + 4][2],Array_Freigaben]);
  }




  for(var y = 0; y < Array_Zugriff.length; y++)   // Prüfe Zugriffe und gib Frei/ Entferne
  {
    Logger.log("Schutz für " + Array_Zugriff[y][0]) + "\tID: " + Array_Zugriff[y][1] + "\tFreigabe: " + Array_Zugriff[y][2];

    try
    {
      var Sheet_Export = SpreadsheetApp.openById(Array_Zugriff[y][1]);
    }
    catch(err)
    {
      Logger.log(err.stack);
      break;
    }


    var Protection = Sheet_Export.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    Schutz_Freigabe(Protection);

    var Protection = Sheet_Export.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    Schutz_Freigabe(Protection);

    function Schutz_Freigabe(Protection)
{
  for(var i = 0; i < Protection.length; i++)
  {
    var Sheet_Name = Protection[i].getRange().getSheet().getSheetName();
    var Titel = Protection[i].getDescription();

    try
    {
      var Array_Bearbeiter = Protection[i].getEditors();
    }
    catch(err)
    {
      Logger.log(err.stack);
      Fehler = true;
      break;
    }

    var Array_Editor = new Array();

    for(var i2 = 0; i2 < Array_Bearbeiter.length; i2++)
    {
      Array_Editor.push(Array_Bearbeiter[i2].getEmail());
    }

    var Array_Freigaben = Array_Zugriff[y][3];
    var Array_Access = new Array();
    var Gefunden = false;

    for(var i3 = 0; i3 < Array_Freigaben.length; i3++)
    {
      if(Array_Freigaben[i3][0] != "" && Titel.includes(Array_Freigaben[i3][0]))
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)   // Wenn Titel Unbekannt
    {
      Logger.log("  " + Titel + " unbekannt");
    }
    else
    {
      for(var y2 = 0; y2 < Array_Zugriff[y][3].length; y2++)
      {
        if(Titel.includes(Array_Freigaben[y2][0]))
        {
          for(var x = 0; x < Array_Freigaben[y2][1].length; x++)
          {
            Array_Access.push(Array_Freigaben[y2][1][x]);
          }
        }
      }

      var Array_Invite = new Array();
      var Array_Remove = new Array();

      for(var x1 = 0; x1 < Array_Access.length; x1++)
      {
        var Gefunden = false;

        for(var x2 = 0; x2 < Array_Editor.length; x2++)
        {
          if(Array_Access[x1] == Array_Editor[x2])
          {
            Gefunden = true;
            break;
          }
        }

        if(Gefunden == false)
        {
          Array_Invite.push(Array_Access[x1]);
        }
      }

      for(var x1 = 0; x1 < Array_Editor.length; x1++)
      {
        var Gefunden = false;

        for(var x2 = 0; x2 < Array_Access.length; x2++)
        {
          if(Array_Editor[x1] == Array_Access[x2])
          {
            Gefunden = true;
            break;
          }
        }

        if(Gefunden == false)
        {
          Array_Remove.push(Array_Editor[x1]);
        }
      }
      if(Array_Invite.length != 0 || Array_Remove.length != 0)
      {
        Logger.log("   Sheet: " + Sheet_Name + "\tSchutz Titel: " + Titel + "\n      Invite: " + Array_Invite + "\n      Remove: " + Array_Remove);
      }

      try
      {
        Protection[i].addEditors(Array_Invite);
        Protection[i].removeEditors(Array_Remove);
      }
      catch(err)
      {
        Logger.log(err.stack);
        Fehler = true;
      }
    }
  }
}
  }

  if(Fehler == true)
  {
    Fehler++;
  }
}

