function Minuten_Trigger()
{
  Dienstblatt_Sicherung();

  var Sheet_Copy = SpreadsheetApp.getActive().getSheetByName("Copy Master");

  var Array_Key = LSPD_Copy.Properties_Get_All_Keys();
  var Array_Ausgabe = new Array();

  for(var i = 0; i < Array_Key.length; i++)
  {
    if(Array_Key[i][0] != "Popup Copy Aktiv" && Array_Key[i][0] != "Popup Copy User")
    {
      var Array_Temp = Array_Key[i][1].toString().split("|#|");

      Array_Ausgabe.push([Array_Key[i][0],Array_Temp[0].toString().replace("Kopie von ","").replace("Copy from ","") ,Array_Temp[1],Array_Temp[2],Array_Temp[3]]);
    }
  }

  Array_Ausgabe = Array_Ausgabe.sort(function(a, b)
  {
    if (new Date(a[0]) < new Date(b[0])) return -1;
    if (new Date(a[0]) > new Date(b[0])) return 1;
    return 0;
  });

  Logger.log(Array_Ausgabe);
  
  if(Array_Ausgabe.length > 0)
  {
    Sheet_Copy.getRange(5,2,Array_Ausgabe.length,Array_Ausgabe[0].length).setValues(Array_Ausgabe);
    Sheet_Copy.getRange(5,8,Array_Ausgabe.length,3).insertCheckboxes();
  }

  SpreadsheetApp.flush();
  
  Benachrichtigung();
}


// -- Benachrichtigung import aus Tabellenblatt auswertungsgedöns -- //

function Benachrichtigung()
{
  var Sheet_Copy = SpreadsheetApp.getActive().getSheetByName("Copy Master");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Array_Copy = Sheet_Copy.getRange("B5:J" + Sheet_Copy.getRange("B3").getValue()).getValues();
  var Array_Mails = Sheet_Auswertung.getRange("E3:E" + (Sheet_Auswertung.getRange("E1").getValue()-1)).getValues();
  var Array_Popup = Sheet_Auswertung.getRange("G3:G" + (Sheet_Auswertung.getRange("G1").getValue()-1)).getValues();

  var Popup_Gefunden = false;

  for(var i = 0; i < Array_Copy.length; i++)
  {
    if(Array_Copy[i][0] != "")
    {
      if(Array_Copy[i][6] == false) // Popup Check
      {
        Popup_Gefunden = true;
      }

      if(Array_Copy[i][8] == false) // Email Check
      {
        for(var y = 0; y < Array_Mails.length; y++)
        {
          MailApp.sendEmail(Array_Mails[y][0],"LSPD Tabellen Blatt Kopiert " + Array_Copy[i][1],"Datum: " + Utilities.formatDate(Array_Copy[i][0],SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy HH:mm:ss") + "\nBlatt Name: " + Array_Copy[i][1] + "\nName: " + Array_Copy[i][2] + "\nEmail: " + Array_Copy[i][3] + "\nLink: " + Array_Copy[i][4] + "\n\nCopy Master Link: https://docs.google.com/spreadsheets/d/19-XqdxjVdy0nOGIS7uNJ2aL53GuOS9RYiFp0NiAln0s/edit#gid=944417460");
        }

        Sheet_Copy.getRange("J" + (i+5)).setValue(true);
      }
    }
  }

  LSPD_Copy.Propertie_Setzen("Popup Copy Aktiv",Popup_Gefunden,"Script")
  LSPD_Copy.Propertie_Setzen("Popup Copy User",Array_Popup.join("|#|"),"Script")
}

function Sort(a,b)
{
  Logger.log(new Date(b[0]).getTime() - new Date(a[0]).getTime())
  return new Date(b[0]).getTime() - new Date(a[0]).getTime()
}