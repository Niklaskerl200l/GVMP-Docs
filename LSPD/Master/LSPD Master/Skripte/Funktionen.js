function Spalte_in_Index(Text)    // String in Spaltenindex
{
  Text = Text.toUpperCase();
  
  if(Text.length == 1)
  {
    return Switch_ABC(Text);
  }
  else if(Text.length == 2)
  {
    return (Switch_ABC(Text[0]) * 26) + Switch_ABC(Text[1]);
  }
}

function Switch_ABC(Text)
{
  switch(Text)
    {
      case "A" : return 1; break;
      case "B" : return 2; break;
      case "C" : return 3; break;
      case "D" : return 4; break;
      case "E" : return 5; break;
      case "F" : return 6; break;
      case "G" : return 7; break;
      case "H" : return 8; break;
      case "I" : return 9; break;
      case "J" : return 10; break;
      case "K" : return 11; break;
      case "L" : return 12; break;
      case "M" : return 13; break;
      case "N" : return 14; break;
      case "O" : return 15; break;
      case "P" : return 16; break;
      case "Q" : return 17; break;
      case "R" : return 18; break;
      case "S" : return 19; break;
      case "T" : return 20; break;
      case "U" : return 21; break;
      case "V" : return 22; break;
      case "W" : return 23; break;
      case "X" : return 24; break;
      case "Y" : return 25; break;
      case "Z" : return 26; break;
      default: return 0;
    }
}


function Dienstblatt_Sicherung()
{
  var Zeitstempel = new Date();
  if(Utilities.formatDate(Zeitstempel, "CET", "HH:mm") == "00:05" || Utilities.formatDate(Zeitstempel, "CET", "HH:mm") == "08:05" || Utilities.formatDate(Zeitstempel, "CET", "HH:mm") == "16:05")
  {
    try
    {
      var SS_Dienstblatt = SpreadsheetApp.openById(SpreadsheetApp.getActive().getSheetByName("LSPD").getRange("C3").getValue());
      var Array_Dienstblatt = SS_Dienstblatt.getEditors();

      var Array_Dienstblatt_EMails = [];
      for(var i = 0; i < Array_Dienstblatt.length; i++)
      {
        Array_Dienstblatt_EMails.push(Array_Dienstblatt[i].getEmail());
      }

      var Whitelist = ["1@1.de"];

      for(var o = 0; o < Array_Dienstblatt_EMails.length; o++)
      {
        if(Whitelist.toString().toUpperCase().includes(Array_Dienstblatt_EMails[o].toString().toUpperCase()) == false)
        {
          SS_Dienstblatt.removeEditor(Array_Dienstblatt_EMails[o]);
          Logger.log("\tEntfernt: " + Array_Dienstblatt_EMails[o]);
        }
      }

      Logger.log("Remove abgeschlossen.");
      Utilities.sleep(5000);

      for(var p = 0; p < Array_Dienstblatt_EMails.length; p++)
      {
        SS_Dienstblatt.addEditor(Array_Dienstblatt_EMails[p]);
        Logger.log("\t\tFüge hinzu: " + Array_Dienstblatt_EMails[p]);
      }

      Logger.log("Add abgeschlossen.");

      Logger.log("Dienstblatt-Sicherung abgeschlossen!");
    }
    catch(err)
    {
      Logger.log(err.stack);
    }
  }
}