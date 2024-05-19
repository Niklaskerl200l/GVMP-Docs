function Dienstblatt()
{
  var Sheet_Dienstblatt = SpreadsheetApp.getActive().getSheetByName("Dienstblatt Übersicht");

  var Array_Suche = Sheet_Dienstblatt.getRange("J6:L20").getValues();

  var Array_Anwesend = Sheet_Dienstblatt.getRange("C6:C150").getValues();

  for(var i = 0; i < Array_Suche.length; i++)
  {
    for(var j = 0; j < Array_Anwesend.length; j++)
    {
      if(Array_Suche[i][0] == Array_Anwesend[j][0] && Array_Anwesend[j][0] != "")
      {
        var Datum = new Date();

        Datum.setHours(Datum.getHours() - 1);

        if(Array_Suche[i][1] <= Datum || Array_Suche[i][1] == "")
        {
          if(Array_Suche[i][2] == "") Array_Suche[i][2] = "1@1.de" ;

          Logger.log("Send mail ");
          MailApp.sendEmail(Array_Suche[i][2],"Frak DB Suche " + Array_Suche[i][0],"Person: " + Array_Suche[i][0] + " ist Wach!")
          Sheet_Dienstblatt.getRange("K" + (i+6)).setValue(new Date());
        }

        break;
      }
    }
  }
}
