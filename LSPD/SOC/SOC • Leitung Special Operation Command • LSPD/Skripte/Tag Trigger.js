function Tagges_Trigger()
{
  var Sheet_Auswertung_Meldung = SpreadsheetApp.getActive().getSheetByName("Auswertung Meldungsblatt");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Array_Meldung = Sheet_Auswertung_Meldung.getRange("B4:D50").getValues();
  var Array_Meldung_Namen = Sheet_Auswertung.getRange("B3:B" + Sheet_Auswertung.getLastRow()).getValues();
  var Array_Meldung_Tickets = Sheet_Auswertung.getRange("C3:C" + Sheet_Auswertung.getLastRow()).getValues();

  for(var y = 0; y < Array_Meldung_Tickets.length; y++)
  {
    if(Array_Meldung_Tickets[y][0] >= 3)
    {
      var Gefunden = false;

      for(var y2 = 0; y2 < Array_Meldung.length; y2++)
      {
        if(Array_Meldung[y2][0] == Array_Meldung_Namen[y][0])
        {
          var Monat = new Date().getMonth()
          Gefunden = true;

          if(Array_Meldung[y2][1].getMonth() != Monat)
          {
            Logger.log("Add " + Array_Meldung_Namen[y][0]);

            Sheet_Auswertung_Meldung.insertRowAfter(4);

            Sheet_Auswertung_Meldung.getRange("B5:D5").setValues([[Array_Meldung_Namen[y][0],new Date(),"Offen"]]);

            break;
          }
          else
          {
            break;
          }
        }
      }

      if(Gefunden == false)
      {
        Logger.log("Add " + Array_Meldung_Namen[y][0]);

        Sheet_Auswertung_Meldung.insertRowAfter(4);

        Sheet_Auswertung_Meldung.getRange("B5:D5").setValues([[Array_Meldung_Namen[y][0],new Date(),"Offen"]]);
      }
    }
  }
}
