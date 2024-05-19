function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  var user = Session.getTemporaryActiveUserKey();
  
  Logger.log("Benutzer: " + user);

  ui.createMenu('Funktionen')
    .addItem('Personalliste Sortieren', 'Personal_Sortieren')
    .addToUi();

  LSPD.onOpen();

  var Sheet_Auswertung_Ticket = SpreadsheetApp.getActive().getSheetByName("Auswertung Meldungsblatt");
  var Array_Tickets = Sheet_Auswertung_Ticket.getRange("B5:D50").getValues();

  for(var y = 0; y < Array_Tickets.length; y++)
  {
    if(Array_Tickets[y][2] == "Offen")
    {
      var Button = ui.alert("Meldungen",Array_Tickets[y][0] + " hat 3+ Meldungen diesen Monat.\nFall ans Detective Bureau weitergeleitet?",ui.ButtonSet.YES_NO_CANCEL);

      if(Button == ui.Button.YES)
      {
        Sheet_Auswertung_Ticket.getRange("D" + (y+5)).setValue("Angenommen " + LSPD.Umwandeln());
      }
      else if(Button == ui.Button.NO)
      {
        Sheet_Auswertung_Ticket.getRange("D" + (y+5)).setValue("Abgelehnt " + LSPD.Umwandeln());
      }
    }
  }
}