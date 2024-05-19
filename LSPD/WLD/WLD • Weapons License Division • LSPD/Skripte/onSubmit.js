function onSubmit(e)
{
  var Werte = e.namedValues;

  var Termin = Werte.Termin;
  var Name = Werte.Name;
  var Tel = Werte["IC-Telefonnummer"];

  try
  {
    var Sheet_Termin = SpreadsheetApp.getActive().getSheetByName("WLD " + Termin);

    var Letzte_Zeile = Sheet_Termin.getRange("E3").getValue() + 1;

    var Array_Namen = Sheet_Termin.getRange("D5:D" + Letzte_Zeile).getValues();

    var Gefunden = false;

    for(var i = 0; i < Array_Namen.length; i++)
    {
      if(Name == Array_Namen[i][0])
      {
        Logger.log(Name + " ist bereits im Kurs: " + Termin)
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      Sheet_Termin.getRange("D" + Letzte_Zeile + ":E" + Letzte_Zeile).setValues([[Name,Tel]]);
    }
  }
  catch(err)
  {
    Logger.log(err.stack);
    
    var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");

    Sheet_Startseite.getRange("M32").setValue(Sheet_Startseite.getRange("M32").getValue() + "Anmeldung von " + Name + "(" + Tel + ") für den " + Termin + " Fehlerhaft\n");
  }
}
