function Geschwindigkeitsüberschreitungen(e)
{
  var error = undefined;
  var Fehler = false;

  try
  {
    Geschwindigkeitsüberschreitungen_Übertragung(e);
  }
  catch(err)
  {
    error = err;
    Logger.log(err.stack);
    Fehler = true;
  }
}

function Geschwindigkeitsüberschreitungen_Übertragung(e)
{
  var Sheet = SpreadsheetApp.openById(LSPD.ID_Geschwindigkeits_Tickets).getSheetByName("Tickets");
  var Werte = e.namedValues;

  var Daten = 
  [[
    Werte["Ihr Name"],
    Werte["1 - 20 KM/H Überschreitungen"],
    Werte["21 - 50 KM/H Überschreitungen"],
    Werte["51 - 100 KM/H Überschreitungen"],
    "",
    Werte["Art"],
    Werte["Notiz, Bemerkung oder Begründung"],
    "LSPD IT",
    new Date()
  ]];

  Sheet.insertRowBefore(13);
  Sheet.getRange("B13:J13").setValues(Daten);
}