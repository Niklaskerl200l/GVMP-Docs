function onSubmit(e)
{
  var Sheetname = e.range.getSheet().getSheetName();

  var Lock = LockService.getScriptLock();
    
  try
  {
    Lock.waitLock(118000);
  }
  catch(e)
  {
    Logger.log('Timeout wegen Lock bei Beschlagnahmungen');
    Fehler++;
  }

  switch(Sheetname)
  {
    case "LSPD Einstellung" : Einstellung(e); break;
    case "LSPD Entlassungen" : Entlassung(e); break;
    case "Telefonnummeränderung" : Telefonnummeraenderung(e); break;
    case "Namensänderung" : Namensaenderung(e); break;
    case "Zugehörigkeit" :  Zugehoerigkeit(e); break;
    case "Geschwindigkeitsticket": Geschwindigkeitsüberschreitungen(e); break;
  }

  SpreadsheetApp.flush();
  Lock.releaseLock();
}

