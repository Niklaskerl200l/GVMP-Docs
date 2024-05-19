function Minuten_Trigger()
{
  try
  {
    var Jetzt = new Date();

    var Wochentag = Jetzt.getDay();
    var Tag = Jetzt.getDate();
    var Stunde = Jetzt.getHours();
    var Minute = Jetzt.getMinutes();

    Logger.log(Stunde + " " + Minute)

    try
    {
      if(Stunde == 0 && Minute == 0)
      {
        Tag_Trigger();

        if(Stunde == 0 && Minute == 0 && Wochentag == 1)
        {
          Woche_Trigger();
        }
        if(Stunde == 0 && Minute == 0 && Tag == 1)
        {
          Monat_Trigger();
        }
      }
    }
    catch(err)
    {
      MailApp.sendEmail("1@1.de","GVMP Zeitsystem Fehler!!!!","Fehler bei einen der Tag / Wochen / Monat Trigger dringend!!!!");
      Logger.log(err.stack);
    }
    
    Stempeluhr_Zeitsystem();

    Leitstelle_Zeitsystem();

    Entlassung();
  }
  catch(err)
  {
    Logger.log(err.stack);
    
    Fehler;
  }
}
