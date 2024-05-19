function Minuten_Trigger() 
{
  var Datum = new Date();

  var Stunde = Datum.getHours();
  var Minute = Datum.getMinutes();

  Logger.log(Stunde + " " + Minute);
  
  if(Stunde == 23 && Minute == 57)
  {

    Kalender_Tages_Wechsel();
  }
}