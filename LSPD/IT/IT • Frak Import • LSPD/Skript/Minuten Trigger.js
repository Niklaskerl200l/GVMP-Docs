function Minuten_Trigger()
{
  var Datum = new Date();

  var Stunde = Datum.getHours();
  var Minute = Datum.getMinutes();

  Logger.log(Stunde + " " + Minute);

  if(Stunde == 23 && Minute == 55)
  {
    Namen();
  }

  if(Stunde == 0 && Minute == 0)
  {
    Fraktionen();
  }

  if(Minute % 5 == 0)
  {
    Update_GOV_Namen()
    Dienstblatt();
  }
}
