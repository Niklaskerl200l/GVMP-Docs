function Minutentrigger()
{
  var Zeitstempel = new Date();
  var Zeit = Utilities.formatDate(Zeitstempel, "CET", "HH:mm");

  Logger.log(Zeit);

  if(Zeit == "00:05")
  {
    Tankkontrolle_Zufall();
  }

  if(Zeitstempel.getMinutes() == 00)
  {
    Detective_Meldung();
  }
}