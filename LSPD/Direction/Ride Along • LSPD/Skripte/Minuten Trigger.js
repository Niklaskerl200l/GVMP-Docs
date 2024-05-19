function Minuten_Trigger()
{
  var Datum = new Date();

  var Stunde = Datum.getHours();
  var Minute = Datum.getMinutes();

  Logger.log(Stunde + " " + Minute);

  var Sheet_Ride_Along = SpreadsheetApp.getActive().getSheetByName("Ride Along");

  var Array_RA = Sheet_Ride_Along.getRange("K5:K8").getValues();

  for(var i = 0; i < Array_RA.length; i++)
  {
    if(Array_RA[i][0] != "" && new Date(Array_RA[i][0]) < new Date())
    {
      Logger.log("Delete RA");
      Sheet_Ride_Along.getRange("I" + (i+5) + ":M" + (i+5)).setValue("");
    }
  }
}
