function Minuten_Trigger()
{
  var Datum = new Date();

  var Stunde = Datum.getHours();
  var Minute = Datum.getMinutes();
  var Tag = Datum.getDay()

  Logger.log(Stunde + " " + Minute + " " + Tag);

  var Sheet_Export_Personal = SpreadsheetApp.getActive().getSheetByName("Export Personal");

  Sheet_Export_Personal.getRange("B4").setFormula("={'Personal Master'!$D$6:$D\\ARRAYFORMULA(IF(ISBLANK('Personal Master'!$D$6:$D);;VLOOKUP('Personal Master'!$D$6:$D;'Auswertungsgedöns'!$E$5:$G$17;3;FALSE)))\\'Personal Master'!$C$6:$C\\'Personal Master'!$B$6:$B\\'Personal Master'!$E$6:$E\\'Personal Master'!$S$6:$S\\'Personal Master'!$I$6:$I\\'Personal Master'!$J$6:$J}");

  Zeitbearbeitung_Zeitsystem();

  if(Stunde == 0 && Minute == 0 && Tag == 1)
  {
    Delete_Nicht_Rankup();
  }
}
