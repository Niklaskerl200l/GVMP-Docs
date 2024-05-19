function Beobachtungsliste()
{
  Sheet_Beobachtungsliste = SpreadsheetApp.getActive().getSheetByName("Beobachtungsliste").getRange("I3").setValue(true);
}
