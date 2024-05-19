var ID_GTF = LSPD.ID_GTF;
var ID_DCI = "1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w";

function Zugehoerigkeit(e)
{
  var Werte = e.namedValues;

  var Sheet_GTF = SpreadsheetApp.openById(ID_GTF).getSheetByName("Zugehörigkeit");

  var Letzte_Zeile = Sheet_GTF.getLastRow() + 1;

  var Array_Ausgabe = [[new Date(), Werte["Name der Person"], Werte["Telefonnummer der Person"], Werte["Aktuelle Zugehörigkeit"], Werte["Gewünschte Zugehörigkeit"], Werte["Name. Dienstnummer des Officers"], '=IFERROR(VLOOKUP($C' + Letzte_Zeile + ';\'Import Aktuell\'!$B$3:$I;2;FALSE);)', '=IFERROR(VLOOKUP($C' + Letzte_Zeile + ';\'Import Aktuell\'!$B$3:$I;8;FALSE);)']];

  Sheet_GTF.insertRowAfter(Letzte_Zeile-1);
  Sheet_GTF.getRange("B" + Letzte_Zeile + ":I" + Letzte_Zeile).setValues(Array_Ausgabe);
  Sheet_GTF.getRange("N" + Letzte_Zeile).insertCheckboxes();

  SpreadsheetApp.getActive().getSheetByName("Zugehörigkeit").getRange(e.range.getRow(),e.range.getLastColumn() + 1).setValue(true).insertCheckboxes();

  SpreadsheetApp.flush();

  //------------- DCI ---------------//

  var Sheet_DCI = SpreadsheetApp.openById(ID_DCI).getSheetByName("Zugehörigkeit");

  Sheet_DCI.insertRowAfter(Letzte_Zeile-1);
  Sheet_DCI.getRange("B" + Letzte_Zeile + ":I" + Letzte_Zeile).setValues(Array_Ausgabe);
  Sheet_DCI.getRange("N" + Letzte_Zeile).insertCheckboxes();

  SpreadsheetApp.getActive().getSheetByName("Zugehörigkeit").getRange(e.range.getRow(),e.range.getLastColumn() + 2).setValue(true).insertCheckboxes();
}
