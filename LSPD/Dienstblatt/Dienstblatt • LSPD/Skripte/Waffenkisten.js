function Waffenkisten_Statistik()
{
  var Vorlage_Letzte_Zeile = 103;

  var Zeitstempel = new Date()

  var SS_Export = SpreadsheetApp.openById("1nrgOIfkdLeFdzwmKF7awR1xx6cNVQHRDo6Q0qqIDTfE");
  var Sheet_Vorlage = SS_Export.getSheetByName("Vorlage");
  var Sheet_Statistik = SS_Export.getSheetByName("KW " + Kalenderwoche() + " "  + Zeitstempel.getFullYear());

  var Tag = Utilities.formatDate(Zeitstempel, "CET", "EEEE");
  var Uhrzeit = Utilities.formatDate(Zeitstempel, "CET", "HH:mm");

  var Stand_Waffenkisten = SpreadsheetApp.getActive().getSheetByName("Startseite").getRange("U10").getValue();
  var Stand_Waffenkisten_Speicher = Stand_Waffenkisten;
  var Stand_Vorher = PropertiesService.getScriptProperties().getProperty("wk_Stand");

  Stand_Waffenkisten = Math.floor(Stand_Vorher - Stand_Waffenkisten);
  Logger.log("Differenz Waffenkisten: " + Stand_Waffenkisten);

  if(Stand_Waffenkisten < 0)
  {
    Stand_Waffenkisten = 0;
  }

  PropertiesService.getScriptProperties().setProperty("wk_Stand", Stand_Waffenkisten_Speicher)

  var Spalte_1;
  var Spalte_2;

  if(Sheet_Statistik == null)
  {
    Sheet_Statistik = Sheet_Vorlage.copyTo(SS_Export);
    Sheet_Statistik.setName("KW " + Kalenderwoche() + " "  + Zeitstempel.getFullYear());
    Sheet_Statistik.getRange("B2").setValue(Sheet_Statistik.getName());
    Sheet_Statistik.getRange("C4").setValue(Zeitstempel);
  }

  switch(Tag)
  {
    case "Monday"   : Spalte_1 = 'B'; Spalte_2 = 'C'; break;
    case "Tuesday"  : Spalte_1 = 'E'; Spalte_2 = 'F'; break;
    case "Wednesday": Spalte_1 = 'H'; Spalte_2 = 'I'; break;
    case "Thursday" : Spalte_1 = 'K'; Spalte_2 = 'L'; break;
    case "Friday"   : Spalte_1 = 'N'; Spalte_2 = 'O'; break;
    case "Saturday" : Spalte_1 = 'Q'; Spalte_2 = 'R'; break;
    case "Sunday"   : Spalte_1 = 'T'; Spalte_2 = 'U'; break;
  }

  var Zeile = Sheet_Statistik.getRange(Spalte_1 + "3").getValue();

  if(Zeile > Vorlage_Letzte_Zeile)
  {
    return;
  }

  Sheet_Statistik.getRange(Spalte_1 + Zeile).setValue(Uhrzeit);
  Sheet_Statistik.getRange(Spalte_2 + Zeile).setValue(Stand_Waffenkisten);
}