function Anwesenheit() // Niklas_Kerl®
{  
  //*********** <Manuele Eingabe> **********//
  //                                        //
        var Vorlage_Letzte_Zeile  = 104;    //
  //                                        //
  //********** <\Manuele Eingabe> **********//

  var Datum = new Date();

  var SS_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Beamtenzahl);
  var Sheet_BSE = SS_Export.getSheetByName("Beamten Statistik Export");
  var Sheet_Startseite = SpreadsheetApp.getActive().getSheetByName("Startseite");
  var Sheet = SS_Export.getSheetByName(Datum.getFullYear() + " KW " + Kalenderwoche());
  
  var Tag = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "EEEE");
  var Uhrzeit = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "HH:mm");
  
  var Beamte = Number(Sheet_Startseite.getRange("N20").getValue()) + Number(Sheet_Startseite.getRange("R20").getValue());
  
  var Spalte1, Spalte2;

  if(Sheet == null)
  {
    var Vorlage = SS_Export.getSheetByName("Auswertung_Vorlage");
    
    Vorlage.copyTo(SS_Export).setName(Datum.getFullYear() + " KW " + Kalenderwoche());
    
    Sheet_BSE.getRange("C2").setValue(Kalenderwoche());

    Sheet = SS_Export.getSheetByName(Datum.getFullYear() + " KW " + Kalenderwoche());
  }

  switch(Tag) 
  {
    case "Monday"   : Spalte1 = 'B'; Spalte2 = 'C'; break;
    case "Tuesday"  : Spalte1 = 'E'; Spalte2 = 'F'; break;
    case "Wednesday": Spalte1 = 'H'; Spalte2 = 'I'; break;
    case "Thursday" : Spalte1 = 'K'; Spalte2 = 'L'; break;
    case "Friday"   : Spalte1 = 'N'; Spalte2 = 'O'; break;
    case "Saturday" : Spalte1 = 'Q'; Spalte2 = 'R'; break;
    case "Sunday"   : Spalte1 = 'T'; Spalte2 = 'U'; break;
  }
  
  var Letzte_Zeile = Sheet.getRange(Spalte1 + '5').getValue();
  var Feld_Datum   = Sheet.getRange(Spalte2 + '6').getValue();
  
  if(Letzte_Zeile > Sheet.getLastRow())
  {
    Sheet.insertRowAfter(Letzte_Zeile);
  }
  
  if(Letzte_Zeile > Vorlage_Letzte_Zeile - 1)
  {
    Sheet.getRange(Spalte1 + Letzte_Zeile).setBackground('#999999');
    Sheet.getRange(Spalte1 + Letzte_Zeile).setHorizontalAlignment("center");
    Sheet.getRange(Spalte1 + Letzte_Zeile).setBorder(true, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.DOTTED);
    Sheet.getRange(Spalte1 + Letzte_Zeile).setBorder(null, true, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    Sheet.getRange(Spalte2 + Letzte_Zeile).setBackground('#999999');
    Sheet.getRange(Spalte2 + Letzte_Zeile).setHorizontalAlignment("center");
    Sheet.getRange(Spalte2 + Letzte_Zeile).setBorder(true, null, null, null, null, null, 'black', SpreadsheetApp.BorderStyle.DOTTED);
    Sheet.getRange(Spalte2 + Letzte_Zeile).setBorder(null, null, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  
  Sheet.getRange(Spalte1 + Letzte_Zeile).setValue(Uhrzeit);
  Sheet.getRange(Spalte2 + Letzte_Zeile).setValue(Beamte);
  
  if(Feld_Datum == "")
  {
    Sheet.getRange(Spalte2 + '6').setValue(Datum);
  }
}