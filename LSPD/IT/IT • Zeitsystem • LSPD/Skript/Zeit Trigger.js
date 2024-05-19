function Tag_Trigger()
{
  var Sheet_Zeitsystem = SpreadsheetApp.getActive().getSheetByName("Zeitsystem");

  var Letzte_Zeile = Sheet_Zeitsystem.getRange("B1").getValue();

  var Array_Woche = Sheet_Zeitsystem.getRange("G4:G" + Letzte_Zeile).getValues();
  var Array_Monat = Sheet_Zeitsystem.getRange("L4:L" + Letzte_Zeile).getValues();

  var Array_Neu_Woche = [];
  var Array_Neu_Monat = [];

  for(var y = 0; y < Array_Woche.length; y++)   // Erzeugen neuer Fromeln der Woche
  {
    Array_Neu_Woche[y] = new Array();

    Array_Neu_Woche[y][0] = "=SUM(\"" + Zeit_Dauer(Array_Woche[y][0])[2] + "\";F" + (y + 4) + ")";
  }

  for(var y = 0; y < Array_Monat.length; y++)   // Erzeugen neuer Fromeln des Monat
  {
    Array_Neu_Monat[y] = new Array();

    Array_Neu_Monat[y][0] = "=SUM(\"" + Zeit_Dauer(Array_Monat[y][0])[2] + "\";F" + (y + 4) + ")";
  }

  Sheet_Zeitsystem.getRange("F4:F" + Letzte_Zeile).setValue("0:00");
  Sheet_Zeitsystem.getRange("F4:F").setNumberFormat("[h]:mm");
  Sheet_Zeitsystem.getRange("G4:G" + Letzte_Zeile).setFormulas(Array_Neu_Woche);
  Sheet_Zeitsystem.getRange("L4:L" + Letzte_Zeile).setFormulas(Array_Neu_Monat);


  Archivieren_Leitstelle();
}

function Woche_Trigger()
{
  var Sheet_Zeitsystem = SpreadsheetApp.getActive().getSheetByName("Zeitsystem");

  var Letzte_Zeile = Sheet_Zeitsystem.getRange("B1").getValue();

  var Array_Woche = Sheet_Zeitsystem.getRange("G4:I" + Letzte_Zeile).getValues();
  var Array_Neu_Woche = [];

  for(var y = 0; y < Array_Woche.length; y++)   // Erzeugen neuer Fromeln der Woche
  {
    Array_Neu_Woche[y] = new Array();

    Array_Neu_Woche[y][0] = "=SUM(\"0:00\";F" + (y + 4) + ")";
  }

  Sheet_Zeitsystem.getRange("G4:G" + Letzte_Zeile).setFormulas(Array_Neu_Woche);
  Sheet_Zeitsystem.getRange("H4:J" + Letzte_Zeile).setValues(Array_Woche);
}

function Monat_Trigger()
{
  var Sheet_Zeitsystem = SpreadsheetApp.getActive().getSheetByName("Zeitsystem");

  var Letzte_Zeile = Sheet_Zeitsystem.getRange("B1").getValue();

  var Array_Monat = Sheet_Zeitsystem.getRange("L4:N" + Letzte_Zeile).getValues();
  var Array_Monat_Letzer = Sheet_Zeitsystem.getRange("O4:O" + Letzte_Zeile).getValues();
  var Array_Monat_Archiv = Sheet_Zeitsystem.getRange("P4:P" + Letzte_Zeile).getValues();

  var Array_Neu_Monat = [];

  for(var y = 0; y < Array_Monat_Letzer.length; y++)
  {
    Array_Neu_Monat[y] = new Array();

    Array_Monat_Archiv[y][0].setMinutes(Array_Monat_Archiv[y][0].getMinutes() + Array_Monat_Letzer[y][0].getMinutes());
    Array_Monat_Archiv[y][0].setHours(Array_Monat_Archiv[y][0].getHours() + Number(Zeit_Dauer(Array_Monat_Letzer[y][0])[0]));
    Array_Monat_Archiv[y][0] = Zeit_Dauer(Array_Monat_Archiv[y][0])[2];

    Array_Neu_Monat[y][0] = "=SUM(\"0:00\";F" + (y + 4) + ")";
  }

  Sheet_Zeitsystem.getRange("L4:L" + Letzte_Zeile).setValues(Array_Neu_Monat);
  Sheet_Zeitsystem.getRange("M4:O" + Letzte_Zeile).setValues(Array_Monat);
  Sheet_Zeitsystem.getRange("P4:P" + Letzte_Zeile).setValues(Array_Monat_Archiv);


  //-------------------- Leitstelle -----------------------------------//

  var Array_Monat = Sheet_Zeitsystem.getRange("T4:V" + Letzte_Zeile).getValues();
  var Array_Monat_Letzer = Sheet_Zeitsystem.getRange("W4:W" + Letzte_Zeile).getValues();
  var Array_Monat_Archiv = Sheet_Zeitsystem.getRange("X4:X" + Letzte_Zeile).getValues();

  for(var y = 0; y < Array_Monat_Letzer.length; y++)
  {
    Array_Neu_Monat[y] = new Array();

    Array_Monat_Archiv[y][0].setMinutes(Array_Monat_Archiv[y][0].getMinutes() + Array_Monat_Letzer[y][0].getMinutes());
    Array_Monat_Archiv[y][0].setHours(Array_Monat_Archiv[y][0].getHours() + Number(Zeit_Dauer(Array_Monat_Letzer[y][0])[0]));
    Array_Monat_Archiv[y][0] = Zeit_Dauer(Array_Monat_Archiv[y][0])[2];
  }

  Sheet_Zeitsystem.getRange("T4:T" + Letzte_Zeile).setValue("0:00");
  Sheet_Zeitsystem.getRange("T4:T").setNumberFormat("[h]:mm");
  Sheet_Zeitsystem.getRange("U4:W" + Letzte_Zeile).setValues(Array_Monat);
  Sheet_Zeitsystem.getRange("X4:X" + Letzte_Zeile).setValues(Array_Monat_Archiv);
}
