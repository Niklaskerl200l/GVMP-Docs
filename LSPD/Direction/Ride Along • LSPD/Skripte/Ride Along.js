function Ride_Along(e)
{
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Sheet_Ride_Along = SpreadsheetApp.getActive().getSheetByName("Ride Along");

  if(Spalte == Spalte_in_Index("B") && Zeile == 5 && Value != "" && Value != undefined)
  {
    var Array_Ausgabe;

    var Datum = new Date();

    Datum.setDate(Datum.getDate() + 1);

    Array_Ausgabe = [[new Date(), Datum, LSPD.Umwandeln()]];

    Sheet_Ride_Along.getRange("D5:F5").setValues(Array_Ausgabe);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile == 5 && (Value == "" || Value == undefined))
  {
    Sheet_Ride_Along.getRange("B5:G5").setValue("");
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile == 5 && Value == "TRUE")
  {
    var Sheet_Blacklist = SpreadsheetApp.getActive().getSheetByName("Blacklist");
    var Sheet_RAbisEinstellung = SpreadsheetApp.getActive().getSheetByName("RA bis Einstellung");

    var Array_Eingabe = Sheet_Ride_Along.getRange("B5:G5").getValues();
    var Array_Archiv = Sheet_Ride_Along.getRange("B10:G" + Sheet_Ride_Along.getLastRow()).getValues();
    var Array_Blacklist = Sheet_Blacklist.getRange("B3:B" + Sheet_Blacklist.getLastRow()).getValues();
    var Array_Einstellung = Sheet_RAbisEinstellung.getRange("B4:B" + Sheet_RAbisEinstellung.getLastRow()).getValues();

    Array_Eingabe[0][0] = Array_Eingabe[0][0].toString().replace("_"," ");
    Array_Eingabe[0][5] = false;

    var Letzte_Zeile = Sheet_Ride_Along.getRange("I3").getValue();

    if(Array_Eingabe[0][1] == "" || Array_Eingabe[0][1] == undefined)
    {
      SpreadsheetApp.getUi().alert("Fehler","Bitte gib eine Telefonnummer an!",SpreadsheetApp.getUi().ButtonSet.OK);
      return 0;
    }

    for(var i = 0; i < Array_Archiv.length; i++)
    {
      if(Array_Archiv[i][0] == Array_Eingabe[0][0])
      {
        var Datum = new Date(Array_Archiv[i][3]);

        Datum = Datum.setDate(Datum.getDate() + 3);

        if(Datum > new Date())
        {
          SpreadsheetApp.getUi().alert("Fehler","Person hatte vor weniger als 3 Tagen einen RA!",SpreadsheetApp.getUi().ButtonSet.OK);
          return 0;
        }
      }
    }

    for(var i = 0; i < Array_Blacklist.length; i++)
    {
      if(Array_Blacklist[i][0] == Array_Eingabe[0][0])
      {
        SpreadsheetApp.getUi().alert("Fehler","Person hat einen Aktiven Blacklist eintrag offen!",SpreadsheetApp.getUi().ButtonSet.OK);
        return 0;
      }
    }

    for(var i = 0; i < Array_Einstellung.length; i++)
    {
      if(Array_Einstellung[i][0] == Array_Eingabe[0][0])
      {
        SpreadsheetApp.getUi().alert("Fehler","Person ist schon RA bis Einstellung!",SpreadsheetApp.getUi().ButtonSet.OK);
        return 0;
      }
    }

    if(Letzte_Zeile >= 9)
    {
      SpreadsheetApp.getUi().alert("Fehler","Kein freier RA Platz mehr!",SpreadsheetApp.getUi().ButtonSet.OK);
      return 0;
    }

    Sheet_Ride_Along.insertRowAfter(9);

    Sheet_Ride_Along.getRange("B10:G10").setValues(Array_Eingabe);

    Sheet_Ride_Along.getRange("I" + Letzte_Zeile + ":L" + Letzte_Zeile).setValues([[Array_Eingabe[0][0],Array_Eingabe[0][1],Array_Eingabe[0][3],Array_Eingabe[0][4]]]);

    Sheet_Ride_Along.getRange("B5:G5").setValue("");
  }

  else if(Spalte == Spalte_in_Index("G") && Zeile >= 10 && Value == "TRUE")
  {
    Sheet_Ride_Along.deleteRow(Zeile);
  }

  else if(Spalte == Spalte_in_Index("M") && Zeile >= 5 && Zeile <= 8 && Value == "TRUE")
  {
    Sheet_Ride_Along.getRange("I" + Zeile + ":M" + Zeile).setValue("");
  }
}
