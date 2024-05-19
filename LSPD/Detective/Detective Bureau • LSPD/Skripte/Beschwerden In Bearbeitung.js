function Beschwerden_In_Bearbeitung(e)
{
  var Sheet_Beschwerde_In_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden In Bearbeitung");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("I") && Zeile >= 3 && Zeile <= 34 && Value != undefined)
  {
    var Sheet_Strafkatalog = SpreadsheetApp.getActive().getSheetByName("Bußgeldkatalog");
    var Array_Strafkatalog = Sheet_Strafkatalog.getRange("H17:K29").getValues();

    var TVO_Rang = Sheet_Beschwerde_In_Bearbeitung.getRange("E" + Zeile).getValue();

    var Gefunden = false;
    for(var i = 0; i < Array_Strafkatalog.length; i++)
    {
      if(Array_Strafkatalog[i][0].toString() == TVO_Rang.toString())
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == true)
    {
      switch(Value)
      {
        case "Geldstrafe 1": Sheet_Beschwerde_In_Bearbeitung.getRange("J" + Zeile).setValue(Array_Strafkatalog[i][1]); break;
        case "Geldstrafe 2": Sheet_Beschwerde_In_Bearbeitung.getRange("J" + Zeile).setValue(Array_Strafkatalog[i][2]); break;
        case "Geldstrafe 3": Sheet_Beschwerde_In_Bearbeitung.getRange("J" + Zeile).setValue(Array_Strafkatalog[i][3]); break;
        case "Geldstrafe 4": Sheet_Beschwerde_In_Bearbeitung.getRange("J" + Zeile).setValue(250000); break;
        case "Geldstrafe 4 1/2": Sheet_Beschwerde_In_Bearbeitung.getRange("J" + Zeile).setValue(Math.floor(250000 + (250000 / 2))); break;
        case "Degradierung": Sheet_Beschwerde_In_Bearbeitung.getRange("J" + Zeile).setValue("Degradierung auf Rang " + Math.floor(TVO_Rang - 1).toString()); break;
        case "Suspendierung": Sheet_Beschwerde_In_Bearbeitung.getRange("J" + Zeile).setValue("Suspendierung von 7 Tagen"); break;
      }
    }
  }
  else if(Spalte == Spalte_in_Index("AC") && Zeile >= 3 && Zeile <= 34 && Value == "TRUE")
  {
    Sheet_Beschwerde_In_Bearbeitung.getRange(Zeile, Spalte).clearContent().removeCheckboxes();

    var Array_Beschwerden_In_Bearbeitung = Sheet_Beschwerde_In_Bearbeitung.getRange("B" + Zeile + ":AD" + Zeile).getValues();
    Array_Beschwerden_In_Bearbeitung = Array_Beschwerden_In_Bearbeitung[0];

    if(Array_Beschwerden_In_Bearbeitung[5] == "Ausstehend" || Array_Beschwerden_In_Bearbeitung[5] == "In Bearbeitung" || Array_Beschwerden_In_Bearbeitung[5] == "")
    {
      Sheet_Beschwerde_In_Bearbeitung.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
      return Browser.msgBox("Fehler!\nSie müssen den Status der Sanktion erst umsetzen!");
    }
    else if(Array_Beschwerden_In_Bearbeitung[6] == "In Bearbeitung")
    {
      Sheet_Beschwerde_In_Bearbeitung.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
      return Browser.msgBox("Fehler!\nSie müssen den Status der Beschwerde erst umsetzen!");
    }
    else if(Array_Beschwerden_In_Bearbeitung[7] == "")
    {
      Sheet_Beschwerde_In_Bearbeitung.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
      return Browser.msgBox("Fehler!\nGeben Sie eine Sanktion angeben!");
    }
    else if(Array_Beschwerden_In_Bearbeitung[9] == "")
    {
      Sheet_Beschwerde_In_Bearbeitung.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
      return Browser.msgBox("Fehler!\nGeben Sie einen Strike angeben!");
    }

    var Sheet_Beschwerden_Abgeschlossen = SpreadsheetApp.getActive().getSheetByName("Beschwerden Abgeschlossen");

    Sheet_Beschwerden_Abgeschlossen.insertRowAfter(3);
    Sheet_Beschwerden_Abgeschlossen.getRange("B4:AC4").setValues([[
      new Date(),
      Array_Beschwerden_In_Bearbeitung[0],
      Array_Beschwerden_In_Bearbeitung[2],
      Array_Beschwerden_In_Bearbeitung[3],
      Array_Beschwerden_In_Bearbeitung[4],
      Array_Beschwerden_In_Bearbeitung[5],
      Array_Beschwerden_In_Bearbeitung[6],
      Array_Beschwerden_In_Bearbeitung[7],
      Array_Beschwerden_In_Bearbeitung[8],
      Array_Beschwerden_In_Bearbeitung[9],
      Array_Beschwerden_In_Bearbeitung[10],
      Array_Beschwerden_In_Bearbeitung[11],
      Array_Beschwerden_In_Bearbeitung[12],
      Array_Beschwerden_In_Bearbeitung[13],
      Array_Beschwerden_In_Bearbeitung[14],
      Array_Beschwerden_In_Bearbeitung[15],
      Array_Beschwerden_In_Bearbeitung[16],
      Array_Beschwerden_In_Bearbeitung[17],
      Array_Beschwerden_In_Bearbeitung[18],
      Array_Beschwerden_In_Bearbeitung[19],
      Array_Beschwerden_In_Bearbeitung[20],
      Array_Beschwerden_In_Bearbeitung[21],
      Array_Beschwerden_In_Bearbeitung[22],
      Array_Beschwerden_In_Bearbeitung[23],
      "-",
      Array_Beschwerden_In_Bearbeitung[25],
      Array_Beschwerden_In_Bearbeitung[26],
      false
    ]]);

    Sheet_Beschwerde_In_Bearbeitung.getRange("B" + Zeile + ":AD" + Zeile).clearContent().setFontColor("white").setBackground(null);

    if(Array_Beschwerden_In_Bearbeitung[7].toString().includes("Geldstrafe") == true)
    {
      var Sheet_Geldverwaltung = SpreadsheetApp.getActive().getSheetByName("Geldverwaltung");

      Sheet_Geldverwaltung.insertRowAfter(3);
      Sheet_Geldverwaltung.getRange("B4:J4").setValues([[Array_Beschwerden_In_Bearbeitung[0], Array_Beschwerden_In_Bearbeitung[2], Array_Beschwerden_In_Bearbeitung[8], "", "", false, "", false, ""]]);

      Sheet_Geldverwaltung.getRangeList(["G4", "I4"]).setValue(false).insertCheckboxes();
      Sheet_Geldverwaltung.getRange("E4").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(SpreadsheetApp.getActive().getSheetByName("Personalliste").getRange("B6:B25")).build());

      Sheet_Geldverwaltung.setActiveSelection("E4");
    }

    Sheet_Beschwerde_In_Bearbeitung.getRange("B3:AD34").sort([{column: Spalte_in_Index("C"), ascending: true}, {column: Spalte_in_Index("AD"), ascending: true}]);
  }
}