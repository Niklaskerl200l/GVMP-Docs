function Aktenklaerung(e)
{
  var Sheet_Aktenklaerung = SpreadsheetApp.getActive().getSheetByName("Aktenklärungen");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 34)
  {
    if(Value != undefined && OldValue == undefined) // Neuer Eintrag
    {
      if(Sheet_Aktenklaerung.getRange("B5:B34").getValues().toString().toUpperCase().includes(Value.toString().trim().toUpperCase()) == true)
      {
        return SpreadsheetApp.getUi().alert("Fehler! Diese Person ist bereits in Klärung mit einem Fall!");
      }

      Sheet_Aktenklaerung.getRange(Zeile, Spalte).setValue(Value.toString().trim());

      var Zeitstempel = new Date();
      Zeitstempel.setDate(Zeitstempel.getDate() + 3);

      Sheet_Aktenklaerung.getRange("K" + Zeile).setValue(Zeitstempel);

      Sheet_Aktenklaerung.getRange("M" + Zeile + ":N" + Zeile).setValues([[new Date(), LSPD.Umwandeln()]]);
    }
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 5 && Zeile <= 34 && Value == "TRUE") // Kontaktaufnahme
  {
    Sheet_Aktenklaerung.getRange(Zeile, Spalte).setValue(false);

    var Kontaktaufnahme = Sheet_Aktenklaerung.getRange(Zeile, Spalte - 1).getDisplayValue();

    if(Kontaktaufnahme == "")
    {
      Kontaktaufnahme = (Utilities.formatDate(new Date(), "CET", "dd.MM.yy HH:mm"));
    }
    else
    {
      Kontaktaufnahme = (Kontaktaufnahme += ("\n" + Utilities.formatDate(new Date(), "CET", "dd.MM.yy HH:mm")));
    }

    Sheet_Aktenklaerung.getRange(Zeile, Spalte - 1).setValue(Kontaktaufnahme);
  }
  else if(Spalte == Spalte_in_Index("L") && Zeile >= 5 && Zeile <= 34 && Value == "TRUE") // In Klärung bis...
  {
    Sheet_Aktenklaerung.getRange(Zeile, Spalte).setValue(false);

    var Zeitstempel = Sheet_Aktenklaerung.getRange(Zeile, Spalte - 1).getValue();

    var UI = SpreadsheetApp.getUi();
    var Confirmation = UI.alert("Aktenklärung verlängern...", "Möchten Sie die Fallklärung vom " + (Utilities.formatDate(Zeitstempel, "CET", "dd.MM.yy HH:mm")) + " um 7 Tage verlängern?\nEs müssen ALLE Parteien damit einverstanden sein!", UI.ButtonSet.YES_NO);

    if(Confirmation == UI.Button.YES)
    {
      Zeitstempel.setDate(Zeitstempel.getDate() + 7);

      Sheet_Aktenklaerung.getRange(Zeile, Spalte - 1).setValue(Zeitstempel);
      Sheet_Aktenklaerung.getRange(Zeile, Spalte - 1).setNote("Aktenklärung um 7 Tage verlängert am " + (Utilities.formatDate(new Date(), "CET", "dd.MM.yy HH:mm") + " von " + LSPD.Umwandeln()));
    }
  }
  else if(Spalte == Spalte_in_Index("O") && Zeile >= 5 && Zeile <= 34 && Value == "TRUE") // Löschen
  {
    Sheet_Aktenklaerung.getRange(Zeile, Spalte).setValue(false);

    Sheet_Aktenklaerung.getRange("B" + Zeile + ":O" + Zeile).clearContent();
    Sheet_Aktenklaerung.getRange("B" + Zeile + ":O" + Zeile).clearNote();

    Sheet_Aktenklaerung.getRange("B5:O34").sort({column: Spalte_in_Index("M"), ascending: false});
  }
}

function Aktenklaerung_onOpen()
{
  var Sheet_Aktenklaerung = SpreadsheetApp.getActive().getSheetByName("Aktenklärungen");
  var Array_Aktenklaerung = Sheet_Aktenklaerung.getRange("E5:H34").getValues();

  var Benutzer = Umwandeln()[2];
  var Gefunden = 0;

  for(var i = 0; i < Array_Aktenklaerung.length; i++)
  {
    for(var o = 0; o < Array_Aktenklaerung[0].length; o++)
    {
      if(Array_Aktenklaerung[i][o] != "" && Array_Aktenklaerung[i][o].toString().toUpperCase().includes(Benutzer.toString().toUpperCase()) == true)
      {
        Gefunden = Gefunden + 1;
      }
    }
  }

  if(Gefunden > 0)
  {
    var UI = SpreadsheetApp.getUi();
    UI.alert("Aktenklärungen", "Sie wurden in " + Gefunden + (Gefunden == 1 ? " Aktenklärung" : " Aktenklärungen") + " als Beteiligte Person hinterlegt!", UI.ButtonSet.OK);
  }
}

function Aktenklaerung_Verjaehrung()
{
  var Sheet_Aktenklaerung = SpreadsheetApp.getActive().getSheetByName("Aktenklärungen");
  var Array_Aktenklaerung = Sheet_Aktenklaerung.getRange("B5:O34").getValues();

  var Zeitstempel = new Date();
  var Array_Fahndung = [];

  for(var i = 0; i < Array_Aktenklaerung.length; i++)
  {
    if(Array_Aktenklaerung[i][0] && new Date(Array_Aktenklaerung[i][9]) <= Zeitstempel)
    {
      Sheet_Aktenklaerung.getRange("B" + (i + 5) + ":N" + (i + 5)).clearContent();
      Array_Fahndung.push([Array_Aktenklaerung[i][0], Array_Aktenklaerung[i][1], Array_Aktenklaerung[i][2].toString() + "\n\nAKTENKLÄRUNG AUSGELAUFEN!", "-", false, new Date(), Array_Aktenklaerung[i][12]]);
    }
  }

  if(Array_Fahndung.length > 0)
  {
    var Sheet_Fahndungen = SpreadsheetApp.getActive().getSheetByName("Fahndungen");
    var Zeile_Fahndungen = Sheet_Fahndungen.getRange("B4").getValue();

    Sheet_Fahndungen.getRange(Zeile_Fahndungen, Spalte_in_Index("B"), Array_Fahndung.length, Array_Fahndung[0].length).setValues(Array_Fahndung);
    Sheet_Fahndungen.getRange("B6:H35").sort({column: Spalte_in_Index("G"), ascending: true});
  }
}