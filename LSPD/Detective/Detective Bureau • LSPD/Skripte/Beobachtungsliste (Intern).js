function Beobachtungsliste_Intern(e)
{
  var Sheet_Beobachtungsliste_Intern = SpreadsheetApp.getActive().getSheetByName("Beobachtungsliste (Intern)");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 5 && Zeile <= 34 && Value != undefined && e.oldValue == undefined)
  {
    Sheet_Beobachtungsliste_Intern.getRange("H" + Zeile + ":I" + Zeile).setValues([[new Date(), LSPD.Umwandeln()]]);
  }
  else if(Spalte == Spalte_in_Index("F") && Zeile >= 5 && Zeile <= 34)
  {
    if(Value == "TRUE")
    {
      Sheet_Beobachtungsliste_Intern.getRange(Zeile, Spalte).setNote("Angeordnet durch: " + LSPD.Umwandeln() + "\nam: " + Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy HH:mm"));
    }
    else if(Value == "FALSE")
    {
      Sheet_Beobachtungsliste_Intern.getRange(Zeile, Spalte).clearNote();
    }
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 5 && Zeile <= 34)
  {
    if(Value == "TRUE")
    {
      Sheet_Beobachtungsliste_Intern.getRange(Zeile, Spalte).setNote("Angeordnet durch: " + LSPD.Umwandeln() + "\nam: " + Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy HH:mm"));
    }
    else if(Value == "FALSE")
    {
      Sheet_Beobachtungsliste_Intern.getRange(Zeile, Spalte).clearNote();
    }
  }
  else if(Spalte == Spalte_in_Index("K") && Zeile >= 5 && Zeile <= 34 && Value == "TRUE")
  {
    Sheet_Beobachtungsliste_Intern.getRange(Zeile, Spalte).setValue(false);
    SpreadsheetApp.flush();

    var UI = SpreadsheetApp.getUi();
    var Confirmation = UI.alert("Beobachtungsliste...", "Möchten Sie diese Person von der Beobachtungsliste nehmen?", UI.ButtonSet.YES_NO);

    if(Confirmation == UI.Button.YES)
    {
      Sheet_Beobachtungsliste_Intern.getRange("B" + Zeile + ":I" + Zeile).setValues([["", "", "", "", false, false, "", ""]]);
      Sheet_Beobachtungsliste_Intern.getRange("F" + Zeile + ":G" + Zeile).clearNote();

      Beobachtungsliste_Intern_Sortieren();
    }
  }
}

function Beobachtungsliste_Intern_Sortieren()
{
  var Sheet_Beobachtungsliste_Intern = SpreadsheetApp.getActive().getSheetByName("Beobachtungsliste (Intern)");
  Sheet_Beobachtungsliste_Intern.getRange("B5:I34").sort({column: Spalte_in_Index("H"), ascending: true});
}