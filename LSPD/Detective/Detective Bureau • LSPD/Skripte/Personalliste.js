function Personalliste(e)
{
  var Sheet_Personalliste = SpreadsheetApp.getActive().getSheetByName("Personalliste");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("J") && Zeile >= 6 && Zeile <= 38)
  {
    if(Value == undefined)
    {
      return Sheet_Personalliste.getRange(Zeile, Spalte).clearNote();
    }

    Sheet_Personalliste.getRange(Zeile, Spalte).setNote(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM. HH:mm"));
  }
  else if(Spalte == Spalte_in_Index("M") && Zeile >= 24 && Zeile <= 28)
  {
    if(Value == undefined)
    {
      return Sheet_Personalliste.getRange(Zeile, Spalte).clearNote();
    }
    
    Sheet_Personalliste.getRange(Zeile, Spalte).setNote(LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM. HH:mm"));
  }
}

function Personalliste_installOnEdit(e)
{
  var Sheet_Personalliste = SpreadsheetApp.getActive().getSheetByName("Personalliste");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("F") && Zeile >= 6 && Zeile <= 25 && Value == "TRUE")
  {
    Sheet_Personalliste.getRange(Zeile, Spalte).setValue(false);

    var Detective = Sheet_Personalliste.getRange("B" + Zeile).getValue();
    var Status = Sheet_Personalliste.getRange("E" + Zeile).getValue();

    if(Detective == "")
    {
      return;
    }

    var SS_Dienstblatt = SpreadsheetApp.openById(LSPD.ID_Dienstblatt);
    SS_Dienstblatt.getSheetByName("Log Stempeluhr").appendRow(["", Detective, 1, new Date(), ""]);

    if(Status == "⚫")
    {
      SS_Dienstblatt.getSheetByName("Log Einsatz").appendRow(["", Detective, "", new Date(), ""]);
    }
    else
    {
      SS_Dienstblatt.getSheetByName("Log Einsatz").appendRow(["", Detective, "Unmarked", new Date(), ""]);
    }
  }
}