function Tickets(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Zeile >= 5 && Zeile <= 10 && Spalte == Spalte_in_Index("J") && Value == "TRUE")
  {
    var Sheet_Tickets = SpreadsheetApp.getActive().getSheetByName("Tickets");

    var Array_Archiv = Sheet_Tickets.getRange("B" + Zeile + ":J" + Zeile).getValues();

    if((Array_Archiv[0][5] == "Einsatz" || Array_Archiv[0][5] == "Sonstige") && Array_Archiv[0][6] == "")
    {
      SpreadsheetApp.getUi().alert("Gib bitte einen Grund in Notiz ein. Einsatz Tickets sind auch eigenverschuldet");
      return 0;
    }

    Sheet_Tickets.getRange("B" + Zeile + ":J" + Zeile).setValue("");

    Array_Archiv[0][8] = new Date();

    Sheet_Tickets.insertRowBefore(13).getRange("B13:J13").setValues(Array_Archiv);

    Sheet_Tickets.setActiveSelection(Sheet_Tickets.getRange("A1"));
  }

  else if(Zeile >= 5 && Zeile <= 10 && Spalte == Spalte_in_Index("B") && Value != undefined)
  {
    var Sheet_Tickets = SpreadsheetApp.getActive().getSheetByName("Tickets").getRange("I" + Zeile).setValue(LSPD.Umwandeln())
  }

  else if(Zeile >= 5 && Zeile <= 10 && Spalte == Spalte_in_Index("B") && Value == undefined)
  {
    var Sheet_Tickets = SpreadsheetApp.getActive().getSheetByName("Tickets").getRange("I" + Zeile).setValue("")
  }
}
