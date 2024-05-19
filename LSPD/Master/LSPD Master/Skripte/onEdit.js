function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(SheetName == "Copy Master" && Spalte == Spalte_in_Index("I") && Zeile >= 5 && Value == "TRUE")
  {
    var Sheet_Copy = SpreadsheetApp.getActive().getSheetByName("Copy Master");

    var Array_Auswahl = Sheet_Copy.getRange("B" + Zeile + ":H" + Zeile).getValues();

    var Letzte_Zeile = Sheet_Copy.getRange("L3").getValue();

    Array_Auswahl[0][6] = "NAME" + "\n" + Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm");

    Sheet_Copy.getRange("L" + Letzte_Zeile + ":R" + Letzte_Zeile).setValues(Array_Auswahl);

    Sheet_Copy.getRange("B" + Zeile + ":J" + Zeile).setValue("").removeCheckboxes();

    LSPD_Copy.Properties_Delete(Utilities.formatDate(Array_Auswahl[0][0],SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy.MM.dd HH:mm:ss").toString(),"Script");

    Sheet_Copy.getRange("B5:J" + Sheet_Copy.getRange("B3").getValue()).sort(Spalte_in_Index("B"));
  }
}

function onOpen(e)
{
  var UI = SpreadsheetApp.getUi();

  UI.createMenu("Funktionen").addItem("Master Start","Open_Formular").addToUi();
}

function Open_Formular()
{
  var UI;

  UI = SpreadsheetApp.getUi();
  
  UI.showModalDialog(HtmlService.createHtmlOutput("<script>window.open(\'" + "https://forms.gle/k3cFmEA3SB1g46hd8" + "\');google.script.host.close();</script>"), "Open Tab")
}

LSPD_Copy.Eingabe_Test();