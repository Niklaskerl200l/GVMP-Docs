function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  Logger.log("Sheet: " + SheetName + "\nZeile: " + Zeile + "\tSpalte: " + Spalte + "\nAlte Value: " + OldValue + "\nValue: " + Value);
  
  LSPD.onEdit(e);

  switch(SheetName)
  {
    case "Auswertung" : Auswertung(e); break;
  }
}

function onOpen()
{
  LSPD.onOpen();
}