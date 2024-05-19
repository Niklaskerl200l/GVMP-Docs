function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  switch(SheetName)
  {
    case "Kursablauf": Zufallsfragen(e); break;
  }

  Logger.log("Benutzer: " + Session.getTemporaryActiveUserKey() + "\nSheet: " + SheetName + "\nZeile: " + Zeile + "\tSpalte: " + Spalte + "\nAlte Value: " + OldValue + "\nValue: " + Value);

  LSPD.onEdit(e);  

  if(SheetName.startsWith("WLD "))
  {
    WLD(e);
  }
}
