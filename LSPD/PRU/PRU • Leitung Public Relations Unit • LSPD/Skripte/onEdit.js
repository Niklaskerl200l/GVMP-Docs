function onEdit(e) 
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  Logger.log("Benutzer: " + Session.getTemporaryActiveUserKey() + "\nSheet: " + SheetName + "\nZeile: " + Zeile + "\tSpalte: " + Spalte + "\nAlte Value: " + OldValue + "\nValue: " + Value);

  LSPD.onEdit(e);

  switch(SheetName)
  {
    case  "Bewerbungen Forum"      :   Bewerbungen_Quellcode(e);      break;
    case  "Personal"      :   Personal(e);      break;
    case  "Export Abteilungen"      :   Abteilungen(e);      break;
  }
}