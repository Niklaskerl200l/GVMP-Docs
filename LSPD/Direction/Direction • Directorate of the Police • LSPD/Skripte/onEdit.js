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
    case "Besprechung"          :   Besprechung(e);           break;
    case "Personal Master"      :   Personal_Master(e);       break;
    case "Entlassungen Archiv"  :   Entlassungen_Archiv(e);   break;
    case "Gespraeche"           :   Gespraeche(e);            break;
    case "Zeitbearbeitung"      :   Zeitbearbeitung(e);       break;
  }
}