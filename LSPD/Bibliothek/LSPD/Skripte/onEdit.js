function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Value != undefined && Value != "" && Value != null)
  {
    var Value_Upper = Value.toString().toUpperCase();

    if(Zeile == 1 && Spalte == 1 && (Value_Upper == "/THEME" || Value_Upper == "/DESIGN" || Value_Upper == "/DARK" || Value_Upper == "DARKMODE" || Value_Upper == "DARK MODE"))
    {
      Sheet.getRange(Zeile,Spalte).setValue(OldValue);
      
      Design_LSPD_Dark();
    }
  }

  Logger.log("Sheet: " + SheetName + "\nZeile: " + Zeile + "\tSpalte: " + Spalte + "\nAlte Value: " + OldValue + "\nValue: " + Value);
  Logger.log("Session Key: " + Session.getTemporaryActiveUserKey());
}