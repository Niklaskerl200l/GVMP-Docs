function onEdit(e)
{
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  LSPD.onEdit(e);
  
  switch(SheetName)
  {
    case "Auszahlungstabelle LSPD" : Auszahlungstabelle(e); break;
  }
}
