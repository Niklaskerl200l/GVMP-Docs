function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  Logger.log("SeMiKoLoNsKy");

  LSPD.onEdit(e);

  switch(SheetName)
  {
    case "Personal": Personal(e); break
    case  "Info"   :   Info(e);   break;
    case  "Bewerber"   :   Bewerber(e);   break;
    case  "Export Abteilungen"      :   Abteilungen(e);      break;
  }
}
