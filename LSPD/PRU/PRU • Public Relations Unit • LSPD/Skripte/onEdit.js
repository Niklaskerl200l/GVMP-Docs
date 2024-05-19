function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();

  LSPD.onEdit(e);

  switch(SheetName)
  {
    case "Personal": Personal(e); break;
    case "Bewerber": Bewerber(e); break;
    case "Export Abteilungen": Export_Abteilungen(e); break
  }
}
