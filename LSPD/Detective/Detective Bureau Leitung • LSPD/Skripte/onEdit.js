function onEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();

  LSPD.onEdit(e);

  switch(SheetName)
  {
    case "Personal": Personal(e); break
    case  "Beschwerden Neu"             :   Beschwerden_Neu(e);             break;
    case  "Beschwerden In Bearbeitung"  :   Beschwerden_Berabeitung(e);     break;
    case  "Beschwerden Abgeschlossen"   :   Beschwerden_Abgeschlossen(e);    break;
  }
}