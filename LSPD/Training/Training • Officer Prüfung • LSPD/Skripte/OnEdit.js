function onEdit(e) 
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();

  LSPD.onEdit(e);
  
  switch(SheetName)
  {
    case  "Neue Termine"      :   Neue_Termine(e);      break;
  }
}