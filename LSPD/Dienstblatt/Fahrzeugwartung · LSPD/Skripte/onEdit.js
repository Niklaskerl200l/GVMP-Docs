function onEdit(e)
{
  var Sheet = e.source.getActiveSheet().getName();
  switch(Sheet)
  {
    case "Kontrolle": Kontrolle(e); break;
    case "Offene Kontrollen": Offene_Kontrollen(e); break;
    case "Schäden": Schaeden(e); break;
    case "Tankkontrolle": Tankkontrolle(e); break;
  }
  
  LSPD.onEdit(e);
}