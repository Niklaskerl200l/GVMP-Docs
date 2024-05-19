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
    case "Ausbildungsblatt" : Ausbildungsblatt(e); break;
    case "OFC Theoriefragen" : ofcfragen(e); break;
    case "Entziehen" : Entziehen(e); break;
    case  "Kalender" :    Kalender(e); break;
  
  }


  if(SheetName.substr(0,21) == "Besprechungsprotokoll" && SheetName != "Besprechungsprotokoll Vorlage")
  {
    Besprechung_Anwesenheit(e);
  }
}
