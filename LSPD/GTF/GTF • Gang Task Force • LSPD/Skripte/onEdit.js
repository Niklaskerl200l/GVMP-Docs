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
    case  "Startseite"   :   Startseite(e);   break;
    case  "Dokumentation" : Dokumentation(e); break;
    case  "Bearbeiten" : Bearbeiten(e); break;
    case  "Mitgliederlisten" : Mitgliederliste(e); break;
    case  "Namensänderungen" : Namensaenderung(e); break;
    case  "Zugehörigkeit"   :   Zugehoerigkeit(e);   break;
    case  "Akteneinträge" : Akteneintrag(e);  break;
    case  "Fraktionsgespräche" : FG_Uebersicht(e); break;
    case  "Fraktionsfahrzeuge" : Fraktionsfahrzeuge(e); break;
    case  "Rückrufanfragen"     : Rueckrufe(e); break;
  }  

  if(SheetName.substr(0,21) == "Besprechungsprotokoll" && SheetName != "Besprechungsprotokoll Vorlage")
  {
    Besprechung_Anwesenheit(e);
  }
  else  if(SheetName.substr(0,3) == "FG ")
  {
    FG_Neu(e);
  }
}
