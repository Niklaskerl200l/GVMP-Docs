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
    case  "Personalliste"               :   Personalliste(e);               break;
    case  "Beschwerden Neu"             :   Beschwerden_Neu(e);             break;
    case  "Beschwerden In Bearbeitung"  :   Beschwerden_In_Bearbeitung(e);  break;
    case  "Streifenkontrolle"           :   Streifenkontrolle(e);           break;
    case  "Geldverwaltung"              :   Geldverwaltung(e);              break;
    case  "Beobachtungsliste"           :   Beobachtungsliste();            break;
    case  "Beschwerden Abgeschlossen"   :   Beschwerden_Abgeschlossen(e);   break;
    case  "Minus Leitstelle"            :   Minus_Leitstelle(e);            break;
    case  "Plus Leitstelle"             :   LST_Plus(e);                    break;
    case  "Rückrufanfragen"             :   Rueckrufe(e);                   break;
    case  "LST Rechner"                 :   LST_Rechner(e);                 break;
    case  "Kalender"                    :   Kalender(e);                    break;
    case  "Beobachtungsliste (Intern)"  :   Beobachtungsliste_Intern(e);    break;
    case  "Anwesenheitskontrolle"       :   Anwesenheitskontrolle(e);       break;
    case  "Fuhrparkmeldungen"           :   Fuhrparkmeldungen(e);           break;
  }

  if(SheetName.substr(0,21) == "Besprechungsprotokoll" && SheetName != "Besprechungsprotokoll Vorlage")
  {
    Besprechung_Anwesenheit(e);
  }
}

function installOnEdit(e)
{
  var Sheet = e.source.getActiveSheet();
  var SheetName = Sheet.getName();

  switch(SheetName)
  {
    case "LST Suche": LST_Suche(e); break;
    case "Personalliste": Personalliste_installOnEdit(e); break;
    case "Einträgefinder (NEU)": Eintraegefinder(e); break;
  }
}