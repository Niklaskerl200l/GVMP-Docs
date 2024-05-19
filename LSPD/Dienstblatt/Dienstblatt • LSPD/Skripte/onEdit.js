var Fehler = false;

function onEdit(e) 
{
  LSPD.onEdit(e);

  var Sheet = e.source.getActiveSheet().getName();
  switch(Sheet)
  {
    case  "Startseite"      :   Startseite_NEU(e);          break;
    case  "Einsatz Archiv"  :   Einsatz_Dokumentation(e);   break;
    case  "DOC"             :   DOC(e);                     break;
    case  "Einsatz"         :   Szenario(e);                break;
    case  "Minderung"       :   Minderung(e);               break;
    case  "Wartung"         :   Wartung(e);                 break;
    case  "10-80 Auswertung":   Auswertung_1080(e);         break;
    case  "Fahndungen"      :   Fahndungen(e);              break;
    case  "Beschlagnahme"   :   Beschlagnahme(e);           break;
    case  "Parkkrallen (NEU)":  Fahrzeugsperren(e);         break;
    case  "Aktenklärungen"  :   Aktenklaerung(e);           break;
  }
  
  if(Fehler == true)
  {
    throw Error("Fehler!");
  }
}