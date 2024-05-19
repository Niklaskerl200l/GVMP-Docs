function Startseite(e) 
{
  var Sheet = e.source.getActiveSheet();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("G") && Zeile >= 6 && Zeile <= 17 && Value != null)   //  auto. Name
  {
    Sheet.getRange(Zeile,Spalte_in_Index("N")).setValue(LSPD.Umwandeln());
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 6 && Zeile <= 17 && Value == null)   //  auto. Löschen Name
  {
    Sheet.getRange(Zeile,Spalte_in_Index("N")).setValue("");
  }
  if(Spalte == Spalte_in_Index("G") && Zeile >= 20 && Zeile <= 31 && Value != null)   // auto. Name
  {
    Sheet.getRange(Zeile,Spalte_in_Index("N")).setValue(LSPD.Umwandeln());
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 20 && Zeile <= 31 && Value == null)   // auto. Löschen Name
  {
    Sheet.getRange(Zeile,Spalte_in_Index("N")).setValue("");
  }
}
