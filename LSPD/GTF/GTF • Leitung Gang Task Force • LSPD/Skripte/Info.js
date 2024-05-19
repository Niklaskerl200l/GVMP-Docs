function Info(e) 
{
  var Sheet = e.source.getActiveSheet();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 3 && Zeile <= 21 && Value != null)   // auto. Name
  {
    Sheet.getRange(Zeile,Spalte_in_Index("D")).setValue(LSPD.Umwandeln());
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 3 && Zeile <= 21 && Value == null)   // auto. Löschen Name
  {
    Sheet.getRange(Zeile,Spalte_in_Index("D")).setValue("");
  }
}
