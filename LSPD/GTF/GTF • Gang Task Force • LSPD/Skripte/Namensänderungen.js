function Namensaenderung(e)
{
  var Sheet_Namensaenderung = SpreadsheetApp.getActive().getSheetByName("Namensänderungen");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile == 5 && Value != undefined && Value != "")
  {
    Sheet_Namensaenderung.getRange("D5:E5").setValues([[LSPD.Umwandeln(),new Date()]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile == 5 && (Value == undefined && Value == ""))
  {
    Sheet_Namensaenderung.getRange("D5:E5").setValues([[LSPD.Umwandeln(),new Date()]]);
  }
  else if(Spalte == Spalte_in_Index("F") && Zeile == 5 && Value == "TRUE")
  {
    var Array_Eingabe = Sheet_Namensaenderung.getRange("B5:E5").getValues();
    
    Eintrag_Check([[Array_Eingabe[0][1],Array_Eingabe[0][2]]],0,-1,-1,1,true);
  }
}
