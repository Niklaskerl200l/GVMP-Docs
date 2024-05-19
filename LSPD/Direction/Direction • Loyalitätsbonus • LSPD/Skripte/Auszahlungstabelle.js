function Auszahlungstabelle(e)
{
  var Sheet_Auszahlung = SpreadsheetApp.getActive().getSheetByName("Auszahlungstabelle LSPD");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("H") && Zeile >= 5 && Value == "TRUE")         // Beantragt
  {
    Sheet_Auszahlung.getRange("I" + Zeile).setValue(new Date());

    var Name_Chief = "Heindrick+Flash";
    var Tel_Chief = "80949";
    var Datum_Auszahlung = Utilities.formatDate(Sheet_Auszahlung.getRange("F" + Zeile).getValue(),"CET","yyyy-MM-dd")
    var Summe = Sheet_Auszahlung.getRange("G" + Zeile).getValue()   //.toLocaleString('de-DE', { style: 'currency', currency: 'USD', minimumFractionDigits: 0  }).replace(" ","+");
    var Name = Sheet_Auszahlung.getRange("B" + Zeile).getValue();
    var Monate = Sheet_Auszahlung.getRange("E" + Zeile).getValue();
    var Einstellung = Sheet_Auszahlung.getRange("C" + Zeile).getValue();
    var Text = "LSPD Loyalitätsbonus\nName: " + Name + "\nMonat: " + Monate + "\nEinstellung: " + Utilities.formatDate(Einstellung,"CET","dd.MM.yyyy")

    var URL = 'https://docs.google.com/forms/d/e/1FAIpQLSe6w0f9VHZNuP2cXkcA2rwdisk8DT5qtZs1j9nMzQFVVaOlbw/viewform?usp=pp_url&entry.1358042652='+Name_Chief+'&entry.735174979='+Tel_Chief+'&entry.953333295='+Datum_Auszahlung+'&entry.1154255028='+Summe+'&entry.454638620='+Text+'&entry.1022989839=Ja&entry.681278857='+Name_Chief+'&entry.1344497255='

    Sheet_Auszahlung.getRange(Zeile,Spalte).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")").setFontColor("#0000ff").removeCheckboxes();

    Utilities.sleep(10000);

    Sheet_Auszahlung.getRange(Zeile,Spalte).setFontColor("#f3f3f3").insertCheckboxes().setValue(true);
  }
  else if(Spalte == Spalte_in_Index("H") && Zeile >= 5 && Value == "FALSE")   // Beantragt
  {
    Sheet_Auszahlung.getRange("I" + Zeile).setValue("");
  }


  else if(Spalte == Spalte_in_Index("J") && Zeile >= 5 && Value == "TRUE")   // Erhalten
  {
    Sheet_Auszahlung.getRange("K" + Zeile).setValue(new Date());
  }
  else if(Spalte == Spalte_in_Index("J") && Zeile >= 5 && Value == "FALSE")   // Erhalten
  {
    Sheet_Auszahlung.getRange("K" + Zeile).setValue("");
  }


  else if(Spalte == Spalte_in_Index("L") && Zeile >= 5 && Value == "TRUE")   // Ausgezahlt
  {
    var Datum = new Date(Sheet_Auszahlung.getRange("F" + Zeile).getValue());

    Datum.setMonth(Datum.getMonth() + 3);

    
    Sheet_Auszahlung.getRange("F" + Zeile).setValue(Datum);
    Sheet_Auszahlung.getRange("H" + Zeile + ":M" + Zeile).setValues([["","","","","",new Date()]])
    Sheet_Auszahlung.getRange("O5").setValue(Sheet_Auszahlung.getRange("O5").getValue() + Sheet_Auszahlung.getRange("G" + Zeile).getValue());
  }


  else if(Spalte == Spalte_in_Index("T") && Zeile >= 29 && Value == "TRUE")  // Rückzahlung
  {
    Sheet_Auszahlung.getRange("Q" + Zeile + ":T" + Zeile).setValue("");
    Sheet_Auszahlung.getRange("T" + Zeile).removeCheckboxes();

    Sheet_Auszahlung.getRange("Q29:T200").sort(Spalte_in_Index("R"));
  }
}

function Sortieren_Auszahlungstabelle()
{
  var Sheet_Auszahlung = SpreadsheetApp.getActive().getSheetByName("Auszahlungstabelle LSPD");

  Sheet_Auszahlung.getRange("B5:M").sort({column: Spalte_in_Index("F"),ascending: true});
}