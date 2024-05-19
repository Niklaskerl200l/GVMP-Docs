function Schaeden(e)
{
  var Sheet_Schaeden = SpreadsheetApp.getActive().getSheetByName("Schäden");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("E") && Zeile >= 3 && Zeile <= 252 && Value == "TRUE")
  {
    Sheet_Schaeden.getRange(Zeile, Spalte).uncheck();

    var Array_Eintrag = Sheet_Schaeden.getRange(`B${Zeile}:D${Zeile}`).getValues();
    Array_Eintrag = Array_Eintrag[0];

    var Betrag = 0;

    if(Array_Eintrag[2].toString().includes("Reifenschaden") == true)
    {
      var Sheet_Meistermechaniker = SpreadsheetApp.getActive().getSheetByName("Meistermechaniker");
      var Array_Meistermechaniker = Sheet_Meistermechaniker.getRange("B3:B17").getValues();

      var Reifenkosten = 2500;
      var Reparaturbetrag = Array_Meistermechaniker.toString().includes(LSPD.Umwandeln()) == true ? Reparaturbetrag = 0 : Reparaturbetrag = 24000;
      Betrag = Math.floor(Betrag + (Reifenkosten * 4) + Reparaturbetrag);
    }

    if(Array_Eintrag[2].toString().includes("Motorschaden") == true)
    {
      Betrag = Math.floor(Betrag + 25000);
    }

    if(Array_Eintrag[2].toString().includes("Zylinderkopf") == true)
    {
      Betrag = Math.floor(Betrag + 12000);
    }

    if(Array_Eintrag[2].toString().includes("Fehlbetankung") == true)
    {
      Betrag = Math.floor(Betrag + 12500);
    }

    var Formular = `https://docs.google.com/forms/d/e/1FAIpQLSe6w0f9VHZNuP2cXkcA2rwdisk8DT5qtZs1j9nMzQFVVaOlbw/viewform?usp=pp_url&entry.1358042652=${LSPD.Umwandeln()}&entry.735174979=${"-"}&entry.953333295=${Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd")}&entry.1154255028=${Betrag}&entry.454638620=${`Fahrzeugwartung:+${Array_Eintrag[0] + "+" + Array_Eintrag[2]}`}&entry.1022989839=Ja&entry.681278857=${"Heindrick Flash"}&entry.1344497255=${"-"}`;

    Sheet_Schaeden.getRange(Zeile, Spalte).clearContent().removeCheckboxes().setFormula(`=HYPERLINK("${Formular}"; "KLICK")`);

    SpreadsheetApp.flush();
    Utilities.sleep(15000);

    Sheet_Schaeden.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();
  }
  else if(Spalte == Spalte_in_Index("F") && Zeile >= 3 && Zeile <= 252 && Value == "TRUE")
  {
    Sheet_Schaeden.getRange(Zeile, Spalte).uncheck();

    var Fahrzeug = Sheet_Schaeden.getRange("B" + Zeile).getValue();

    var Sheet_Schadensliste = SpreadsheetApp.getActive().getSheetByName("Schadensliste");
    var Array_Schadensliste = Sheet_Schadensliste.getRange("B3:F").getValues();

    for(var i = 0; i < Array_Schadensliste.length; i++)
    {
      if(Array_Schadensliste[i][0] != "" && Array_Schadensliste[i][0].toString() == Fahrzeug.toString() && Array_Schadensliste[i][4] == "")
      {
        Sheet_Schadensliste.getRange("F" + (i + 3)).setValue(`${Utilities.formatDate(new Date(), "CET", "dd.MM.yy")} ${LSPD.Umwandeln()}`);
      }
    }
  }
}