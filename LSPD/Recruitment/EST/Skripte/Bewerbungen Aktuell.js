function Bewerbungen_Aktuell(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Zeile >= 5 && Zeile <= 35 && Spalte == Spalte_in_Index("O") && Value == "TRUE")
  {
    var Sheet_Bewerbungen_Aktuell = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Aktuell");
    var Sheet_Bewerbungen_Archiv = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Archiv");

    var Array_Aktuell = Sheet_Bewerbungen_Aktuell.getRange("B" + Zeile + ":N" + Zeile).getValues();

    if(Array_Aktuell[0][1] == "Abgelehnt")
    {
      var Array_Archiv =
      [[
        Array_Aktuell[0][0],
        Array_Aktuell[0][1],
        Array_Aktuell[0][2],
        Array_Aktuell[0][3],
        Array_Aktuell[0][4],
        Array_Aktuell[0][5],
        Array_Aktuell[0][6],
        Array_Aktuell[0][7],
        Array_Aktuell[0][8],
        Array_Aktuell[0][9],
        Array_Aktuell[0][10],
        Array_Aktuell[0][11],
        Array_Aktuell[0][12],
        "-"
      ]];

      Sheet_Bewerbungen_Aktuell.getRange("B" + Zeile + ":N" + Zeile).setValue("");
      Sheet_Bewerbungen_Aktuell.getRange("O" + Zeile).removeCheckboxes();

      Sheet_Bewerbungen_Archiv.getRange("B5:O35").sort({column: 5, ascending: false});

      var Zeile_Archiv = Sheet_Bewerbungen_Archiv.getLastRow() + 1;

      Sheet_Bewerbungen_Archiv.getRange("B" + Zeile_Archiv + ":O" + Zeile_Archiv).setValues(Array_Archiv);

      Sheet_Bewerbungen_Archiv.setActiveSelection("B" + Zeile_Archiv);
    }
    else
    {
      var SS = SpreadsheetApp.getActiveSpreadsheet();
      var Sheet_Bewerbung_Vorlage = SpreadsheetApp.getActive().getSheetByName("Bewerbungsgespräch Vorlage");

      var Array_Zufallsfragen = Zufallsfragen_Bewerbung();

      var Array_Aktuell = Sheet_Bewerbungen_Aktuell.getRange("B" + Zeile + ":N" + Zeile).getValues();

      Sheet_Bewerbungen_Aktuell.getRange("B" + Zeile + ":O" + Zeile).setValue("");
      Sheet_Bewerbungen_Aktuell.getRange("O" + Zeile).removeCheckboxes();

      Sheet_Bewerbung_Vorlage.copyTo(SS).setName("Bewerbungsgespräch " + Array_Aktuell[0][0]);

      var Sheet_Bewerbung = SpreadsheetApp.getActive().getSheetByName("Bewerbungsgespräch " + Array_Aktuell[0][0]);

      Sheet_Bewerbung.setActiveSelection("B8");

      var Array_Bewerbung = 
      [[
        Array_Aktuell[0][0],
        "",
        Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy"),
        "",
        LSPD.Umwandeln(),
        "",
        "",
        "",
        Array_Aktuell[0][4],
        "",
        Array_Aktuell[0][5],
        Array_Aktuell[0][6],
        Array_Aktuell[0][7],
        Array_Aktuell[0][8],
        Array_Aktuell[0][9],
        "",
        Array_Aktuell[0][10],
        "",
        "",
        Array_Aktuell[0][11],
        Array_Aktuell[0][12]
      ]];

      Sheet_Bewerbung.getRange("B2").setValue("Bewerbungsgespräch " + Array_Aktuell[0][0]);
      Sheet_Bewerbung.getRange("B5:V5").setValues(Array_Bewerbung);
      
      Sheet_Bewerbung.getRange("H13").setValue(Array_Zufallsfragen[0][0]);
      Sheet_Bewerbung.getRange("N13").setValue(Array_Zufallsfragen[1][0]);
      Sheet_Bewerbung.getRange("B16").setValue(Array_Zufallsfragen[2][0]);
      Sheet_Bewerbung.getRange("H16").setValue(Array_Zufallsfragen[3][0]);
      Sheet_Bewerbung.getRange("N16").setValue(Array_Zufallsfragen[4][0]);

      Sheet_Bewerbung.getRange("H14").setNote(Array_Zufallsfragen[0][1]);
      Sheet_Bewerbung.getRange("N14").setNote(Array_Zufallsfragen[1][1]);
      Sheet_Bewerbung.getRange("B17").setNote(Array_Zufallsfragen[2][1]);
      Sheet_Bewerbung.getRange("H17").setNote(Array_Zufallsfragen[3][1]);
      Sheet_Bewerbung.getRange("N17").setNote(Array_Zufallsfragen[4][1]);


      Sheet_Bewerbungen_Aktuell.getRange("B5:O35").sort(4);
    }
  }
}

function Zufallsfragen_Bewerbung()
{
  var Sheet_Fragen = SpreadsheetApp.getActive().getSheetByName("Fragen");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Zeile_Leicht = Sheet_Fragen.getRange("P3").getValue();
  var Zeile_Mittel = Sheet_Fragen.getRange("S3").getValue();
  var Zeile_Schwer = Sheet_Fragen.getRange("V3").getValue();

  var Array_Leicht = Sheet_Fragen.getRange("P6:Q" + Zeile_Leicht).getValues();
  var Array_Mittel = Sheet_Fragen.getRange("S6:T" + Zeile_Mittel).getValues();
  var Array_Schwer = Sheet_Fragen.getRange("V6:W" + Zeile_Schwer).getValues();

  var Anzahl_Leicht = Sheet_Auswertung.getRange("C22").getValue();
  var Anzahl_Mittel = Sheet_Auswertung.getRange("C23").getValue();
  var Anzahl_Schwer = Sheet_Auswertung.getRange("C24").getValue();

  var Array_Var = [["Array_Leicht","Anzahl_Leicht"],["Array_Mittel","Anzahl_Mittel"],["Array_Schwer","Anzahl_Schwer"]];

  var Array_Zufallszahlen = [];
  var Gefunden = false;
  for(var z = 0; z < Array_Var.length; z++)
  {
    Array_Zufallszahlen[z] = [];
    Gefunden = false;

    for(var y = 1; y <= eval(Array_Var[z][1]);)
    {
      Gefunden = false;
      var Zufallszahl = Math.floor(Math.random() * ((eval(Array_Var[z][0]).length) - 0 ) + 0);
      
      for(var x = 0; x < Array_Zufallszahlen[z].length; x++)
      {
        if(Array_Zufallszahlen[z][x] == Zufallszahl)
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == false)
      {
        Array_Zufallszahlen[z][Array_Zufallszahlen[z].length] = Zufallszahl;
        y++
      }
    }
  }

  Logger.log(Array_Zufallszahlen);

  var Array_Ausgabe = [];
  var x = 0;

  for(var z = 0; z < Array_Var.length; z++)
  {
    for(var y = 0; y < Array_Zufallszahlen[z].length; y++)
    {
      Array_Ausgabe[x] = new Array();
      Array_Ausgabe[x][0] = eval(Array_Var[z][0])[Array_Zufallszahlen[z][y]][0];
      Array_Ausgabe[x][1] = eval(Array_Var[z][0])[Array_Zufallszahlen[z][y]][1];
      x++;
    }
  }

  Logger.log(Array_Ausgabe)
  return Array_Ausgabe;
}
