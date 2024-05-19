function LST_Suche(e)
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("LST Suche");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("E") && Zeile == 5 && Value == "TRUE")
  {
    Sheet.getRange(Zeile, Spalte).setValue(false);
    if(Leer("B5:D5", 3) == false) return Suche_LST_Funktion();
    Suche_LST_Funktion(
      Sheet.getRange(Zeile, Spalte_in_Index("B")).getValue(),
      Sheet.getRange(Zeile, Spalte_in_Index("C")).getValue(),
      Sheet.getRange(Zeile, Spalte_in_Index("D")).getValue(),
    );
  }
}

function Leer(Bereich, Anzahl_Spalten)
{
  var Daten = SpreadsheetApp.getActive().getSheetByName("LST Suche").getRange(Bereich).getValues();
  var Wert = false;
  for(var i = 0, ian = Anzahl_Spalten; i < ian; i++)
  {
    if(Daten[0][i] != "")
    {
      Wert = true;
      break;
    }
  }

  return Wert;
}

function Suche_LST_Funktion(Datum, Erwählter, Ausführender)
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("LST Suche");
  var Dienstblatt = SpreadsheetApp.openById(LSPD.ID_Archiv_LSTSuche).getSheetByName("Archiv");

  Sheet.getRange("B8:E506").clearContent();
  if(Datum == undefined && Erwählter == undefined && Ausführender == undefined) return;

  var DB_LetzteZeile = Dienstblatt.getRange("B1").getValue();
  var Daten = Dienstblatt.getRange("B3:D" + DB_LetzteZeile).getValues();

  var Datum_Morgen = new Date(Datum).setDate(new Date(Datum).getDate() + 1);

  if(Datum != "")       Daten = Daten.filter(function(e){return e[0] != "" && e[0] > Datum && e[0] < Datum_Morgen});
  if(Erwählter != "")   Daten = Daten.filter(function(e){return e[0] != "" && e[1] == Erwählter});
  if(Ausführender != "")Daten = Daten.filter(function(e){return e[0] != "" && e[2] == Ausführender});

  var Ausgabe_Array = new Array();
  var MaximaleAnzahlAnAusgaben = 500;

  for(var i = 0; i < Daten.length; i++)
  {
    if(i <= MaximaleAnzahlAnAusgaben)
    {
      Ausgabe_Array.push([Daten[i][0], Daten[i][1], Daten[i][2], Daten[i][0]]);
    }
  }


  if(Ausgabe_Array.length < 1) return Sheet.getRange("C8").setValue("Kein Fund.");
  Sheet.getRange("B8:E" + Math.floor(Ausgabe_Array.length + 7)).setValues(Ausgabe_Array);
}