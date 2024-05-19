function Sondergenehmigungen(e)
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Sondergenehmigungen");
  var Spalte = e.range.getColumn();
  var Zeile = e.range.getRow();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile == 5 && Value != undefined)
  {
    var ArrayAusgabe;
    var Datum = new Date();
    Datum.setDate(Datum.getDate() + 1);

    ArrayAusgabe = [[new Date() , Datum , LSPD.Umwandeln()]];
    Sheet.getRange(`D5:F5`).setValues(ArrayAusgabe);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile == 5 && Value == undefined)
  {
    Sheet.getRange(`B5:F5`).clearContent();
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile == 5 && Value == "TRUE")
  {
    Sheet.getRange(Zeile , Spalte).setValue(false);

    var Array_Eingabe = Sheet.getRange(`B5:F5`).getValues();
    var Array_AktiveGenehmigungen = Sheet.getRange(`B12:F${Sheet.getRange(`B8`).getValue()}`).getValues();
    var Next_Archiv_Row = 12;

    var UI = SpreadsheetApp.getUi();

    if(Array_Eingabe[0][0] == "")
    {
      return UI.alert("Fehler!" , "Sie müssen einen Namen angeben!" , UI.ButtonSet.OK);
    }

    for(var i = 0; i < Array_AktiveGenehmigungen.length; i++)
    {
      if(Array_AktiveGenehmigungen[i][0] == Array_Eingabe[0][0])
      {
        var Datum_VON = new Date(Array_AktiveGenehmigungen[i][2]);
        var Datum_BIS = new Date(Array_AktiveGenehmigungen[i][3]);

        if(Datum_VON < new Date() && Datum_BIS > new Date() && Array_AktiveGenehmigungen[i][1] == Array_Eingabe[0][1])
        {
          return UI.alert("Fehler!" , "Dieser Beamte hat bereits eine laufende Freistellung!" , UI.ButtonSet.OK);
        }
      }
    }

    Sheet.insertRowAfter(Next_Archiv_Row - 1);
    Sheet.getRange(`B12:F12`).setValues(Array_Eingabe);
    Sheet.getRange(`B5:F5`).clearContent();
  }
  else if(Spalte == Spalte_in_Index("G") && Zeile >= 12 && Value == "TRUE")
  {
    var UI = SpreadsheetApp.getUi();

    if(UI.alert("Sondergenehmigung löschen." , `Möchten Sie die Sondergenehmigung von ${Sheet.getRange(Zeile , Spalte_in_Index("B")).getValue()} löschen?` , UI.ButtonSet.YES_NO) == UI.Button.YES)
    {
      Sheet.deleteRow(Zeile);
    }
  }
  else if(Spalte == Spalte_in_Index("O") && Zeile >= 5 && Zeile <= 10 && Value == "TRUE")
  {
    var UI = SpreadsheetApp.getUi();
    if(UI.alert("Sondergenehmigung löschen." , `Möchten Sie die Sondergenehmigung von ${Sheet.getRange(Zeile , Spalte_in_Index("I")).getValue()} löschen?` , UI.ButtonSet.YES_NO) != UI.Button.YES) return;

    var Array_Uebersicht = Sheet.getRange(`I${Zeile}:N${Zeile}`).getValues();
    var Array_AktiveGenehmigungen = Sheet.getRange(`B12:F${Sheet.getRange(`B8`).getValue()}`).getValues();

    var Gefunden = false;

    for(var i = 0; i < Array_AktiveGenehmigungen.length; i++)
    {
      if(Array_AktiveGenehmigungen[i][0] == Array_Uebersicht[i][0] && Array_AktiveGenehmigungen[i][1] == Array_Uebersicht[0][2])
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden)
    {
      Sheet.deleteRow(i + 12);
    }

    Sheet.getRange(Zeile , Spalte).setValue(false);
  }
}