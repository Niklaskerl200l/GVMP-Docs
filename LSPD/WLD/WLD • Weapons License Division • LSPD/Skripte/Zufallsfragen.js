function Zufallsfragen(e)
{
  var Sheet_Zufallsfragen = SpreadsheetApp.getActive().getSheetByName("Kursablauf");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("J") && (Zeile == 38 || Zeile == 43 || Zeile == 48) && Value == "TRUE")
  {
    Sheet_Zufallsfragen.getRange(Zeile, Spalte).setValue(false);
    Sheet_Zufallsfragen.getRange(Zeile + 1, Spalte - 4, 3, 5).setValues(Zufallsfragen_Generieren());
  }
}

function Zufallsfragen_Generieren()
{
  var Sheet_Fragebogen = SpreadsheetApp.getActive().getSheetByName("Fragenkatalog");
  var Array_Fragebogen = Sheet_Fragebogen.getRange("B3:C32").getValues();

  Array_Fragebogen = Array_Fragebogen.filter(function(e){return e[0] != ""});

  var Array_Ausgabe = [];
  var Maximal_Fragen = 3;

  for(var i = 0; i < Maximal_Fragen; i++)
  {
    var Zufallsfrage = Math.floor(Math.random() * (Array_Fragebogen.length)) + 0;

    var Gefunden = false;

    for(var o = 0; o < Array_Ausgabe.length; o++)
    {
      if(Array_Ausgabe[o][0] == Array_Fragebogen[Zufallsfrage][0])
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == true)
    {
      i--;
    }
    else
    {
      Array_Ausgabe.push([Array_Fragebogen[Zufallsfrage][0].toString().trim(),"","","", Array_Fragebogen[Zufallsfrage][1].toString().trim()]);
    }
  }

  Logger.log(Array_Ausgabe);
  return Array_Ausgabe;
}