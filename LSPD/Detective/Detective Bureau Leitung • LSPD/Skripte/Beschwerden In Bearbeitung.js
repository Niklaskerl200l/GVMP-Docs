function Beschwerden_Berabeitung(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("I") && Zeile >= 3 && Zeile <= 50)
  {
    Sanktionstyp(e);
  }
  if(Spalte == Spalte_in_Index("AA") && Zeile >= 3 && Zeile <= 50 && Value == "TRUE")
  {
    Archivieren_Bearbeitung(e);
  }
}

function Sanktionstyp(e)
{
  var Zeile = e.range.getRow();
  var Value = e.value;

  var Sheet_Beschwerden_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden In Bearbeitung");

  var Array_Strafkatalog = SpreadsheetApp.getActive().getSheetByName("Bußgeldkatalog").getRange("H16:K29").getValues();
  var Array_Strafe = [];

  var Rang = Sheet_Beschwerden_Bearbeitung.getRange("E" + Zeile).getValue();

  for(var i = 0; i < Array_Strafkatalog.length; i++)
  {
    if(Array_Strafkatalog[i][0] == Rang)
    {
      Array_Strafe = Array_Strafkatalog[i];
      break;
    }
  }

  switch(Value)
  {
    case "Geldstrafe 1" : Sheet_Beschwerden_Bearbeitung.getRange("J" + Zeile).setValue(Array_Strafe[1]); break;
    case "Geldstrafe 2" : Sheet_Beschwerden_Bearbeitung.getRange("J" + Zeile).setValue(Array_Strafe[2]); break;
    case "Geldstrafe 3" : Sheet_Beschwerden_Bearbeitung.getRange("J" + Zeile).setValue(Array_Strafe[3]); break;
    case "Degradierung" : Sheet_Beschwerden_Bearbeitung.getRange("J" + Zeile).setValue("Von Rang " + Rang + " auf Rang " + (Rang - 1)); break;
    case "Supendierung" : Sheet_Beschwerden_Bearbeitung.getRange("J" + Zeile).setValue("Suspendierung für 7 Tage"); break;
  }
}

function Archivieren_Bearbeitung(e)
{
  var Zeile = e.range.getRow();

  var Sheet_Beschwerden_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden In Bearbeitung");
  var Sheet_Beschwerden_Abgeschlossen = SpreadsheetApp.getActive().getSheetByName("Beschwerden Abgeschlossen");

  var Array_Bearbeitung = Sheet_Beschwerden_Bearbeitung.getRange("B" + Zeile + ":Z" + Zeile).getValues();
  var Array_Abgeschlossen = [];

  Sheet_Beschwerden_Bearbeitung.getRange("B" + Zeile + ":Z" + Zeile).setValue("");
  Sheet_Beschwerden_Bearbeitung.getRange("AA" + Zeile).removeCheckboxes();

  Array_Abgeschlossen = 
  [
    new Date(),
    Array_Bearbeitung[0][0],
    Array_Bearbeitung[0][2],
    Array_Bearbeitung[0][3],
    Array_Bearbeitung[0][4],
    Array_Bearbeitung[0][5],
    Array_Bearbeitung[0][6],
    Array_Bearbeitung[0][7],
    Array_Bearbeitung[0][8],
    Array_Bearbeitung[0][9],
    Array_Bearbeitung[0][10],
    Array_Bearbeitung[0][11],
    Array_Bearbeitung[0][12],
    Array_Bearbeitung[0][13],
    Array_Bearbeitung[0][14],
    Array_Bearbeitung[0][15],
    Array_Bearbeitung[0][16],
    Array_Bearbeitung[0][17],
    Array_Bearbeitung[0][18],
    Array_Bearbeitung[0][19],
    Array_Bearbeitung[0][20],
    Array_Bearbeitung[0][21],
    Array_Bearbeitung[0][22],
    Array_Bearbeitung[0][23],
    Array_Bearbeitung[0][24]
  ];

  Sheet_Beschwerden_Abgeschlossen.insertRowAfter(Sheet_Beschwerden_Abgeschlossen.getLastRow());
  var Zeile_Abgeschlossen = Sheet_Beschwerden_Abgeschlossen.getLastRow() + 1;

  Sheet_Beschwerden_Abgeschlossen.getRange("B" + Zeile_Abgeschlossen + ":Z" + Zeile_Abgeschlossen).setValues([Array_Abgeschlossen]);

  Sheet_Beschwerden_Bearbeitung.getRange("B3:AA" + Sheet_Beschwerden_Bearbeitung.getLastRow()).sort(12);

  Sheet_Beschwerden_Abgeschlossen.getRange("B4:Z").sort({column: 2, ascending: false});

  Sheet_Beschwerden_Abgeschlossen.setActiveSelection("D4");

  if(Array_Abgeschlossen[7] == "Geldstrafe 1" || Array_Abgeschlossen[7] == "Geldstrafe 2" || Array_Abgeschlossen[7] == "Geldstrafe 3")
  {
    SpreadsheetApp.getUi().alert("Denk daran die Geldstrafe im Detective Blatt bei Geldverwaltung einzutragen!");
  }
}