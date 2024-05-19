function Beschwerden_Neu(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("C") && Zeile >= 3 && Zeile <= 50 && Value != undefined && OldValue == undefined)
  {
    Eintragen_Neu(e);
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile >= 3 && Zeile <= 50 && Value == undefined )
  {
    Austragen_Neu(e);
  }
  else if(Spalte == Spalte_in_Index("Q") && Zeile >= 3 && Zeile <= 50 && Value == "TRUE" )
  {
    Bearbeiten_Neu(e);
  }
}

function Eintragen_Neu(e)
{
  var Zeile = e.range.getRow();

  var Sheet_Beschwerden_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var Fallnummer = Sheet_Auswertung.getRange("C3").getValue();

  Sheet_Beschwerden_Neu.getRange("B" + Zeile).setValue(Fallnummer);
  Sheet_Beschwerden_Neu.getRange("E" + Zeile).setValue("Offen");
  Sheet_Beschwerden_Neu.getRange("F" + Zeile).setValue(new Date());
  Sheet_Beschwerden_Neu.getRange("P" + Zeile).setValue(LSPD.Umwandeln());
}

function Austragen_Neu(e)
{
  var Zeile = e.range.getRow();

  var Sheet_Beschwerden_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");

  Sheet_Beschwerden_Neu.getRange("B" + Zeile).setValue("");
  Sheet_Beschwerden_Neu.getRange("E" + Zeile).setValue("");
  Sheet_Beschwerden_Neu.getRange("F" + Zeile).setValue("");
  Sheet_Beschwerden_Neu.getRange("P" + Zeile).setValue("");
}

function Bearbeiten_Neu(e)
{
  var Zeile = e.range.getRow();

  var Sheet_Beschwerden_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
  var Sheet_Beschwerden_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden In Bearbeitung");

  var Array_Neu = Sheet_Beschwerden_Neu.getRange("B" + Zeile + ":P" + Zeile).getValues();
  var Array_Bearbeitung = [];

  Sheet_Beschwerden_Neu.getRange("Q" + Zeile).setValue(false);
  Sheet_Beschwerden_Neu.getRange("B" + Zeile + ":P" + Zeile).setValue("")
  Sheet_Beschwerden_Neu.getRange("D" + Zeile).setFormula("=IF(C"+Zeile+"=\"\";\"\";VLOOKUP(C"+Zeile+";'Import Personaltabelle'!$A$4:$B;2;FALSE))")

  Array_Bearbeitung = 
  [
    Array_Neu[0][0],
    "",
    Array_Neu[0][1],
    Array_Neu[0][2],
    LSPD.Umwandeln(),
    "In Bearbeitung",
    "In Bearbeitung",
    "",
    "",
    "",
    Array_Neu[0][4],
    Array_Neu[0][5],
    Array_Neu[0][6],
    Array_Neu[0][7],
    Array_Neu[0][8],
    Array_Neu[0][9],
    "",
    "",
    Array_Neu[0][10],
    Array_Neu[0][11],
    Array_Neu[0][12],
    Array_Neu[0][13],
    "",
    "",
    Array_Neu[0][14]
  ];

  var Zeile_Bearbeiten = Sheet_Beschwerden_Bearbeitung.getLastRow() + 1;

  Sheet_Beschwerden_Bearbeitung.getRange("B" + Zeile_Bearbeiten + ":Z" + Zeile_Bearbeiten).setValues([Array_Bearbeitung]);
  Sheet_Beschwerden_Bearbeitung.getRange("AA" + Zeile_Bearbeiten).insertCheckboxes();

  Sheet_Beschwerden_Bearbeitung.setActiveSelection("C" + Zeile_Bearbeiten);

  Sheet_Beschwerden_Neu.getRange("B3:Q" + Sheet_Beschwerden_Neu.getLastRow()).sort(6);
}