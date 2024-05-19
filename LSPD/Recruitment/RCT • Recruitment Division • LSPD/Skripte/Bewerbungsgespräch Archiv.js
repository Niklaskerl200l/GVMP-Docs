function Bewerbung_Angenommen()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_Bewerbung = SpreadsheetApp.getActiveSheet();
  var Sheet_Bewerbungen_Archiv = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Archiv");
  var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Bewerbungsgespräche);

  var Array_Bewerbung = Sheet_Bewerbung.getRange("B5:V5").getValues();
  var Email = Sheet_Bewerbung.getRange("T11").getValue();
  var Zeile_Export = Sheet_Export.getSheetByName("Übersicht").getRange("B1").getValue();

  Academy_Start();

  var Array_Archiv = 
  [[
    Array_Bewerbung[0][0],
    "Angenommen",
    Array_Bewerbung[0][2],
    Array_Bewerbung[0][4],
    Array_Bewerbung[0][8],
    Array_Bewerbung[0][10],
    Array_Bewerbung[0][11],
    Array_Bewerbung[0][12],
    Array_Bewerbung[0][13],
    Email,
    Array_Bewerbung[0][14],
    Array_Bewerbung[0][16],
    Array_Bewerbung[0][19],
    Array_Bewerbung[0][20]
  ]];

  var Array_Export = [[]];
  Array_Export[0][0] = Array_Bewerbung[0][0];
  Array_Export[0][1] = Array_Bewerbung[0][2];
  Array_Export[0][2] = Array_Bewerbung[0][4];
  Array_Export[0][3] = Sheet_Bewerbung.getRange("T8").getValue();
  Array_Export[0][4] = Sheet_Bewerbung.getRange("U8").getValue();
  Array_Export[0][5] = Sheet_Bewerbung.getRange("V8").getValue();
  Array_Export[0][6] = "Angenommen";

  var Sheet_Neu = Sheet_Bewerbung.copyTo(Sheet_Export).setName("Bewerbungsgespräch " + Array_Bewerbung[0][0] + " " + Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM HH:mm"));

  try
  {
    Sheet_Neu.getRange(1,1,Sheet_Bewerbung.getLastRow(),Sheet_Bewerbung.getLastColumn()).setValues(Sheet_Bewerbung.getRange(1,1,Sheet_Bewerbung.getLastRow(),Sheet_Bewerbung.getLastColumn()).getValues());
  }
  catch(err)
  {

  }

  var URL = Sheet_Export.getUrl() + "#gid=" + Sheet_Export.getSheetByName(Sheet_Neu.getName()).getSheetId();

  Sheet_Export = Sheet_Export.getSheetByName("Übersicht");

  Sheet_Export.getRange("B" + Zeile_Export + ":H" + Zeile_Export).setValues(Array_Export);
  Sheet_Export.getRange("I" + Zeile_Export).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");

  var Zeile_Archiv = 6

  Sheet_Bewerbungen_Archiv.insertRowAfter(5);

  Sheet_Bewerbungen_Archiv.getRange("B" + Zeile_Archiv + ":O" + Zeile_Archiv).setValues(Array_Archiv);
  Sheet_Bewerbungen_Archiv.getRange("P" + Zeile_Archiv).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");

  var Forum_ID = SpreadsheetApp.newRichTextValue().setText(Array_Bewerbung[0][13]).setLinkUrl("https://www.gvmp.de/index.php?user/" + Array_Bewerbung[0][13]).build();

  Sheet_Bewerbungen_Archiv.getRange("J" + Zeile_Archiv).setRichTextValue(Forum_ID);
}

function Bewerbung_Abgelehnt()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_Bewerbung = SpreadsheetApp.getActiveSheet();
  var Sheet_Bewerbungen_Archiv = SpreadsheetApp.getActive().getSheetByName("Bewerbungen Archiv");
  var Sheet_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Bewerbungsgespräche);

  var Array_Bewerbung = Sheet_Bewerbung.getRange("B5:V5").getValues();
  var Email = Sheet_Bewerbung.getRange("T11").getValue();
  var Zeile_Export = Sheet_Export.getSheetByName("Übersicht").getRange("B1").getValue();

  var Array_Archiv = 
  [[
    Array_Bewerbung[0][0],
    "Abgelehnt",
    Array_Bewerbung[0][2],
    Array_Bewerbung[0][4],
    Array_Bewerbung[0][8],
    Array_Bewerbung[0][10],
    Array_Bewerbung[0][11],
    Array_Bewerbung[0][12],
    Array_Bewerbung[0][13],
    Email,
    Array_Bewerbung[0][14],
    Array_Bewerbung[0][16],
    Array_Bewerbung[0][19],
    Array_Bewerbung[0][20]
  ]];

  var Array_Export = [[]];
  Array_Export[0][0] = Array_Bewerbung[0][0];
  Array_Export[0][1] = Array_Bewerbung[0][2];
  Array_Export[0][2] = Array_Bewerbung[0][4];
  Array_Export[0][3] = Sheet_Bewerbung.getRange("T8").getValue();
  Array_Export[0][4] = Sheet_Bewerbung.getRange("U8").getValue();
  Array_Export[0][5] = Sheet_Bewerbung.getRange("V8").getValue();
  Array_Export[0][6] = "Abgelehnt";

  var Sheet_Name;
  var Sheet_Neu = Sheet_Bewerbung.copyTo(Sheet_Export);

  try
  {
    Sheet_Neu.setName(Sheet_Bewerbung.getName());
    Sheet_Name = Sheet_Bewerbung.getName();
  }
  catch(err)
  {
    Sheet_Neu.setName(Sheet_Bewerbung.getName() + Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM HH:mm"));
    Sheet_Name = Sheet_Bewerbung.getName() + Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM HH:mm")
  }

  var URL = Sheet_Export.getUrl() + "#gid=" + Sheet_Export.getSheetByName(Sheet_Name).getSheetId();

  Sheet_Export = Sheet_Export.getSheetByName("Übersicht");

  Sheet_Export.getRange("B" + Zeile_Export + ":H" + Zeile_Export).setValues(Array_Export);
  Sheet_Export.getRange("I" + Zeile_Export).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");

  var Zeile_Archiv = 6

  Sheet_Bewerbungen_Archiv.insertRowAfter(5);

  Sheet_Bewerbungen_Archiv.getRange("B" + Zeile_Archiv + ":O" + Zeile_Archiv).setValues(Array_Archiv);
  Sheet_Bewerbungen_Archiv.getRange("P" + Zeile_Archiv).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")");

  SS.deleteSheet(Sheet_Bewerbung);

  Sheet_Bewerbungen_Archiv.setActiveSelection("B6");
}
