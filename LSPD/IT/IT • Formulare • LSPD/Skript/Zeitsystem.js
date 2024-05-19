var ID_Zeitsystem = LSPD.ID_Zeitsystem;

function Zeitsystem_Einstellung(e)
{
  var Sheet_Zeitsystem = SpreadsheetApp.openById(ID_Zeitsystem).getSheetByName("Zeitsystem");

  var Werte = e.namedValues;
  
  var Array_Eingabe = [[Werte.Name,"","","","0:00","","0:00","0:00","0:00","","","0:00","0:00","0:00","0:00","",new Date(),"","0:00","0:00","0:00","0:00","0:00",""]];
  
  var Letzte_Zeile = Sheet_Zeitsystem.getLastRow() + 1;

  Sheet_Zeitsystem.insertRowAfter(Letzte_Zeile - 1);

  Sheet_Zeitsystem.getRange("B" + Letzte_Zeile + ":Y" + Letzte_Zeile).setValues(Array_Eingabe);

  Sheet_Zeitsystem.getRange("C" + Letzte_Zeile).setFormula("=IF($B" + Letzte_Zeile +"=\"\";\"\";VLOOKUP($B" + Letzte_Zeile +";'Import Personaltabelle'!$A$4:$H;2;FALSE))");
  Sheet_Zeitsystem.getRange("D" + Letzte_Zeile).setFormula("=IF($B" + Letzte_Zeile +"=\"\";\"\";VLOOKUP($B" + Letzte_Zeile +";'Import Personaltabelle'!$A$4:$H;8;FALSE))");
  Sheet_Zeitsystem.getRange("G" + Letzte_Zeile).setFormula("=SUM(\"0:00\";F"+Letzte_Zeile+")");
  Sheet_Zeitsystem.getRange("L" + Letzte_Zeile).setFormula("=SUM(\"0:00\";F"+Letzte_Zeile+")");
  Sheet_Zeitsystem.getRange("Q" + Letzte_Zeile).setFormula("=SUM(L"+Letzte_Zeile+":P"+Letzte_Zeile+")");
  Sheet_Zeitsystem.getRange("Y" + Letzte_Zeile).setFormula("=SUM(T"+Letzte_Zeile+":X"+Letzte_Zeile+")");
}

function Zeitsystem_Entlassung(e)
{
  var Sheet_Zeitsystem = SpreadsheetApp.openById(ID_Zeitsystem).getSheetByName("Entlassungen");

  var Werte = e.namedValues;

  Sheet_Zeitsystem.appendRow(["",Werte.Name[0],""])
}

function Zeitsystem_Einstellung_Name(Name = "Junior Kadaver-Jack")
{
  var Sheet_Zeitsystem = SpreadsheetApp.openById(ID_Zeitsystem).getSheetByName("Zeitsystem");

  
  var Array_Eingabe = [[Name,"","","","0:00","","0:00","0:00","0:00","","","0:00","0:00","0:00","0:00","","","","0:00","0:00","0:00","0:00","0:00",""]];
  
  var Letzte_Zeile = Sheet_Zeitsystem.getLastRow() + 1;

  Sheet_Zeitsystem.insertRowAfter(Letzte_Zeile - 1);

  Sheet_Zeitsystem.getRange("B" + Letzte_Zeile + ":Y" + Letzte_Zeile).setValues(Array_Eingabe);

  Sheet_Zeitsystem.getRange("C" + Letzte_Zeile).setFormula("=IF($B" + Letzte_Zeile +"=\"\";\"\";VLOOKUP($B" + Letzte_Zeile +";'Import Personaltabelle'!$A$4:$H;2;FALSE))");
  Sheet_Zeitsystem.getRange("D" + Letzte_Zeile).setFormula("=IF($B" + Letzte_Zeile +"=\"\";\"\";VLOOKUP($B" + Letzte_Zeile +";'Import Personaltabelle'!$A$4:$H;8;FALSE))");
  Sheet_Zeitsystem.getRange("G" + Letzte_Zeile).setFormula("=SUM(\"0:00\";F"+Letzte_Zeile+")");
  Sheet_Zeitsystem.getRange("L" + Letzte_Zeile).setFormula("=SUM(\"0:00\";F"+Letzte_Zeile+")");
  Sheet_Zeitsystem.getRange("Q" + Letzte_Zeile).setFormula("=SUM(L"+Letzte_Zeile+":P"+Letzte_Zeile+")");
  Sheet_Zeitsystem.getRange("Y" + Letzte_Zeile).setFormula("=SUM(T"+Letzte_Zeile+":X"+Letzte_Zeile+")");
}