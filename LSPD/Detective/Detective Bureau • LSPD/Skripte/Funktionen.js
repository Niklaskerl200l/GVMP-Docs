function Spalte_in_Index(Text)    // String in Spaltenindex
{
  Text = Text.toUpperCase();
  
  if(Text.length == 1)
  {
    return Switch_ABC(Text);
  }
  else if(Text.length == 2)
  {
    return (Switch_ABC(Text[0]) * 26) + Switch_ABC(Text[1]);
  }

}

function Switch_ABC(Text)
{
  switch(Text)
    {
      case "A" : return 1; break;
      case "B" : return 2; break;
      case "C" : return 3; break;
      case "D" : return 4; break;
      case "E" : return 5; break;
      case "F" : return 6; break;
      case "G" : return 7; break;
      case "H" : return 8; break;
      case "I" : return 9; break;
      case "J" : return 10; break;
      case "K" : return 11; break;
      case "L" : return 12; break;
      case "M" : return 13; break;
      case "N" : return 14; break;
      case "O" : return 15; break;
      case "P" : return 16; break;
      case "Q" : return 17; break;
      case "R" : return 18; break;
      case "S" : return 19; break;
      case "T" : return 20; break;
      case "U" : return 21; break;
      case "V" : return 22; break;
      case "W" : return 23; break;
      case "X" : return 24; break;
      case "Y" : return 25; break;
      case "Z" : return 26; break;
      default: return 0;
    }
}

function Kalenderwoche(datum = new Date())
{
  var Datum = new Date(datum.valueOf());
  var TagN = (datum.getDay() + 6) % 7;

  Datum.setDate(Datum.getDate() - TagN + 3);

  var Erster_Donnerstag = Datum.valueOf();
  
  Datum.setMonth(0, 1);

  if (Datum.getDay() !== 4) 
  {
    Datum.setMonth(0, 1 + ((4 - Datum.getDay()) + 7) % 7);
  }

  return 1 + Math.ceil((Erster_Donnerstag - Datum) / 604800000);  
}

function Design_LSPD_Dark() // // Niklas_Kerl®
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = SpreadsheetApp.getActiveSheet();

  // Sheet.clear().clearContents().clearFormats().clearNotes().clearConditionalFormatRules();   // Clear Kontent

  var Themes = SpreadsheetApp.ThemeColorType;
  var SSTheme = SS.getPredefinedSpreadsheetThemes()[1];
  var Color = SpreadsheetApp.newColor();

  SSTheme.setFontFamily("Roboto Condensed")
  .setConcreteColor(Themes.TEXT,Color.setRgbColor("#f3f3f3"))
  .setConcreteColor(Themes.BACKGROUND,Color.setRgbColor("#181818"))
  .setConcreteColor(Themes.ACCENT1,Color.setRgbColor("#353535"))
  .setConcreteColor(Themes.ACCENT2,Color.setRgbColor("#565656"))
  .setConcreteColor(Themes.ACCENT3,Color.setRgbColor("#7d7d7d"))
  .setConcreteColor(Themes.ACCENT4,Color.setRgbColor("#434343"))
  .setConcreteColor(Themes.ACCENT5,Color.setRgbColor("#ffffff"))
  .setConcreteColor(Themes.ACCENT6,Color.setRgbColor("#ffffff"))
  .setConcreteColor(Themes.HYPERLINK,Color.setRgbColor("#0000ff"));

  SS.setSpreadsheetTheme(SSTheme);

  var FontColor = Color.setThemeColor(SpreadsheetApp.ThemeColorType.TEXT).build();
  var BackgroundColor = Color.setThemeColor(SpreadsheetApp.ThemeColorType.ACCENT4).build();

  var Range_Sheet = Sheet.getRange(1,1,Sheet.getMaxRows(),Sheet.getMaxColumns());

  Range_Sheet.setFontColorObject(FontColor);
  Range_Sheet.setBackgroundObject(BackgroundColor);
  Range_Sheet.setHorizontalAlignment("center");
  Range_Sheet.setVerticalAlignment("middle");
}

function Zeit_Dauer(Zeit)
{
  Zeit = new Date(Zeit);
  var Stunden;
  var Minuten;
  var Text;

  var Dummy = Zeit;
  Dummy.setDate(Dummy.getDate() + 25569)
  Dummy = ((Dummy.getTime() / 1000 / 60 / 60 + 1).toFixed(2)).toString();

  if(Dummy.length == 4)
  {
    Stunden = Dummy[0];
  }
  else if(Dummy.length == 5)
  {
    Stunden = Dummy[0] + Dummy[1];
  }
  else if(Dummy.length == 6)
  {
    Stunden = Dummy[0] + Dummy[1] + Dummy[2];
  }
  else if(Dummy.length == 7)
  {
    Stunden = Dummy[0] + Dummy[1] + Dummy[2] + Dummy[3];
  }

  Minuten = Zeit.getMinutes().toString();

  if(Minuten.length == 1)
  {
    Text = Stunden + ":0" + Minuten;
  }
  else
  {
    Text = Stunden + ":" + Minuten;
  }

  return [Stunden,Minuten,Text];
 
}

LSPD.Eingabe_Test();

function Sort_Beschwerden()
{
  var Sheet_Beschwerden_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
  var Sheet_Beschwerden_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden In Bearbeitung");
  var Sheet_Abgeschlossen = SpreadsheetApp.getActive().getSheetByName("Beschwerden Abgeschlossen");

  Sheet_Beschwerden_Neu.getRange("B3:Q" + Sheet_Beschwerden_Neu.getLastRow()).sort(Spalte_in_Index("F"));
  Sheet_Abgeschlossen.getRange("B3:AD" + Sheet_Abgeschlossen.getLastRow()).sort([{column: Spalte_in_Index("B"), ascending: false}]);

  Sheet_Beschwerden_Bearbeitung.getRange("B3:AD" + Sheet_Beschwerden_Bearbeitung.getLastRow()).sort([{column: Spalte_in_Index("C"), ascending: true}, {column: Spalte_in_Index("AD"), ascending: true}, {column:Spalte_in_Index("B"),ascending:true}]);
}

function Suche_Beschwerde()
{
  var UI = SpreadsheetApp.getUi();
  var Beschwerde_ID = UI.prompt("Fallnummer", "Geben Sie die Fallnummer an!", UI.ButtonSet.OK).getResponseText();

  if(Beschwerde_ID == undefined || Beschwerde_ID == "" || Beschwerde_ID == null)
  {
    return;
  }

  var Gefunden = false;

  var Sheet_Beschwerde_Neu = SpreadsheetApp.getActive().getSheetByName("Beschwerden Neu");
  var Array_Beschwerde_Neu = Sheet_Beschwerde_Neu.getRange("B3:B36").getValues();

  for(var i = 0; i < Array_Beschwerde_Neu.length; i++)
  {
    if(Array_Beschwerde_Neu[i][0] != "")
    {
      if(Array_Beschwerde_Neu[i][0].toString().includes(Beschwerde_ID.toString()) == true)
      {
        Gefunden = true;
        break;
      }
    }
  }

  if(Gefunden == true)
  {
    return Sheet_Beschwerde_Neu.setActiveSelection("B" + (i + 3));
  }

  var Sheet_Beschwerden_Bearbeitung = SpreadsheetApp.getActive().getSheetByName("Beschwerden In Bearbeitung");
  var Array_Beschwerden_Bearbeitung = Sheet_Beschwerden_Bearbeitung.getRange("B3:B34").getValues();

  for(var i = 0; i < Array_Beschwerden_Bearbeitung.length; i++)
  {
    if(Array_Beschwerden_Bearbeitung[i][0] != "")
    {
      if(Array_Beschwerden_Bearbeitung[i][0].toString().includes(Beschwerde_ID.toString()) == true)
      {
        Gefunden = true;
        break;
      }
    }
  }

  if(Gefunden == true)
  {
    return Sheet_Beschwerden_Bearbeitung.setActiveSelection("B" + (i + 3));
  }
}

function Suche_DN()
{
  var UI = SpreadsheetApp.getUi();
  var DN = UI.prompt("Dienstnummer", "Geben Sie die Dienstnummer an!", UI.ButtonSet.OK).getResponseText();

  if(DN == undefined || DN == "" || DN == null)
  {
    return;
  }

  var Gefunden = false;

  var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Import Personaltabelle");
  var Array_Personal = Sheet_Personal.getRange("C4:E199").getValues();

  for(var i = 0; i < Array_Personal.length; i++)
  {
    if(Array_Personal[i][1] != "")
    {
      if(Array_Personal[i][2].toString() == DN.toString())
      {
        Gefunden = true;
        break;
      }
    }
  }

  if(Gefunden == true)
  {
    UI.alert("Mitarbeiter: " + Array_Personal[i][1] + "\nDienstgrad: " + Array_Personal[i][0]);
  }
  else
  {
    UI.alert("Mitarbeiter nicht gefunden!");
  }
}

function DB_Eintragen(Name = "Darren Wales", State = 1)
{
  SpreadsheetApp.openById(LSPD.ID_Dienstblatt).getSheetByName("Log Stempeluhr").appendRow(["", Name, State, new Date, ""]);
}