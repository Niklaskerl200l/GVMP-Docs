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

LSPD.Eingabe_Test();

String.prototype.replaceAt = function(index, replacement) 
{
    return this.substr(0, index) + replacement + this.substr(index + replacement.length);
}

function Eintrag_Check(Array_Eingabe, Off_Name, Off_Fraktion, Off_Aktivitaet, Off_Beamter, UI = false)
{
  Logger.log("Start Eintrag Check für:\n" + Array_Eingabe);

  if(Off_Name != -1) var Name = Array_Eingabe[0][Off_Name].toString();

  if(Off_Fraktion != -1) var Fraktion = Array_Eingabe[0][Off_Fraktion];

  if(Off_Aktivitaet != -1) var Aktivitaet = Array_Eingabe[0][Off_Aktivitaet];
  
  if(Off_Beamter != -1) var Beamter = Array_Eingabe[0][Off_Beamter];
  
  
  if(Off_Name != -1)
  {
    if(Name == null || Name == "" || Name == undefined)
    {
      Logger.log("Name ist Leer");
      if(UI) SpreadsheetApp.getUi().alert("Bitte gib einen Namen ein!");
      return 1;
    }
  }
  
  if(Off_Fraktion != -1)
  {
    if(Fraktion == null || Fraktion == "" || Fraktion == undefined)
    {
      Logger.log("Fraktion ist Leer");
      if(UI) SpreadsheetApp.getUi().alert("Bitte gib eine Fraktion ein!");
      return 1;
    }
  }

  if(Off_Aktivitaet != -1)
  {
    if(Aktivitaet == null || Aktivitaet == "" || Aktivitaet == undefined)
    {
      Logger.log("Aktivität ist Leer");
      if(UI) SpreadsheetApp.getUi().alert("Bitte gib eine Aktivität ein!");
      return 1;
    }
  }

  if(Off_Beamter != -1)
  {
    if(Beamter == null || Beamter == "" || Beamter == undefined)
    {
      Logger.log("Beamter ist Leer");
      if(UI) SpreadsheetApp.getUi().alert("Bitte gib einen Beamter ein!");
      return 1;
    }
  }

  if(Off_Name != -1)
  {
    Name = Name.toString();

    Name= Name.trim();
    
    if(Name.includes("_") == false)
    {
      Logger.log("Name hat keinen Unterstrich " + Name);
      if(UI) SpreadsheetApp.getUi().alert("Bitte gib den Namen mit _ ein!");
      return 1;
    }

    if(Name.includes(" ") == true)
    {
      Logger.log("Name mit Leerzeichen");
      if(UI) SpreadsheetApp.getUi().alert("Bitte gib den Zweit Namen mit - ein und nicht mit Leerzeichen!");
      return 1;
    }

    for(var i = 0; i < Name.length; i++)
    {
      var Ascii = Name.charCodeAt(i);
      
      if(!(Ascii >= 65 && Ascii <= 90 || Ascii >= 97 && Ascii <= 122 || Ascii == 45 || Ascii == 95))
      {
        Name = Name.substr(0,i) + Name.substr(i+1);
        i--;
      }
    }

    Name = Name.replaceAt(0,Name[0].toUpperCase());
    Name = Name.replaceAt(Name.indexOf("_") + 1,Name[Name.indexOf("_") + 1].toUpperCase());

    if(Name.includes("-") == true)
    {
      Name = Name.replaceAt(Name.indexOf("-") + 1,Name[Name.indexOf("-") + 1].toUpperCase());
    }

    Array_Eingabe[0][Off_Name] = Name;
  }

  Logger.log("Check bestanden Ausgabe:\n" + Array_Eingabe);

  return Array_Eingabe;
}