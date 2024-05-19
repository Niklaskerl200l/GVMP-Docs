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

function Freigabe(Freigabe_ID,Array_Email_1D = false,Array_Email_2D = false)
{
  if(Array_Email_1D == false)
  {
    var Array_Email = new Array();

    for(var i = 0; i < Array_Email_2D.length; i++)
    {
      if(Array_Email_2D[i][0] != "")
      {
        Array_Email.push(Array_Email_2D[i][0]);
      }
    }
  }
  else if(Array_Email_2D == false)
  {
    var Array_Email = new Array();

    for(var i = 0; i < Array_Email_1D.length; i++)
    {
      if(Array_Email_1D[i] != "")
      {
        Array_Email.push(Array_Email_1D[i]);
      }
    }
  }

  var Sheet_Freigabe = SpreadsheetApp.openById(Freigabe_ID);
  var Array_Editor = Sheet_Freigabe.getEditors();

  for(var i = 0; i < Array_Editor.length; i++)    // Entferne Zugriff
  {
    var Gefunden = false;

    for(var y = 0; y < Array_Email.length; y++)
    {
      if(Array_Editor[i].getEmail().toUpperCase() == Array_Email[y].toString().toUpperCase())
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      try
      {
        Logger.log("Remove: " + Array_Editor[i]);

        Sheet_Freigabe.removeEditor(Array_Editor[i]);
      }
      catch(err)
      {
        Logger.log(err.stack);
      }
    }
  }

  for(var i = 0; i < Array_Email.length; i++)    // Add Zugriff
  {
    var Gefunden = false;

    for(var y = 0; y < Array_Editor.length; y++)
    {
      if(Array_Editor[y].getEmail().toUpperCase() == Array_Email[i].toString().toUpperCase())
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false && Array_Email[i] != "")
    {
      try
      {
        Logger.log("Add: " + Array_Email[i]);

        Sheet_Freigabe.addEditor(Array_Email[i]);
      }
      catch(err)
      {
        Logger.log(err.stack);
      }
    }
  }
}