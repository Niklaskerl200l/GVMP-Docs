function Design_LSPD_Dark() // Niklas_Kerl®
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = SpreadsheetApp.getActiveSheet();

  var Themes = SpreadsheetApp.ThemeColorType;
  var SSTheme = SS.getPredefinedSpreadsheetThemes()[1];
  var Color = SpreadsheetApp.newColor();

  SSTheme.setFontFamily("Roboto Condensed")
  .setConcreteColor(Themes.TEXT,Color.setRgbColor("#f3f3f3"))
  .setConcreteColor(Themes.BACKGROUND,Color.setRgbColor("#181818"))
  .setConcreteColor(Themes.ACCENT1,Color.setRgbColor("#353535"))
  .setConcreteColor(Themes.ACCENT2,Color.setRgbColor("#565656"))
  .setConcreteColor(Themes.ACCENT3,Color.setRgbColor("#7d7d7d"))
  .setConcreteColor(Themes.ACCENT4,Color.setRgbColor("#1e1e1e"))
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