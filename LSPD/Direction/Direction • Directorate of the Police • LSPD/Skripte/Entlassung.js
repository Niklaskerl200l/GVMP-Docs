function Entlassung(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Sheet = SpreadsheetApp.getActive().getSheetByName("Personal Master")
  var Werte = Sheet.getRange("B" + Zeile + ":O" + Zeile).getValues();
  var Datum = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd");
  var Name = LSPD.Umwandeln().toString().replace(" ","+");

  var URL = "https://docs.google.com/forms/d/e/1FAIpQLScpx5CaXdCzndSTcqniQyh3I1kTB7nis-vlBEXl4WrD5o9HAw/viewform?usp=pp_url&entry.1362312353="+Werte[0][1].replace(" ","+")+"&entry.1149726796="+Name+"&entry.1351926169=Nein";

  Sheet.getRange(Zeile,Spalte).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")").setFontColor("#0000ff").removeCheckboxes();

  Utilities.sleep(10000);

  Sheet.getRange(Zeile,Spalte).setFontColor("#f3f3f3").insertCheckboxes();
}
