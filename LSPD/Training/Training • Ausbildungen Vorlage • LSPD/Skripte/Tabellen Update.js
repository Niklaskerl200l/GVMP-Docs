function Tabellen_Update()
{
  var Tabellen_Blatt = "Qualifizierungen";
  var Sheet_Vorlage = SpreadsheetApp.getActive().getSheetByName(Tabellen_Blatt);
  var Update_Rang = "D9:D13";
  var Update_Wert = Sheet_Vorlage.getRange(Update_Rang).getRichTextValues();    //Anpassen ob getValue etc



  var Sheet_Training = SpreadsheetApp.openById(LSPD.ID_Training).getSheetByName("Ausbildungsblatt");
  
  var Array_Formeln = Sheet_Training.getRange("E4:E" + Sheet_Training.getLastRow()).getFormulas();

  var Array_IDs = [];

  for(var y = 0; y < Array_Formeln.length; y++)
  {
    Array_IDs.push(Array_Formeln[y][0].substring(Array_Formeln[y][0].indexOf("\"") + 1,Array_Formeln[y][0].indexOf("\"",Array_Formeln[y][0].indexOf("\"") + 1)))
  }

  for(var i = 0; i < Array_IDs.length; i++)
  {
    Logger.log(i + ": Update Tabelle mit ID: " + Array_IDs[i] + " im Tabellenblatt: " + Tabellen_Blatt + " den Bereich: " + Update_Rang);
    SpreadsheetApp.openById(Array_IDs[i]).getSheetByName(Tabellen_Blatt).getRange(Update_Rang).setRichTextValues(Update_Wert);          // <--- Auch ändern ob getValue oder etwas anderes
  }
}
