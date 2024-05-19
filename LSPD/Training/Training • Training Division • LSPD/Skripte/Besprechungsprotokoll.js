function Besprechung_Start()
{
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_Besprechung_Vorlage = SpreadsheetApp.getActive().getSheetByName("Besprechungsprotokoll Vorlage");

  var Datum = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy");

  var Sheet_Besprechung = Sheet_Besprechung_Vorlage.copyTo(SS)

  try
  {
    Sheet_Besprechung.setName("Besprechungsprotokoll " + Datum);
  }
  catch(err)
  {
    try
    {
      Sheet_Besprechung.setName("Besprechungsprotokoll 2 " + Datum);
    }
    catch(err)
    {
      try
      {
        Sheet_Besprechung.setName("Besprechungsprotokoll 3 " + Datum + Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone()," HH:mm"));
      }
      catch(err)
      {
        SpreadsheetApp.getUi().alert("Spam nicht so du Opfer!");

        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(Sheet_Besprechung);

        return 0;
      } 
    } 
  }

  SS.setActiveSheet(Sheet_Besprechung);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(2);

  Sheet_Besprechung.getRange("B1").setValue(Datum);
  Sheet_Besprechung.getRange("F6").setValue(new Date());
  Sheet_Besprechung.getRange("B2").setValue("Besprechungsprotokoll vom " + Datum);
}
function Besprechung_Archivieren()
{
  if(SpreadsheetApp.getActive().getSheetName() != "Besprechungsprotokoll Vorlage")
  {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var Sheet_Besprechung = SpreadsheetApp.getActiveSheet();
    var SS_Export = SpreadsheetApp.openById(LSPD.ID_Archiv_Training_Besprechungen);

    var Datum = Sheet_Besprechung.getRange("B1").getValue();
    var Anzahl = Sheet_Besprechung.getRange("C1").getValue();

    var Sheet_Archiv = Sheet_Besprechung.copyTo(SS_Export);

    try
    {
      Sheet_Archiv.setName(Sheet_Besprechung.getName());
    }
    catch(err)
    {
      try
      {
        Sheet_Archiv.setName(Sheet_Besprechung.getName() + " 2");
      }
      catch(err)
      {
        
        try
        {
          Sheet_Archiv.setName(Sheet_Besprechung.getName() + " 3") + Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone()," HH:mm");
        }
        catch(err)
        {
          SpreadsheetApp.getUi().alert("Spam nicht so du Opfer!");

          SS_Export.deleteSheet(Sheet_Archiv);

          return 0;
        }
      }
    }

    var Link = SS_Export.getUrl() + "#gid=" + Sheet_Archiv.getSheetId();

    Sheet_Archiv.getRange(1,1,Sheet_Archiv.getLastRow(),Sheet_Archiv.getLastColumn()).setValues(Sheet_Besprechung.getRange(1,1,Sheet_Besprechung.getLastRow(),Sheet_Besprechung.getLastColumn()).getValues());

    Sheet_Archiv.getRange(1,1,Sheet_Archiv.getLastRow(),Sheet_Archiv.getLastColumn()).clearDataValidations();

    var Sheet_Uebersicht = SS_Export.getSheetByName("Übersicht");

    Sheet_Uebersicht.getRange("B" + (Sheet_Uebersicht.getLastRow() + 1) + ":C" + (Sheet_Uebersicht.getLastRow() + 1)).setValues([[Datum,Anzahl]])
    Sheet_Uebersicht.getRange("D" + Sheet_Uebersicht.getLastRow()).setFormula("=HYPERLINK(\""+Link+"\";\"Link\")");

    Sheet_Uebersicht.getRange("B5:D" + Sheet_Uebersicht.getLastRow()).sort({column: 2, ascending: false});

    SS.deleteSheet(Sheet_Besprechung);
  }
}

function Besprechung_Anwesenheit(e)
{
  var Sheet = e.source.getActiveSheet();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("C") && Zeile >= 6 && Zeile <= 14 && Value == "TRUE" || Spalte == Spalte_in_Index("E") && Zeile >= 6 && Zeile <= 104 && Value == "TRUE")
  {
    Sheet.getRange(Zeile,Spalte).setValue(Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM HH:mm"));
  }
}