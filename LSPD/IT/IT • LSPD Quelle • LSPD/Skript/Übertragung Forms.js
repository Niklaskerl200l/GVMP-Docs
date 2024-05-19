function Uebertragung_Forms()
{
  var SheetName = "Imports in Formulare"

  Logger.log("Start für Blatt " + SheetName);

  var Sheet_Imports = SpreadsheetApp.getActive().getSheetByName(SheetName);
  
  var Letzte_Zeile = Sheet_Imports.getLastRow();

  var Array_Imports = Sheet_Imports.getRange("B4:L" + Letzte_Zeile).getValues();
  var Zeit = new Date();
  var Minute = Zeit.getMinutes();


  for(var i = 0; i < Array_Imports.length; i++)
  {
    var Fehler = false;

    try
    {

      if(Minute % Array_Imports[i][1] == 0)
      {
        Logger.log("Funktion: '" + Array_Imports[i][0] + "' Takt: '" + Array_Imports[i][1] + "'\nQuelle ID: '" + Array_Imports[i][3] + "' Blatt: '" + Array_Imports[i][4] + "' Bereich: '" + Array_Imports[i][5] + "'\nZiel ID: '" + Array_Imports[i][7] + "' Dropdown: '" + Array_Imports[i][8]);

        if(Array_Imports[i][3] == "")
        {
          Sheet_Imports.getRange("L" + (i+4)).setValue("Quelle ID Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][4] == "")
        {
          Sheet_Imports.getRange("L" + (i+4)).setValue("Quelle Tabellen Name Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][5] == "")
        {
          Sheet_Imports.getRange("L" + (i+4)).setValue("Quelle Bereich Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][7] == "")
        {
          Sheet_Imports.getRange("L" + (i+4)).setValue("Ziel ID Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][8] == "")
        {
          Sheet_Imports.getRange("L" + (i+4)).setValue("Ziel Dropdown Leer");
          Fehler=true;
        }

        if(Fehler == false)
        {
          var SS_Export = SpreadsheetApp.openById(Array_Imports[i][3]);
          var Sheet_Export = SS_Export.getSheetByName(Array_Imports[i][4]);
          var Range_Export = Sheet_Export.getRange(Array_Imports[i][5]);
          var Array_Export = Range_Export.getValues();

          var Dropdown = FormApp.openById(Array_Imports[i][7]).getItemById(Array_Imports[i][8]).asListItem()

          var Array_Ausgabe = new Array();

          if(Array_Export[0].length >= 2)
          {
            Sheet_Imports.getRange("L" + (i+4)).setValue("Quelle hat mehr als eine Spalte");
            Fehler=true;
          }

          if(Fehler == false)
          {
            for(var x = 0; x < Array_Export.length; x++)
            {
              if(Array_Export[x][0] != "" && Array_Ausgabe.includes(Array_Export[x][0]) == false)
              {
                Array_Ausgabe[x] = Array_Export[x][0]
              }
            }

            Array_Ausgabe = Array_Ausgabe.sort()

            Dropdown.setChoiceValues(Array_Ausgabe);
            Sheet_Imports.getRange("K" + (i+4)).setValue(new Date());
          }
        }
      }
    }
    catch(err)
    {
      Logger.log(err.stack);

      if(Fehler == false)
      {
        Sheet_Imports.getRange("L" + (i+4)).setValue(new Date());
        Sheet_Imports.getRange("M" + (i+4)).setValue(err);
      }
    }
  }

  Logger.log("Ende");
}
