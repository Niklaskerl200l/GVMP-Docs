function Minuten_Trigger()
{ 
  var Sheet_Imports = SpreadsheetApp.getActive().getSheetByName("Export");

  var Letzte_Zeile = Sheet_Imports.getLastRow();

  var Array_Imports = Sheet_Imports.getRange("B4:K" + Letzte_Zeile).getValues();
  var Zeit = new Date();
  var Minute = Zeit.getMinutes()

  for(var i = 0; i < Array_Imports.length; i++)
  {
    var Fehler = false;

    try
    {
      if(Minute % Array_Imports[i][1] == 0)
      {
        Logger.log("Start ID: " + Array_Imports[i][2] + " Blatt: " + Array_Imports[i][3] + "   zu   ID: " + Array_Imports[i][5] + " Blatt: " + Array_Imports[i][6]);

        if(Array_Imports[i][2] == "")
        {
          Sheet_Imports.getRange("K" + (i+4)).setValue("Quelle ID Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][3] == "")
        {
          Sheet_Imports.getRange("K" + (i+4)).setValue("Quelle Tabellen Name Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][4] == "")
        {
          Sheet_Imports.getRange("K" + (i+4)).setValue("Quelle Bereich Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][5] == "")
        {
          Sheet_Imports.getRange("K" + (i+4)).setValue("Ziel ID Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][6] == "")
        {
          Sheet_Imports.getRange("K" + (i+4)).setValue("Ziel Tabellen Name Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][7] == "")
        {
          Sheet_Imports.getRange("K" + (i+4)).setValue("Ziel Bereich Leer");
          Fehler=true;
        }

        if(Fehler == false)
        {
          var SS_Export = SpreadsheetApp.openById(Array_Imports[i][2]);
          var Sheet_Export = SS_Export.getSheetByName(Array_Imports[i][3]);
          var Range_Export = Sheet_Export.getRange(Array_Imports[i][4]);
          var Array_Export = Range_Export.getValues();

          var SS_Import = SpreadsheetApp.openById(Array_Imports[i][5]);
          var Sheet_Import = SS_Import.getSheetByName(Array_Imports[i][6]);
          var Range_Import = Sheet_Import.getRange(Array_Imports[i][7]);
          var Array_Import = Range_Import.getValues();

          if(Array_Import.length != Array_Export.length)
          {
            Sheet_Imports.getRange("K" + (i+4)).setValue("Falsche Zeilen Anzahl (Quelle Zeilen: " + Array_Import.length + " | Ziel Zeilen: " + Array_Export.length + ")");
            Fehler=true;
          }
          else if(Array_Import[0].length != Array_Export[0].length)
          {
            Sheet_Imports.getRange("K" + (i+4)).setValue("Falsche Spalten Anzahl (Quelle Spalten: " + Array_Import[0].length + " | Ziel Spalten: " + Array_Export[0].length + ")");
            Fehler=true;
          }
          

          if(Array_Export.toString() != Array_Import.toString())
          {
            for(var y = 0; y < Array_Export.length; y++)
            {
              for(var x = 0; x < Array_Export[0].length; x++)
              {
                if(Array_Export[y][x].toString() != Array_Import[y][x].toString())
                {
                  Logger.log("\tUpdate: y: " + (Range_Import.getRow() + y) + "  x: " + (Range_Import.getColumn() + x))
                  Sheet_Import.getRange(Range_Import.getRow() + y, Range_Import.getColumn() + x).setValue(Array_Export[y][x]);
                }
              }
            }
          }
          
          if(Array_Imports[i][9].toString().length > 64)
          {
            Sheet_Imports.getRange("K" + (i+4)).setValue("");
          }

          Sheet_Imports.getRange("J" + (i+4)).setValue(new Date());
        }
      }
    }
    catch(err)
    {
      Logger.log(err.stack);

      if(Fehler == false)
      {
        Sheet_Imports.getRange("K" + (i+4)).setValue(new Date());
      }
    }
  }
}