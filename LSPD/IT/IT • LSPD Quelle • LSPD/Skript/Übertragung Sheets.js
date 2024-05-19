function Uebertragung(SheetName)
{
  Logger.log("Start für Blatt " + SheetName);

  var Sheet_Imports = SpreadsheetApp.getActive().getSheetByName(SheetName);
  
  var Letzte_Zeile = Sheet_Imports.getLastRow();

  var Array_Imports = Sheet_Imports.getRange("B4:M" + Letzte_Zeile).getValues();
  var Zeit = new Date();
  var Minute = Zeit.getMinutes();

  try
  {
    Logger.log("Update ID Doc");

    var Sheet_ID = SpreadsheetApp.openById("1Iw95dUiLpdJwSLnYPp5IEe4RCSOayM8QEEiEk8gTfUk").getSheetByName("ID");

    Sheet_ID.getRange(3,2,LSPD.Array_URLs.length,4).setValues(LSPD.Array_URLs);
  }
  catch(err)
  {
    Logger.log("Fehler beim Update des ID Doc")
    Logger.log(err.stack);
    Sheet_Imports.getRange("M4:N4").setValues([[new Date(),err]]);
    return 0;
  }

  for(var i = 0; i < Array_Imports.length; i++)
  {
    var Fehler = false;

    try
    {

      if(Minute % Array_Imports[i][1] == 0)
      {
        Logger.log("Funktion: '" + Array_Imports[i][0] + "' Takt: '" + Array_Imports[i][1] + "'\nQuelle ID: '" + Array_Imports[i][3] + "' Blatt: '" + Array_Imports[i][4] + "' Bereich: '" + Array_Imports[i][5] + "'\nZiel ID: '" + Array_Imports[i][7] + "' Blatt: '" + Array_Imports[i][8] + "' Bereich: '" + Array_Imports[i][9] + "'");

        if(Array_Imports[i][3] == "")
        {
          Sheet_Imports.getRange("M" + (i+4)).setValue("Quelle ID Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][4] == "")
        {
          Sheet_Imports.getRange("M" + (i+4)).setValue("Quelle Tabellen Name Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][5] == "")
        {
          Sheet_Imports.getRange("M" + (i+4)).setValue("Quelle Bereich Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][7] == "")
        {
          Sheet_Imports.getRange("M" + (i+4)).setValue("Ziel ID Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][8] == "")
        {
          Sheet_Imports.getRange("M" + (i+4)).setValue("Ziel Tabellen Name Leer");
          Fehler=true;
        }
        else if(Array_Imports[i][9] == "")
        {
          Sheet_Imports.getRange("M" + (i+4)).setValue("Ziel Bereich Leer");
          Fehler=true;
        }

        if(Fehler == false)
        {
          var SS_Export = SpreadsheetApp.openById(Array_Imports[i][3]);
          var Sheet_Export = SS_Export.getSheetByName(Array_Imports[i][4]);
          var Range_Export = Sheet_Export.getRange(Array_Imports[i][5]);
          var Array_Export = Range_Export.getValues();

          var SS_Import = SpreadsheetApp.openById(Array_Imports[i][7]);
          var Sheet_Import = SS_Import.getSheetByName(Array_Imports[i][8]);
          var Range_Import = Sheet_Import.getRange(Array_Imports[i][9]);
          var Array_Import = Range_Import.getValues();
          
          var Array_Updates = new Array();

          if(Array_Import.length != Array_Export.length)
          {
            Sheet_Imports.getRange("M" + (i+4)).setValue("Falsche Zeilen Anzahl (Quelle Zeilen: " + Array_Export.length + " | Ziel Zeilen: " + Array_Import.length + ")");
            Fehler=true;
          }
          else if(Array_Import[0].length != Array_Export[0].length)
          {
            Sheet_Imports.getRange("M" + (i+4)).setValue("Falsche Spalten Anzahl (Quelle Spalten: " + Array_Export[0].length + " | Ziel Spalten: " + Array_Import[0].length + ")");
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
                  Array_Updates.push([y,x])
                }
              }
            }

            Logger.log("\tÄnderungen: " + Array_Updates.length + " von " + Array_Export.length * Array_Export[0].length + " möglichen");

            var Off_Zeile = Range_Import.getRow();
            var Off_Spalte = Range_Import.getColumn();

            if(Array_Updates.length >= (Array_Export.length * Array_Export[0].length) / 100 * 5)
            {
              Logger.log("\tSetze Daten kommplett neu da mehr als " + ((Array_Export.length * Array_Export[0].length) / 100 * 5) + " (5%) änderungen!");
              if(Array_Updates.length <= 100)
              {
                var Ausgabe = "";

                Array_Updates.forEach(x => Ausgabe += "\tUpdate: Zeile: " + (Off_Zeile + x[0]) + "\tSpalte: " + (Off_Spalte + x[1]) + "\tVon: '" + Array_Import[x[0]][x[1]] + "' zu '" + Array_Export[x[0]][x[1]] + "'\n");

                Logger.log(Ausgabe);
              }


              Sheet_Import.getRange(Off_Zeile,Off_Spalte,Array_Export.length, Array_Export[0].length).setValues(Array_Export);
            }
            else
            {
              for(var y = 0; y < Array_Updates.length; y++)
              {
                Logger.log("\tUpdate: Zeile: " + (Off_Zeile + Array_Updates[y][0]) + "\tSpalte: " + (Off_Spalte + Array_Updates[y][1]) + "\tVon: '" + Array_Import[Array_Updates[y][0]][Array_Updates[y][1]] + "' zu '" + Array_Export[Array_Updates[y][0]][Array_Updates[y][1]] + "'");
                Sheet_Import.getRange(Off_Zeile + Array_Updates[y][0],Off_Spalte + Array_Updates[y][1]).setValue(Array_Export[Array_Updates[y][0]][Array_Updates[y][1]]);
              }
            }
          }

          if(!(Array_Imports[i][11] instanceof Date && !isNaN(Array_Imports[i][11].valueOf())))
          {
            Sheet_Imports.getRange("M" + (i+4) + ":N" + (i+4)).setValue("");
          }

          Sheet_Imports.getRange("L" + (i+4)).setValue(new Date());
        }
      }
    }
    catch(err)
    {
      Logger.log(err.stack);

      if(Fehler == false)
      {
        Sheet_Imports.getRange("M" + (i+4)).setValue(new Date());
        Sheet_Imports.getRange("N" + (i+4)).setValue(err);
      }
    }
  }

  Logger.log("Ende");
}
