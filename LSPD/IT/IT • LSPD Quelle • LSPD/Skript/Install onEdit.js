function Install_onEdit(e)
{
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  Logger.log("Sheet: " + SheetName + "\nZeile: " + Zeile + "\tSpalte: " + Spalte + "\nAlte Value: " + OldValue + "\nValue: " + Value);


  if(SheetName.startsWith("Imports"))
  {
    if((Spalte == Spalte_in_Index("D") || Spalte == Spalte_in_Index("H")) && Zeile >= 4 && Value != "" && Value != undefined)
    {
      var Sheet_Imports = SpreadsheetApp.getActive().getSheetByName(SheetName);

      var Eingabe_Name = Sheet_Imports.getRange(Zeile,Spalte).getValue();
      var Eingabe_ID = Sheet_Imports.getRange(Zeile,Spalte + 1).getValue();
      var ID = Eingabe_ID;
      
      if(Eingabe_Name != "")
      {
        var Sheet_IDs = SpreadsheetApp.getActive().getSheetByName("Import IDs");

        var Array_Ids = Sheet_IDs.getRange("B3:C").getValues()

        for(var i = 0; i < Array_Ids.length; i++)
        {
          if(Array_Ids[i][0] == Eingabe_Name)
          {
            ID = Array_Ids[i][1];
            break;
          }
        }
      }


      if(ID != "" && ID != undefined)
      {
        try
        {
          var SS_Temp = SpreadsheetApp.openById(ID);

          var Array_Sheets = SS_Temp.getSheets();
          var Array_Namen = new Array();

          for(var i = 0; i < Array_Sheets.length; i++)
          {
            Array_Namen.push(Array_Sheets[i].getName());
          }

          Sheet_Imports.getRange(Zeile,Spalte+2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array_Namen).build())
        }
        catch(err)
        {
          Logger.log(err.stack);
          Fehler;
        }
      }
    }

    else if((Spalte == Spalte_in_Index("D") || Spalte == Spalte_in_Index("H")) && Zeile >= 4 && (Value == "" || Value == undefined))
    {
      var Sheet_Imports = SpreadsheetApp.getActive().getSheetByName("Imports");

      Sheet_Imports.getRange(Zeile,Spalte+2).clearDataValidations();
    }


    else if((Spalte == Spalte_in_Index("E") || Spalte == Spalte_in_Index("I")) && Zeile >= 4 && Value != "" && Value != undefined)
    {
      var Sheet_Imports = SpreadsheetApp.getActive().getSheetByName(SheetName);

      var ID = Value;

      if(ID != "" && ID != undefined)
      {
        try
        {
          var SS_Temp = SpreadsheetApp.openById(ID);

          var Array_Sheets = SS_Temp.getSheets();
          var Array_Namen = new Array();

          for(var i = 0; i < Array_Sheets.length; i++)
          {
            Array_Namen.push(Array_Sheets[i].getName());
          }

          Sheet_Imports.getRange(Zeile,Spalte+1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array_Namen).build())
          Sheet_Imports.getRange(Zeile,Spalte).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(ID).setLinkUrl("https://docs.google.com/spreadsheets/d/" + ID).build());
        }
        catch(err)
        {
          Logger.log(err.stack);
          Fehler;
        }
      }
    }

    else if((Spalte == Spalte_in_Index("E") || Spalte == Spalte_in_Index("I")) && Zeile >= 4 && (Value == "" || Value == undefined))
    {
      var Sheet_Imports = SpreadsheetApp.getActive().getSheetByName(SheetName);
      
      Sheet_Imports.getRange(Zeile,Spalte+1).clearDataValidations();
    }
  }
}
