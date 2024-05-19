function Install_onEdit(e)
{
  var SheetName = e.source.getActiveSheet().getName();
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Off_Name = 0, Off_Fraktion = 1, Off_Rang = 2, Off_Haus = 3, Off_Tel = 4, Off_Aktivitaet = 7, Off_Beamter = 8, Off_Akteneintrag = 9, Off_Datenbank = 10;

  //------------------------------------- Akteneinträge ----------------------------------------//

  if(SheetName == "Akteneinträge" && Spalte == Spalte_in_Index("K") && Zeile >= 7 && Value == "TRUE")
  {
    var SS_Datenbank = SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle);

    var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Akteneinträge");
    var Sheet_GTF = SpreadsheetApp.openById(LSPD.ID_GTF).getSheetByName("Akteneinträge");
    var Sheet_Eintrag = SpreadsheetApp.getActive().getSheetByName("Akteneinträge");
    var Sheet_Aktuell = SS_Datenbank.getSheetByName("Aktuell");

    var Array_DB_Aktuell = Sheet_Aktuell.getRange("B3:B" + Sheet_Aktuell.getLastRow()).getValues();
    var Array_Eintrag = Sheet_Eintrag.getRange("B" + Zeile + ":J" + Zeile).getValues();
    
    for(var i = 0; i < Array_DB_Aktuell.length; i++)
    {
      if(Array_DB_Aktuell[i][Off_Name] == Array_Eintrag[0][0])   // Suche nach Namen in DB Aktuell
      {
        Logger.log("Setzt Akteneintrag von: " + Array_DB_Aktuell[i][Off_Name]);
        Sheet_Aktuell.getRange("K" + (i+3)).setValue(new Date());   // Setze Datum

        if(Array_Eintrag[0][2] != Array_Eintrag[0][6])
        {
          Logger.log("Setzt Rang von: " + Array_DB_Aktuell[i][Off_Name] + " auf: " + Array_Eintrag[0][6]);
          Sheet_Aktuell.getRange("D" + (i+3)).setValue(Array_Eintrag[0][6]);   // Setze Rang
        }

        if(Array_Eintrag[0][3] != Array_Eintrag[0][7])
        {
          Logger.log("Setzt Haus Nummer von: " + Array_DB_Aktuell[i][Off_Name] + " auf: " + Array_Eintrag[0][7]);
          Sheet_Aktuell.getRange("E" + (i+3)).setValue(Array_Eintrag[0][7]);   // Setze HN
        }

        if(Array_Eintrag[0][4] != Array_Eintrag[0][8]) 
        {
          Logger.log("Setzt Tel von: " + Array_DB_Aktuell[i][Off_Name] + " auf: " + Array_Eintrag[0][8]);
          Sheet_Aktuell.getRange("F" + (i+3)).setValue(Array_Eintrag[0][8]);   // Setze Tel
        }

        Sheet_Eintrag.getRange("B" + Zeile +":K" + Zeile).setValue("");
        Sheet_Eintrag.getRange("H" + Zeile +":J" + Zeile).setValues([["=D"+Zeile,"=E"+Zeile,"=F"+Zeile]]);

        if(Sheet_FIB.getRange("B" + Zeile).getValue() == Array_Eintrag[0][0])
        {
          Sheet_FIB.getRange("B" + Zeile +":K" + Zeile).setValue("");
          Sheet_FIB.getRange("H" + Zeile +":J" + Zeile).setValues([["=D"+Zeile,"=E"+Zeile,"=F"+Zeile]]);
        }
        else
        {
          var Array_FIB =  Sheet_FIB.getRange("B7:B").getValues();

          for(var y = 0; y < Array_FIB.length; y++)
          {
            if(Array_FIB[y][0] == Array_Eintrag[0][0])
            {
              Sheet_FIB.getRange("B" + (y+7) +":K" + (y+7)).setValue("");
              Sheet_FIB.getRange("H" + (y+7) +":J" + (y+7)).setValues([["=D"+(y+7),"=E"+(y+7),"=F"+(y+7)]]);
              break;
            }
          }
        }

        if(Sheet_GTF.getRange("B" + Zeile).getValue() == Array_Eintrag[0][0])
        {
          Sheet_GTF.getRange("B" + Zeile +":K" + Zeile).setValue("");
          Sheet_GTF.getRange("H" + Zeile +":J" + Zeile).setValues([["=D"+Zeile,"=E"+Zeile,"=F"+Zeile]]);
        }
        else
        {
          var Array_GTF =  Sheet_GTF.getRange("B7:B").getValues();

          for(var y = 0; y < Array_GTF.length; y++)
          {
            if(Array_GTF[y][0] == Array_Eintrag[0][0])
            {
              Sheet_GTF.getRange("B" + (y+7) +":K" + (y+7)).setValue("");
              Sheet_GTF.getRange("H" + (y+7) +":J" + (y+7)).setValues([["=D"+(y+7),"=E"+(y+7),"=F"+(y+7)]]);
              break;
            }
          }
        }

        break;
      }
    }
  }

  else if(SheetName == "Akteneinträge" && Spalte == Spalte_in_Index("J") && Zeile == 3)
  {
    var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Akteneinträge");
    var Sheet_GTF = SpreadsheetApp.openById(LSPD.ID_GTF).getSheetByName("Akteneinträge");

    Sheet_FIB.getRange("J3").setValue(Value);
    Sheet_GTF.getRange("J3").setValue(Value);
  }

  else if(SheetName == "Akteneinträge" && Spalte >= Spalte_in_Index("H") && Spalte <= Spalte_in_Index("J") && Zeile >= 7)
  {
    var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Akteneinträge");
    var Sheet_GTF = SpreadsheetApp.openById(LSPD.ID_GTF).getSheetByName("Akteneinträge");
    var Sheet_Eintrag = SpreadsheetApp.getActive().getSheetByName("Akteneinträge");

    if(Value == "" || Value == undefined)
    {
      Sheet_FIB.getRange(Zeile,Spalte).setFormula(Sheet_Eintrag.getRange(Zeile,Spalte).getFormula());
      Sheet_GTF.getRange(Zeile,Spalte).setFormula(Sheet_Eintrag.getRange(Zeile,Spalte).getFormula());
    }
    else
    {
      Sheet_FIB.getRange(Zeile,Spalte).setValue(Value);
      Sheet_GTF.getRange(Zeile,Spalte).setValue(Value);
    }
  }

  //-------------------------------------------- ENDE ------------------------------------------//

}