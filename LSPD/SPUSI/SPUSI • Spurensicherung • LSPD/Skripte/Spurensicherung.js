// in Arbeit Fabio

function Sicherung(e) 
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Spurensicherung");

  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Letzte_Zeile = Sheet.getLastRow();

  if(Spalte == Spalte_in_Index("B") || Spalte == Spalte_in_Index("C") && Zeile >= 7 && Zeile <= 12 && Value != undefined)   // Beschlagnahmungen auto. Datum + Uhrzeit, Name
  {
    Sheet.getRange("D" + Zeile + ":G" + Zeile).setValues([[new Date(),new Date(),,LSPD.Umwandeln()]]);
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >= 7 && Zeile <= 12 && Value == undefined)
  {
    Sheet.getRange("B" + Zeile + ":L" + Zeile).setValue("");
  }
  else if(Spalte == Spalte_in_Index("L") && Zeile >=7 && Zeile <= 12 && Value == "TRUE")
  {

   var Array_Schieben = Sheet.getRange("B" + Zeile + ":L" + Zeile).getValues();

   
   if(Array_Schieben[0][0] == "")     
    {
      Logger.log("Ohne Name");
      SpreadsheetApp.getActive().toast("Bitte Namen eingeben bei Beschlagnahmung!")

      return 0;
    }

   Sheet.getRange("L" + Zeile).setValue("");  //checkbox


    Sheet.insertRowBefore(15)
    Sheet.getRange("B15:L15").setValues(Array_Schieben);


    Sheet.getRange("B" + Zeile + ":K" + Zeile).setValue(""); //Löschen

    Sheet.getRange("B13:L13").setValues([[Array_Schieben[0][0], Array_Schieben[0][1], Array_Schieben[0][2], Array_Schieben[0][3], Array_Schieben[0][4], Array_Schieben[0][5], Array_Schieben[0][6], Array_Schieben[0][7], Array_Schieben[0][8] ]]);


 










   Logger.log(Array_Schieben);

    SpreadsheetApp.flush();
    Lock.releaseLock();
  



  }

} 

