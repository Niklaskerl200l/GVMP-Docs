function Minuten_Trigger_Key()
{
  Minuten_Trigger("Otter");
}

function Minuten_Trigger(Key)
{
  if(Key != "Otter") return;

  var Datum = new Date();

  var Stunde = Datum.getHours();
  var Minute = Datum.getMinutes();

  Logger.log(Stunde + " " + Minute);

  //var Sheet_FIB = SpreadsheetApp.openById("1sMYY7clyNObbbWsm7puCoYs3376zJ6h1TCryBfB9H0w").getSheetByName("Akteneinträge");
  var Sheet_LSPD = SpreadsheetApp.openById(LSPD.ID_GTF).getSheetByName("Akteneinträge");
  var Sheet_Akten = SpreadsheetApp.openById(LSPD.ID_Akteneinträge).getSheetByName("Akteneinträge");

  var Intervall = Sheet_LSPD.getRange("J3").getValue();

  if(Minute % Intervall == 0 && Minute != 99)
  {
    Logger.log("Update Akteneinträge");
    
    var Sheet_Schnittstelle = SpreadsheetApp.getActive().getSheetByName("Akteneinträge");

    var Letzte_Zeile_Schnittstelle = Sheet_Schnittstelle.getLastRow();
    var Letzte_Zeile_LSPD = Sheet_LSPD.getRange("B4").getValue();
    //var Letzte_Zeile_FIB = Sheet_FIB.getRange("B4").getValue();
    var Letzte_Zeile_Akte = Sheet_Akten.getRange("B4").getValue();

    var Array_Schnittstelle = Sheet_Schnittstelle.getRange("B3:G" + Letzte_Zeile_Schnittstelle).getValues();
    var Array_LSPD = Sheet_LSPD.getRange("B7:G" + Letzte_Zeile_LSPD).getValues();
    //var Array_FIB = Sheet_FIB.getRange("B7:G" + Letzte_Zeile_FIB).getValues();
    var Array_Akten = Sheet_Akten.getRange("B7:G" + Letzte_Zeile_Akte).getValues();

    if(Array_Schnittstelle.toString() != Array_LSPD.toString())
    {
      Logger.log("Update Akteneinträge LSPD");

      var Array_Formeln = new Array();

      for(var i = 7; i <= Letzte_Zeile_LSPD; i++)
      {
        Array_Formeln.push(["=D"+i,"=E"+i,"=F"+i]);
      }

      Sheet_LSPD.getRange(7,2,Array_Schnittstelle.length,6).setValues(Array_Schnittstelle);

      if(Array_Schnittstelle.length < Array_LSPD.length)
      {
        Sheet_LSPD.getRange(7 + Array_Schnittstelle.length,2,Letzte_Zeile_LSPD-Array_Schnittstelle.length,6).setValue("");
      }
      
      Sheet_LSPD.getRange("H7:J" + Letzte_Zeile_LSPD).setValues(Array_Formeln);
    }




    /*if(Array_Schnittstelle.toString() != Array_FIB.toString())
    {
      Logger.log("Update Akteneinträge FIB");

      var Array_Formeln = new Array();

      for(var i = 7; i <= Letzte_Zeile_FIB; i++)
      {
        Array_Formeln.push(["=D"+i,"=E"+i,"=F"+i]);
      }

      Sheet_FIB.getRange(7,2,Array_Schnittstelle.length,6).setValues(Array_Schnittstelle);

      if(Array_Schnittstelle.length < Array_FIB.length)
      {
        Sheet_FIB.getRange(7 + Array_Schnittstelle.length,2,Letzte_Zeile_FIB-Array_Schnittstelle.length,6).setValue("");
      }
      
      Sheet_FIB.getRange("H7:J" + Letzte_Zeile_FIB).setValues(Array_Formeln);
    }*/



    if(Array_Schnittstelle.toString() != Array_Akten.toString())
    {
      Logger.log("Update Akteneinträge Akte");

      var Array_Formeln = new Array();

      for(var i = 7; i <= Letzte_Zeile_Akte; i++)
      {
        Array_Formeln.push(["=D"+i,"=E"+i,"=F"+i]);
      }

      Sheet_Akten.getRange(7,2,Array_Schnittstelle.length,6).setValues(Array_Schnittstelle);

      if(Array_Schnittstelle.length < Array_Akten.length)
      {
        Sheet_Akten.getRange(7 + Array_Schnittstelle.length,2,Letzte_Zeile_Akte-Array_Schnittstelle.length,6).setValue("");
      }
      
      Sheet_Akten.getRange("H7:J" + Letzte_Zeile_Akte).setValues(Array_Formeln);
    }
  }


  if(Stunde == 7 && Minute == 57 || Stunde == 15 && Minute == 57 || Stunde == 23 && Minute == 57)
  {
    Logger.log("Sort")
    SpreadsheetApp.getActive().getSheetByName("Aktuell").getRange("B3:K").sort({column: 9, ascending: false})
  }
}
