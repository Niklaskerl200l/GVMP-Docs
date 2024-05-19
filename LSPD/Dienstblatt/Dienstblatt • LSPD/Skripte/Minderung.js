function Minderung(e)
{
  var Sheet_Minderung = SpreadsheetApp.getActive().getSheetByName("Minderung");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 6 && Zeile <= 9 && Value != undefined)
  {
    Sheet_Minderung.getRange("G" + Zeile).setValue(LSPD.Umwandeln());
  }
  else if(Spalte == Spalte_in_Index("C") && Zeile >= 6 && Zeile <= 9 && Value != undefined)
  {
    if(Number(Value) >= 90)
    {
      Sheet_Minderung.getRange("E" + Zeile).setValue(30);
    }
    else if(Number(Value) >= 75)
    {
      Sheet_Minderung.getRange("E" + Zeile).setValue(25);
    }
    else if(Number(Value) >= 60)
    {
      Sheet_Minderung.getRange("E" + Zeile).setValue(20);
    }
    else if(Number(Value) >= 45)
    {
      Sheet_Minderung.getRange("E" + Zeile).setValue(15);
    }
    else if(Number(Value) >= 25)
    {
      Sheet_Minderung.getRange("E" + Zeile).setValue(10);
    }
    else if(Number(Value) >= 15)
    {
      Sheet_Minderung.getRange("E" + Zeile).setValue(5);
    }
    else if(Number(Value) < 15)
    {
      Sheet_Minderung.getRange(Zeile, Spalte).clearContent();
      SpreadsheetApp.flush();
      SpreadsheetApp.getUi().alert("Fehler!\nFrühstens ab 15 HE kann eine Minderung vergeben werden!");
    }
  }
  else if(Spalte == Spalte_in_Index("H") && Zeile >= 6 && Zeile <= 9 && Value == "TRUE")
  {
    Sheet_Minderung.getRange(Zeile, Spalte).setValue(false);

    var Lock = LockService.getDocumentLock();
    try
    {
      Lock.waitLock(28000);
    }
    catch(err)
    {
      throw Error("HE Minderung: Zeitüberschreitung");
    }

    var Array_Minderung = Sheet_Minderung.getRange("B" + Zeile + ":G" + Zeile).getValues();

    Sheet_Minderung.insertRowAfter(12);
    Sheet_Minderung.getRange("B13:G13").setValues(Array_Minderung);
    Sheet_Minderung.getRange("H13").setValue(new Date());

    Sheet_Minderung.getRange("B" + Zeile + ":G" + Zeile).clearContent();
    Log_Zaehler("Haftzeitminderung", Array_Minderung[0][0] + " von " + Array_Minderung[0][3]);

    Lock.releaseLock();
  }
  else if(Spalte == Spalte_in_Index("B") && Zeile >=6 && Zeile <=9 && Value =="")

  Sheet_Minderung.getRange(Zeile, Spalte).setValue("")
}

function Auto_Entfernen_Minderung()
{
  var Sheet_Minderung = SpreadsheetApp.getActive().getSheetByName("Minderung");
  var Sheet_Direktion = SpreadsheetApp.openById(LSPD.ID_Direction).getSheetByName("Minderung");

  var Letzte_Zeile = Sheet_Minderung.getLastRow();
  var Array_Minderung = Sheet_Minderung.getRange("B13:H" + Letzte_Zeile).getValues();
  var Datum2 = new Date();
  var Array_Archiv = new Array();
  var Anfangs_Zeile = 0;

  Datum2.setDate(Datum2.getDate() - 3)

  for(var i = Array_Minderung.length - 1; i >= 0; i--)
  {
    if(new Date(Array_Minderung[i][6]) <= Datum2)
    {
      Array_Archiv.push(Array_Minderung[i]);
      Anfangs_Zeile = i + 13;
    }
    else
    {
      break;
    }
  }

  if(Array_Archiv.length > 0)
  {
    var Anzahl = Letzte_Zeile - Anfangs_Zeile + 1;

    Logger.log("Archiviere Minderung");
    Logger.log(Array_Archiv);

    Sheet_Direktion.insertRowsAfter(5,Array_Archiv.length);

    Sheet_Direktion.getRange(6,2,Array_Archiv.length,Array_Archiv[0].length).setValues(Array_Archiv);

    Sheet_Direktion.getRange("B6:H" + Sheet_Direktion.getLastRow()).sort({column: Spalte_in_Index("H"), ascending: false})

    Sheet_Minderung.deleteRows(Anfangs_Zeile,Anzahl);
  }
}