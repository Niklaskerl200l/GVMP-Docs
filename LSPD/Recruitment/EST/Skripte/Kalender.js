function Kalender(e)
{
  var Zeile_Start = e.range.rowStart;
  var Zeile_Ende = e.range.rowEnd;
  var Spalte_Start = e.range.columnStart;
  var Spalte_Ende = e.range.columnEnd;
  var Values = e.range.getValues()

  for(var y = Zeile_Start; y <= Zeile_Ende; y++)
  {
    for(var x = Spalte_Start; x <= Spalte_Ende; x++)
    {
      var Temp_e = e;

      Temp_e.range.columnEnd = x;
      Temp_e.range.columnStart = x;
      Temp_e.range.rowEnd = y;
      Temp_e.range.rowStart = y;

      Temp_e.range.getRow = function() {return y}
      Temp_e.range.getColumn = function() {return x}
      Temp_e.value = Values[y-Zeile_Start][x-Spalte_Start];

      Kalender_onEdit(Temp_e);
    }
  }
}

function Kalender_onEdit(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  Logger.log("Zeile: " + Zeile + "\tSpalte: " + Spalte  + "\nValue: " + Value);

  var Sheet_Kalender = SpreadsheetApp.getActive().getSheetByName("Kalender");
  var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Archiv Kalender");

  var Array_Archiv = Sheet_Archiv.getRange(2,2,1,Sheet_Archiv.getLastColumn() - 1).getValues();

  var Datum = Sheet_Kalender.getRange(4,Spalte).getValue();
  var Uhrzeit = Sheet_Kalender.getRange(Zeile,Spalte_in_Index("B")).getValue();

  Logger.log(Datum);
  Logger.log(Uhrzeit);

  for(var x = 1; x < Array_Archiv[0].length; x++)
  {
    if(Array_Archiv[0][x].toString() == Datum.toString())
    {
      Sheet_Archiv.getRange(Zeile - 3, x + 2).setValue(Value);
      break;
    }
  }
}

function Kalender_Links()
{
  var Sheet_Kalender = SpreadsheetApp.getActive().getSheetByName("Kalender");
  var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Archiv Kalender");

  var Array_Archiv = Sheet_Archiv.getRange(2,2,17,Sheet_Archiv.getLastColumn() - 1).getValues();
  var Array_Ausgabe = new Array();

  var Start_Datum = Sheet_Kalender.getRange("C4").getValue();

  for(var x = 1; x < Array_Archiv[0].length; x++)
  {
    if(Array_Archiv[0][x].toString() == Start_Datum.toString())
    {
      if(x-1 == 0)
      {
        SpreadsheetApp.getUi().alert("Ende");
        return 0;
      }

      Array_Ausgabe = Kalender_Laden(x-1,Array_Archiv)
      break;
    }
  }

  Logger.log(Array_Ausgabe)

  Sheet_Kalender.getRange("C4:I4").setValues([Array_Ausgabe[0]]);

  Array_Ausgabe.shift();

  Sheet_Kalender.getRange("C6:I21").setValues(Array_Ausgabe);
}

function Kalender_Rechts()
{
  var Sheet_Kalender = SpreadsheetApp.getActive().getSheetByName("Kalender");
  var Sheet_Archiv = SpreadsheetApp.getActive().getSheetByName("Archiv Kalender");

  var Array_Archiv = Sheet_Archiv.getRange(2,2,17,Sheet_Archiv.getLastColumn() - 1).getValues();
  var Array_Ausgabe = new Array();

  var Start_Datum = Sheet_Kalender.getRange("C4").getValue();
  var End_Datum = Sheet_Archiv.getRange(2,Sheet_Archiv.getLastColumn()).getValue()

  for(var x = 1; x < Array_Archiv[0].length; x++)
  {
    if(Array_Archiv[0][x].toString() == Start_Datum.toString())
    {
      if(x+7 <= Array_Archiv[0].length)
      {
        var Letzes_Datum = new Date(End_Datum);
        Letzes_Datum.setDate(Letzes_Datum.getDate() + 1);

        Sheet_Archiv.getRange(2,Sheet_Archiv.getLastColumn()+1).setValue(Letzes_Datum);

        SpreadsheetApp.flush();

        Array_Archiv = Sheet_Archiv.getRange(2,2,17,Sheet_Archiv.getLastColumn() - 1).getValues();
      }

      Array_Ausgabe = Kalender_Laden(x+1,Array_Archiv)
      break;
    }
  }

  //Logger.log(Array_Ausgabe)

  Sheet_Kalender.getRange("C4:I4").setValues([Array_Ausgabe[0]]);

  Array_Ausgabe.shift();

  Sheet_Kalender.getRange("C6:I21").setValues(Array_Ausgabe);
}


function Kalender_Laden(temp,Array_Archiv)
{
  Logger.log(Array_Archiv)
  var Array_Ausgabe = new Array();

  for(var i = 0, y = 0; i < Array_Archiv.length; i++, y++)
  {
    Array_Ausgabe[y] = new Array();

    for(var i2 = temp, x = 0; i2 <= (temp + 6); i2++, x++)
    {
      Array_Ausgabe[y][x] = Array_Archiv[i][i2];
    }
  }
  Logger.log(Array_Ausgabe)
  
  return Array_Ausgabe;
}



function Kalender_Tages_Wechsel()
{
  var Sheet_Kalender = SpreadsheetApp.getActive().getSheetByName("Kalender");

  var Start_Datum = Sheet_Kalender.getRange("C4").getValue();
  var Heute_Datum = new Date();

  var Tage = Heute_Datum.getDate() - Start_Datum.getDate();

  if(Tage < 0)
  {
    for(var i = 0; i > Tage; i--)
    {
      Kalender_Links();
    }
  }
  else if(Tage > 0)
  {
    for(var i = 0; i < Tage; i++)
    {
      Kalender_Rechts();
    }
  }
}