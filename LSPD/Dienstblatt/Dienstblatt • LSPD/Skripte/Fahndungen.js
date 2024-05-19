function Fahndungen(e)
{
  var Sheet_Fahndungen = SpreadsheetApp.getActive().getSheetByName("Fahndungen");
  var Spalte = e.range.getColumn();
  var Zeile = e.range.getRow();
  var Value = e.value;
  var OldValue = e.oldValue;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 6 && Zeile <= 35 && Value != undefined && OldValue == undefined)
  {
    Sheet_Fahndungen.getRange("G" + Zeile).setValue(new Date());
    Sheet_Fahndungen.getRange("H" + Zeile).setValue(LSPD.Umwandeln());

    Log_Zaehler("Fahndungen\nEingetragen", Value);
  }
  else if(Spalte == Spalte_in_Index("I") && Zeile >= 6 && Zeile <= 35 && Value == "TRUE")
  {
    Sheet_Fahndungen.getRange(Zeile, Spalte).setValue(false);

    var Array_Fahndung = Sheet_Fahndungen.getRange("B" + Zeile + ":H" + Zeile).getValues();
    Array_Fahndung = Array_Fahndung[0];
    
    if(Array_Fahndung[4] == true)
    {
      SpreadsheetApp.flush();

      var UI = SpreadsheetApp.getUi();
      UI.alert("Fehler!", "Diese Akte ist in Klärung und kann demnach nicht vergeben werden!\nFalls dies nicht der Fall sein sollte, so entferne den Haken unter '🚨'.", UI.ButtonSet.OK);

      return;
    }

    var Lock = LockService.getDocumentLock();
    try
    {
      Lock.waitLock(28000);
    }
    catch(err)
    {
      throw Error("Fahndungen: Zeitüberschreitung!");
    }

    Sheet_Fahndungen.insertRowAfter(38);
    Sheet_Fahndungen.getRange("F39:G39").merge();
    Sheet_Fahndungen.getRange("B39:I39").setValues(
      [[
        Array_Fahndung[0],
        Array_Fahndung[1],
        Array_Fahndung[2],
        Array_Fahndung[3],
        Array_Fahndung[5],
        "",
        Array_Fahndung[6],
        LSPD.Umwandeln() + "\n" + Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy HH:mm")
      ]]
    );

    Sheet_Fahndungen.getRange("B" + Zeile + ":H" + Zeile).clearContent();
    Sheet_Fahndungen.getRange("F" + Zeile).setValue(false).insertCheckboxes();

    Sheet_Fahndungen.getRange("B6:H35").sort({column: Spalte_in_Index("G"), ascending: true});

    Log_Zaehler("Akte\nVergeben", Array_Fahndung[0]);
  }
}

function Fahndungen_Verjaehrung()
{
  var Sheet_Fahndungen = SpreadsheetApp.getActive().getSheetByName("Fahndungen");
  var Array_Fahndungen = Sheet_Fahndungen.getRange("B6:H35").getValues();

  var Zeitgrenze = new Date();
  Zeitgrenze.setDate(Zeitgrenze.getDate() - 28);

  for(var i = 0; i < Array_Fahndungen.length; i++)
  {
    if(Array_Fahndungen[i][0] != "" && Array_Fahndungen[i][5] <= Zeitgrenze)
    {
      Sheet_Fahndungen.insertRowAfter(38);
      Sheet_Fahndungen.getRange("F39:G39").merge();

      Sheet_Fahndungen.getRange("B39:I39").setValues(
        [[
          Array_Fahndungen[i][0],
          Array_Fahndungen[i][1],
          Array_Fahndungen[i][2],
          Array_Fahndungen[i][3],
          Array_Fahndungen[i][5],
          "",
          Array_Fahndungen[i][6],
          "Verjährt am\n" + Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy HH:mm")
        ]]
      );

      Sheet_Fahndungen.getRange("B" + (i + 6) + ":H" + (i + 6)).clearContent();
      Sheet_Fahndungen.getRange("F" + (i + 6)).setValue(false).insertCheckboxes();
    }
  }

  Sheet_Fahndungen.getRange("B6:H35").sort({column: Spalte_in_Index("G"), ascending: true});
}

function Fahndungen_Archivierung()
{
  var Sheet_Fahndungen = SpreadsheetApp.getActive().getSheetByName("Fahndungen");
  var Array_Fahndungen = Sheet_Fahndungen.getRange("B38:I").getValues();

  var Zeitgrenze = new Date();
  Zeitgrenze.setDate(Zeitgrenze.getDate() - 90);

  var Archiv_Array = [];

  for(var i = Array_Fahndungen.length - 1; i >= 0; i--)
  {
    if(Array_Fahndungen[i][0] != "" && new Date(Array_Fahndungen[i][4]) <= Zeitgrenze || Array_Fahndungen[i][0] != "" && Array_Fahndungen[i][4] == "")
    {
      Archiv_Array.push(
        [
          Array_Fahndungen[i][0],
          Array_Fahndungen[i][1],
          Array_Fahndungen[i][2],
          Array_Fahndungen[i][3],
          Array_Fahndungen[i][6],
          Array_Fahndungen[i][4],
          Array_Fahndungen[i][7]
        ]
      );

      Logger.log(Archiv_Array);
      Sheet_Fahndungen.deleteRow(i + 38);
    }
  }

  if(Archiv_Array.length > 0)
  {
    return;

    var SS_Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_Aktenvergabe);
    var Sheet_Archiv = SS_Archiv.getRange("Archiv");

    Sheet_Archiv.insertRowsAfter(3, Archiv_Array.length);
    Sheet_Archiv.getRange("B4:H" + (Archiv_Array.length + 3)).setValues(Archiv_Array);
  }
}

function Fahndungen_Sortieren()
{
  var Sheet_Fahndungen = SpreadsheetApp.getActive().getSheetByName("Fahndungen");
  Sheet_Fahndungen.getRange("B6:H35").sort({column: Spalte_in_Index("G"), ascending: true});
}