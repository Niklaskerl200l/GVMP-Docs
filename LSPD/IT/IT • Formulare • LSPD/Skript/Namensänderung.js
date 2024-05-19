function Namensaenderung(e)
{
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var error = undefined;
  var Fehler = false;
  var Werte = e.namedValues;

  var Alter_Name = Werte["Alter Name"];
  var Neuer_Name = Werte["Neuer Name"];

  var Array_ID = Sheet_Auswertung.getRange("B3:C" + Sheet_Auswertung.getLastRow()).getValues();

  Logger.log("Ersetzen " + Alter_Name + " zu " + Neuer_Name);
  for(var y = 0; y < Array_ID.length; y++)
  {
    try
    {
      Logger.log(Array_ID[y][0] + ":\t" + SpreadsheetApp.openById(Array_ID[y][1]).createTextFinder(Alter_Name).replaceAllWith(Neuer_Name) + " mal ersetzt");
    }
    catch(err)
    {
      error = err;
      Logger.log(err.stack);
      SpreadsheetApp.getActive().getSheetByName("Namensänderung").getRange(e.range.getRow(),e.range.getLastColumn() + (y + 1)).setValue(false).insertCheckboxes();
      Fehler = true;
    }
    if(error == undefined)
    {
      SpreadsheetApp.getActive().getSheetByName("Namensänderung").getRange(e.range.getRow(),e.range.getLastColumn() + (y + 1)).setValue(true).insertCheckboxes();
    }
  }

  try
  {
    Alter_Name = Alter_Name.toString().replace(" ","_");
    Neuer_Name = Neuer_Name.toString().replace(" ","_");

    Logger.log(Alter_Name + " " + SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle).createTextFinder(Alter_Name).replaceAllWith(Neuer_Name) + " mal ersetzt");
  }
  catch(err)
  {}

  if(Fehler == true)
  {
    Problem;
  }
}

function temp()
{
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Namensänderung");
  var Array_temp = Sheet_Auswertung.getRange("B33:C46").getValues();

  for(var i = 0; i < Array_temp.length; i++)
  {
    Namensaenderung_Manuell(Array_temp[i][0],Array_temp[i][1])
  }
}

function Namensaenderung_Manuell(Alter_Name = "Leory Loewe", Neuer_Name = "Leroy Loewe")
{
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");

  var error = undefined;
  var Fehler = false;

  var Array_ID = Sheet_Auswertung.getRange("B3:C" + Sheet_Auswertung.getLastRow()).getValues();

  Logger.log("Ersetzen " + Alter_Name + " zu " + Neuer_Name);
  for(var y = 0; y < Array_ID.length; y++)
  {
    try
    {
      //Logger.log("ID: " + Array_ID[y][1]);
      var Sheet = SpreadsheetApp.openById(Array_ID[y][1]);
      var Anzahl = Sheet.createTextFinder(Alter_Name).replaceAllWith(Neuer_Name)
      Logger.log(Array_ID[y][0] + ":\t" + Anzahl + " mal ersetzt");
    }
    catch(err)
    {
      error = err;
      Logger.log(err.stack);
      //Fehler = true;
    }
  }

  try
  {
    Alter_Name = Alter_Name.toString().replace(" ","_");
    Neuer_Name = Neuer_Name.toString().replace(" ","_");

    Logger.log(Alter_Name + " " + SpreadsheetApp.openById(LSPD.ID_GTF_DCI_Schnittstelle).createTextFinder(Alter_Name).replaceAllWith(Neuer_Name) + " mal ersetzt");
  }
  catch(err)
  {}

  if(Fehler == true)
  {
    Problem;
  }
}
