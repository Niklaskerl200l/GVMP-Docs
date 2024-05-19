function Streifenkontrolle(e)
{
  var Sheet_Streifenkontrolle = SpreadsheetApp.getActive().getSheetByName("Streifenkontrolle");
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("B") && Zeile >= 3 && Zeile <= 8)
  {
    if(Value != undefined)
    {
      Sheet_Streifenkontrolle.getRange("H" + Zeile).setNote("/alktest " + Value.toString().trim().replace(" ", "_"));
      Sheet_Streifenkontrolle.getRange("I" + Zeile).setNote("/drugtest " + Value.toString().trim().replace(" ", "_"));

      var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Import Personaltabelle");
      var Array_Personal = Sheet_Personal.getRange("B4:D199").getValues();

      var Gefunden = false;
      for(var i = 0; i < Array_Personal.length; i++)
      {
        if(Array_Personal[i][0] != "" && Array_Personal[i][2].toString() == Value)
        {
          Gefunden = true;
          break;
        }
      }

      if(Gefunden == true)
      {
        var Bewaffnung;
        switch(Array_Personal[i][0])
        {
          case 0: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle"; break;
          case 1: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle"; break;
          case 2: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle"; break;
          case 3: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle"; break;
          case 4: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle"; break;
          case 5: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle"; break;
          case 6: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle"; break;
          case 7: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle\nMilitaryrifle"; break;
          case 8: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle\nMilitaryrifle\nTacticalrifle\nHeavyshotgun"; break;
          case 9: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle\nMilitaryrifle\nTacticalrifle\nHeavyshotgun"; break;
          case 10: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle\nMilitaryrifle\nTacticalrifle\nHeavyshotgun"; break;
          case 11: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle\nMilitaryrifle\nTacticalrifle\nHeavyshotgun"; break;
          case 12: Bewaffnung = "Heavypistol\nSMG\nAdvancedrifle\nBullpuprifle\nCarbinerifle\nMilitaryrifle\nTacticalrifle\nHeavyshotgun"; break;
        }

        Sheet_Streifenkontrolle.getRange("F" + Zeile).setNote(Bewaffnung);
      }
    }
    else if(Value == undefined)
    {
      Sheet_Streifenkontrolle.getRange("F" + Zeile + ":I" + Zeile).clearNote();
    }
  }
  else if(Spalte == Spalte_in_Index("L") && Zeile >= 3 && Zeile <= 8 && Value == "TRUE")
  {
    Sheet_Streifenkontrolle.getRange(Zeile, Spalte).setValue(false).insertCheckboxes();

    var Lock = LockService.getScriptLock();
    try
    {
      Lock.waitLock(28000);
    }
    catch(err)
    {
      throw Error("Streifenkontrollen: Zeitüberschreitung!");
    }

    var Array_Kontrolle = Sheet_Streifenkontrolle.getRange("B" + Zeile + ":K" + Zeile).getValues();

    Sheet_Streifenkontrolle.insertRowAfter(11);
    Sheet_Streifenkontrolle.getRange("B12:K12").setValues(Array_Kontrolle);
    Sheet_Streifenkontrolle.getRange("L12:N12").setValues([[new Date(), Kalenderwoche(new Date()), LSPD.Umwandeln()]]);

    Sheet_Streifenkontrolle.getRange("B" + Zeile + ":K" + Zeile).clearContent();
    Sheet_Streifenkontrolle.getRangeList(["D" + Zeile + ":G" + Zeile, "K" + Zeile]).setValue(false).insertCheckboxes();

    Sheet_Streifenkontrolle.getRange("F" + Zeile + ":I" + Zeile).clearNote();

    Lock.releaseLock();

    var Array_Streifenkontrollen = Sheet_Streifenkontrolle.getRange("B12:L" + Sheet_Streifenkontrolle.getLastRow()).getValues();

    var Zeitgrenze = new Date();
    Zeitgrenze.setDate(Zeitgrenze.getDate() - 30);

    var Anzahl_Fehlerhafte_Kontrollen = [];

    for(var i = 0; i < Array_Streifenkontrollen.length; i++)
    {
      if(Array_Streifenkontrollen[i][0] != "")
      {
        if(Array_Streifenkontrollen[i][0] == Array_Kontrolle[0][0])
        {
          if(Array_Streifenkontrollen[i][10] >= Zeitgrenze)
          {
            if(Array_Streifenkontrollen[i][9] == false)
            {
              Anzahl_Fehlerhafte_Kontrollen.push(Array_Streifenkontrollen[i][10]);
            }
          }
        }
      }
    }

    if(Anzahl_Fehlerhafte_Kontrollen.length >= 3)
    {
      SpreadsheetApp.flush();
      var UI = SpreadsheetApp.getUi();
      UI.alert("Detective Bureau", Array_Kontrolle[0][0] + " ist binnen der letzten 30 Tage " + Anzahl_Fehlerhafte_Kontrollen.length + "x durch Streifenkontrollen durchgefallen.", UI.ButtonSet.OK);
    }
  }
}