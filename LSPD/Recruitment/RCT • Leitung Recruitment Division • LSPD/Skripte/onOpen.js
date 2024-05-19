function onOpen()
{
  var ui = SpreadsheetApp.getUi();
  var user = Session.getTemporaryActiveUserKey();
  
  Logger.log("Benutzer: " + user);

  ui.createMenu('Funktionen')
    .addItem('Personalliste Sortieren', 'Personal_Sortieren')
    .addItem('Letzter EST suchen', 'EST_Suche')
    .addToUi();
  
  LSPD.onOpen();
}


function EST_Suche()
{
  var UI = SpreadsheetApp.getUi();
  var Confirmation = UI.prompt("EST Suche", "Geben Sie den Namen der Person an.", UI.ButtonSet.OK).getResponseText();

  if(Confirmation != undefined && Confirmation != null && Confirmation != "")
  {
    Confirmation = Confirmation.toString().replace(" ", "_");

    var SS_EST_Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_EST);
    var Sheet_Uebersicht = SS_EST_Archiv.getSheetByName("Übersicht");
    var Array_Uebersicht = Sheet_Uebersicht.getRange("B3:H").getValues();

    Array_Uebersicht = Array_Uebersicht.filter(function(e){return e[0] != "" && e[0].toString().replace(" ", "_").toUpperCase() == Confirmation.toUpperCase() && e[6].toString().includes("Nicht Bestanden") == false});

    if(Array_Uebersicht.length > 0)
    {
      Array_Uebersicht = Array_Uebersicht.sort(EST_Suche_Sort);

      return UI.alert("EST Suche", `${Array_Uebersicht[0][0]} hat seinen letzten bestandenen EST gehabt am ${Utilities.formatDate(Array_Uebersicht[0][1], "CET", "dd.MM.yyyy")} ${Array_Uebersicht[0][6]} bestanden.`, UI.ButtonSet.OK);
    }
    else
    {
      return UI.alert("EST Suche", "Diese Person hat keinen EST bisher bestanden.", UI.ButtonSet.OK);
    }
  }
}

function EST_Suche_Sort(a, b)
{
  var Offset_Name = 0;
  var Offset_Datum = 1;

  if(a[Offset_Name] != "")
  {
    if(b[Offset_Name] != "")
    {
      if(new Date(a[Offset_Datum]) < new Date(b[Offset_Datum]))
      {
        return 1;
      }
      else
      {
        return -1;
      }
    }

    if(a[Offset_Name] != "" && b[Offset_Name] == "")
    {
      return -1;
    }
  }

  return 0;
}