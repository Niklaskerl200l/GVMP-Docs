function onOpen() 
{
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Funktionen')
    .addItem('💻 - Besprechungsprotokoll Erstellen', 'Besprechung_Start')
    .addItem('🔁 - Beschwerden Sortieren', 'Sort_Beschwerden')
    .addItem('🔎 - Beschwerde suchen', 'Suche_Beschwerde')
    .addItem('📞 - Dienstnummer suchen', 'Suche_DN')
  .addToUi();

  ui.createMenu("DB-Postservice")
    .addItem("Einladung versenden...", "GS_DB_Vorladung_Menu")
  .addToUi();

  LSPD.onOpen();

  var SS_Detective = SpreadsheetApp.getActive();
  var UI_Ausgabe = "";
  
  /**
   * Rückanfragen
   */
  var Sheet_Rueckrufe = SS_Detective.getSheetByName("Rückrufanfragen");
  var Array_Rueckrufe = Sheet_Rueckrufe.getRange("B5:I" + Sheet_Rueckrufe.getLastRow()).getValues();

  Array_Rueckrufe = Array_Rueckrufe.filter(function(e){return e[0] != "" && (e[4] == false && e[7] == false)});

  if(Array_Rueckrufe.length > 0)
  {
    Sheet_Rueckrufe.showSheet();
    UI_Ausgabe += `Es liegen ${Array_Rueckrufe.length} offene Rückrufe vor...`
  }

  /**
   * Fuhrparkmeldungen
   */
  var Sheet_Fuhrparkmeldungen = SS_Detective.getSheetByName("Fuhrparkmeldungen");
  var Array_Fuhrparkmeldungen = Sheet_Fuhrparkmeldungen.getRange("B4:I2003").getValues();

  Array_Fuhrparkmeldungen = Array_Fuhrparkmeldungen.filter(function(e){return e[0] != "" && e[0] == "In Bearbeitung"});

  if(Array_Fuhrparkmeldungen.length > 0)
  {
    if(UI_Ausgabe != "")
    {
      UI_Ausgabe += "\n"
    }

    UI_Ausgabe += `Es liegen ${Array_Fuhrparkmeldungen.length} offene Fuhrparkmeldungen vor...`;
  }

  if(UI_Ausgabe != "")
  {
    ui.alert(UI_Ausgabe.toString());
  }

  Bewerber_onOpen();
}