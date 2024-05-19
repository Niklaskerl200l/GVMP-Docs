function onOpen(e)
{
  LSPD.onOpen();

  var UI = SpreadsheetApp.getUi();
  
  UI.createMenu("Funktionen")
    .addItem("Rückleitung verwalten...", "Rueckleitung_Offene_Kontrollen")
  .addToUi();

  if(PropertiesService.getUserProperties().getProperty("rueckleitung") == null)
  {
    PropertiesService.getUserProperties().setProperty("rueckleitung", false);
  }
}

function Rueckleitung_Offene_Kontrollen()
{
  var UI = SpreadsheetApp.getUi();
  var Confirmation = UI.alert("LSPD", "Hallo! Möchten Sie nach einer eingetragenen Kontrolle zurückgeleitet werden zu den offenen Kontrollen?", UI.ButtonSet.YES_NO_CANCEL);

  if(Confirmation == UI.Button.YES)
  {
    PropertiesService.getUserProperties().setProperty("rueckleitung", true);
  }
  else if(Confirmation == UI.Button.NO)
  {
    PropertiesService.getUserProperties().setProperty("rueckleitung", false);
  }
}