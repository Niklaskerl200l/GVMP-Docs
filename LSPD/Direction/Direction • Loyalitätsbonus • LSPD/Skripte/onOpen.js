function onOpen(e)
{
   SpreadsheetApp.getUi().createMenu('Funktionen')
    .addItem('HTML Code', 'HTML')
    .addItem('Sortieren', 'Sortieren_Auszahlungstabelle')
    .addToUi();

  LSPD.onOpen();
}