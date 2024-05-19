function onOpen()
{
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Funktionen")      
    .addItem("Einstellung", "Einstellung")
    .addSubMenu
    (
      ui.createMenu('Personal Master')
        .addItem("Sortieren nach DN", "MasterSortDN")
        .addItem("Sortieren nach Rang", "MasterSortRang")
        .addItem("Sortieren nach Beitrittsdatum", "MasterSortBeitritt")
    ) 
    .addToUi();

    LSPD.onOpen();
}

function MasterSortDN()         // Sortieren nach Dienstnummer
{
  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal Master") 
  var sortierBereich=sheet.getRange("B6:S");
  sortierBereich.sort([{column: 2, ascending: true}, {column: 3, ascending: true}]);
}

function MasterSortRang()       // Sortieren nach Rang
{
  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal Master") 
  var sortierBereich=sheet.getRange("B6:S");  
  sortierBereich.sort([{column: 4, ascending: false}, {column: 3, ascending: true}]); 
}


function MasterSortBeitritt()     // Sortieren nach Beitritt
{
  var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Personal Master") 
  var sortierBereich=sheet.getRange("B6:S");  
  sortierBereich.sort([{column: Spalte_in_Index("I"), ascending: true}, {column: 3, ascending: true}]);  
}