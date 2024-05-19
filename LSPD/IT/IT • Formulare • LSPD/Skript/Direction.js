var ID_Direction = LSPD.ID_Direction;

function Direction_Einstellung(e)
{
  var Sheet_Direction = SpreadsheetApp.openById(ID_Direction).getSheetByName("Personal Master");
  var Sheet_Direction_Einstellung = SpreadsheetApp.openById(ID_Direction).getSheetByName("Einstellungen Archiv");

  var Werte = e.namedValues;

  if(Werte.Wiedereinstellung == "Ja")
  {
    var Array_Wiedereinstellung = SpreadsheetApp.openById(ID_Direction).getSheetByName("Entlassungen Archiv").getRange("B4:C").getValues();

    for(var y = 0; y < Array_Wiedereinstellung.length; y++)
    {
      if(Array_Wiedereinstellung[y][1] == Werte.Name)
      {
        var Letze_DN = Array_Wiedereinstellung[y][0];
        break;
      }
    }
  }
  else
  {
    var Letze_DN = SpreadsheetApp.openById(ID_Direction).getSheetByName("Auswertungsgedöns").getRange("C7").getValue();
  }
  Logger.log(Werte);
  var Array_Eingabe = [[Letze_DN,Werte.Name,Werte.Rang,Werte.Telefonnummer,Werte.ID,Werte.GWD,Werte.ZD,Werte.Beitritt,Werte.Beitritt,"","",Werte["Forum ID"],Werte.Email,Werte.Notiz]];

  var Array_Einstellung = [[Werte.Zeitstempel,Letze_DN,Werte.Name,Werte.Rang,Werte.Telefonnummer,Werte.ID,Werte["Sozial Stufe"],Werte.GWD,Werte.ZD,Werte.Beitritt,Werte["Forum ID"],Werte.Email,Werte.Notiz,Werte.Wiedereinstellung,Werte["Ausführender"]]];
  
  var Letzte_Zeile = Sheet_Direction.getLastRow() + 1;

  Sheet_Direction.insertRowAfter(Letzte_Zeile - 1);

  Sheet_Direction.getRange("B" + Letzte_Zeile + ":O" + Letzte_Zeile).setValues(Array_Eingabe);

  Sheet_Direction.getRange("A" + Letzte_Zeile).setFormula("=C" + Letzte_Zeile);

  Sheet_Direction.getRange("K" + Letzte_Zeile).setFormula("=IF(ISNA(FILTER(Abmeldungen!$D$3:$E; Abmeldungen!$C$3:$C = $C"+Letzte_Zeile+"; Abmeldungen!$D$3:$D <= TODAY(); Abmeldungen!$E$3:$E >= TODAY()));;INDEX(SORT(FILTER(Abmeldungen!$D$3:$E; Abmeldungen!$C$3:$C = $C"+Letzte_Zeile+"; Abmeldungen!$D$3:$D <= TODAY(); Abmeldungen!$E$3:$E >= TODAY());2;FALSE);1;))");

  Sheet_Direction.getRange("P" + Letzte_Zeile + ":R" + Letzte_Zeile).insertCheckboxes();

  Sheet_Direction.getRange("B6:S" + Letzte_Zeile).sort([{column: 4, ascending: false}, {column: 3, ascending: true}]);

  Sheet_Direction_Einstellung.insertRowAfter(3);

  Sheet_Direction_Einstellung.getRange("B4:P4").setValues(Array_Einstellung);
  
  var Forum_ID = SpreadsheetApp.newRichTextValue().setText(Werte["Forum ID"]).setLinkUrl("https://www.gvmp.de/index.php?user/" + Werte["Forum ID"]).build();

  Sheet_Direction_Einstellung.getRange("L4").setRichTextValue(Forum_ID);
}

function Direction_Entlassung(e)
{
  var Sheet_Direction = SpreadsheetApp.openById(ID_Direction).getSheetByName("Personal Master");
  var Sheet_Direction_Entlassung = SpreadsheetApp.openById(ID_Direction).getSheetByName("Entlassungen Archiv");

  var Werte = e.namedValues;

  Logger.log(JSON.stringify(Werte));

  var Array_Master = Sheet_Direction.getRange("B6:O").getValues();

  for(var y = 0; y < Array_Master.length; y++)
  {
    if(Array_Master[y][1] == Werte.Name)
    {
      var Array_Entlassung = Array_Master[y];
      Logger.log(Array_Entlassung);

      Sheet_Direction.deleteRow(y + 6);
      
      break;
    }
  }

  var Array_Archiv = 
  [[
    Array_Entlassung[0],
    Array_Entlassung[1],
    Array_Entlassung[2],
    Array_Entlassung[3],
    Array_Entlassung[4],
    Array_Entlassung[5],
    Array_Entlassung[6],
    Array_Entlassung[7],
    Array_Entlassung[8],
    new Date(),
    Array_Entlassung[11],
    Array_Entlassung[12],
    Werte["Interner Grund"],
    Werte["Externer Grund"],
    Werte["Art"],
    Werte["Ausführender"],
    "",
    "",
    "",
    "",
    ""
  ]];

  var Forum_ID = SpreadsheetApp.newRichTextValue().setText(Array_Entlassung[11]).setLinkUrl("https://www.gvmp.de/index.php?user/" + Array_Entlassung[11]).build();

  Sheet_Direction_Entlassung.insertRowAfter(3);

  Sheet_Direction_Entlassung.getRange("B4:V4").setValues(Array_Archiv);

  Sheet_Direction_Entlassung.getRange("L4").setRichTextValue(Forum_ID);

  Sheet_Direction_Entlassung.getRange("R4:V4").insertCheckboxes(); 

  if(Werte["Blacklist"] == "Ja")
  {
    Sheet_Direction_Entlassung.getRange("R4").insertCheckboxes().setValue(true);
  }
}

function Direction_Telefonnummer(e)
{
  var Sheet_Direction = SpreadsheetApp.openById(ID_Direction).getSheetByName("Personal Master");

  var Werte = e.namedValues;

  var Array_Namen = Sheet_Direction.getRange("C6:C").getValues();

  for(var y = 0; y < Array_Namen.length; y++)
  {
    if(Array_Namen[y][0] == Werte.Name)
    {
      Sheet_Direction.getRange("E" + (y + 6)).setValue(Werte["Neue Telefonnummer"]);
    }
  }
}
