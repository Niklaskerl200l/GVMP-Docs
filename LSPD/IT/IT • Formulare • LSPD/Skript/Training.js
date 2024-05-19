var ID_Training = LSPD.ID_Training;

function Training_Einstellung(e)
{
  var Werte = e.namedValues;

  var Name = Werte.Name;

  Logger.log("Einstellung Training " + Name);

  var Folder_Personalakten = DriveApp.getFolderById("1PRzIqgBH_fK1Dpffs1ZRjM7u0FoaFrbl");
  var File_Vorlage = DriveApp.getFileById("1eSxvhzib1P9YiLpcHbQYtg9ZzvWF5gw34d9T3DYvhlY");

  var Sheet_Training = SpreadsheetApp.openById(ID_Training).getSheetByName("Ausbildungsblatt");

  var New_ID = File_Vorlage.makeCopy("Ausbildungen " + Name,Folder_Personalakten).getId();

  var Sheet_Neu = SpreadsheetApp.openById(New_ID);

  var Schutz = Sheet_Neu.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  for(var i = 0; i < Schutz.length; i++)
  {
    if(Schutz[i].getDescription().search("Generalschutz") != undefined && Schutz[i].canEdit())
    {
      Schutz[i].addEditors(["gvmp.lspd.bot@gmail.com","niklaskerl2001@gmail.com"]);
    }
  }

  Sheet_Neu.getSheetByName("Übersicht").getRange("B3").setValue("Soz: " + Werte["Sozial Stufe"] + (Werte.Wiedereinstellung == "Ja" ? " (Wiedereinstellung)" : ""));

  var Letzt_Zeile = Sheet_Training.getLastRow() + 1;

  Sheet_Training.insertRowAfter(Letzt_Zeile - 1);

  var Text_Stile = SpreadsheetApp.newTextStyle().setForegroundColor("#f3f3f3").setUnderline(false).build();

   Sheet_Training.getRange("A" + Letzt_Zeile).setValue('=IF($B'+Letzt_Zeile+'="";"";IF(ISTEXT(VLOOKUP($B'+Letzt_Zeile+';\'Import Stempeluhr\'!$D$3:$D;1;FALSE));"🟢";IF(VLOOKUP($B'+Letzt_Zeile+';\'Import Personaltabelle\'!$D$4:$I;1;FALSE) = TRUE;"⚪";"🔴")))');

  Sheet_Training.getRange("B" + Letzt_Zeile).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(Name).setLinkUrl("https://docs.google.com/spreadsheets/d/"+New_ID).setTextStyle(Text_Stile).build());

  Sheet_Training.getRange("C" + Letzt_Zeile + ":E" + Letzt_Zeile).setFormulas([['=IF($B'+Letzt_Zeile+'="";"";VLOOKUP($B'+Letzt_Zeile+';\'Import Personaltabelle\'!$A$4:$I;2;FALSE))','=IF($B'+Letzt_Zeile+'="";"";VLOOKUP($B'+Letzt_Zeile+';\'Import Personaltabelle\'!$A$4:$I;9;FALSE))','=IMPORTRANGE("'+New_ID+'";"Auswertungsgedöns!B14:ES14")']]);

  Sheet_Training.getRange("F" + Letzt_Zeile + ":V" + Letzt_Zeile).setValue("");

  Sheet_Training.getRange("EW" + Letzt_Zeile + ":EY" + Letzt_Zeile).setFormulas([['=IF(ISERR($E'+Letzt_Zeile+');"";IF($B'+Letzt_Zeile+'="";"";IFERROR(JOIN(CHAR(10);QUERY(IF($C'+Letzt_Zeile+'=1;{IFERROR(IF($H'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($X$3:$AM$3);TRANSPOSE($X$3:$AM$3) <> "Ausbilder";TRANSPOSE($X$3:$AM$3) <> "";TRANSPOSE($X'+Letzt_Zeile+':$AM'+Letzt_Zeile+') = "");"");"");IFERROR(IF($I'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($AN$3:$BC$3);TRANSPOSE($AN$3:$BC$3) <> "Ausbilder";TRANSPOSE($AN$3:$BC$3) <> "";TRANSPOSE($AN'+Letzt_Zeile+':$BC'+Letzt_Zeile+') = ""););"");IFERROR(IF($J'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BD$3:$BS$3);TRANSPOSE($BD$3:$BS$3) <> "Ausbilder";TRANSPOSE($BD$3:$BS$3) <> "";TRANSPOSE($BD'+Letzt_Zeile+':$BS'+Letzt_Zeile+') = "");"");"")};IF($C'+Letzt_Zeile+'=2;{IFERROR(IF($H'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($X$3:$AM$3);TRANSPOSE($X$3:$AM$3) <> "Ausbilder";TRANSPOSE($X$3:$AM$3) <> "";TRANSPOSE($X'+Letzt_Zeile+':$AM'+Letzt_Zeile+') = "");"");"");IFERROR(IF($I'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($AN$3:$BC$3);TRANSPOSE($AN$3:$BC$3) <> "Ausbilder";TRANSPOSE($AN$3:$BC$3) <> "";TRANSPOSE($AN'+Letzt_Zeile+':$BC'+Letzt_Zeile+') = "");"");"");IFERROR(IF($J'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BD$3:$BS$3);TRANSPOSE($BD$3:$BS$3) <> "Ausbilder";TRANSPOSE($BD$3:$BS$3) <> "";TRANSPOSE($BD'+Letzt_Zeile+':$BS'+Letzt_Zeile+') = "");"");"");IFERROR(IF($K'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BT$3:$CI$3);TRANSPOSE($BT$3:$CI$3) <> "Ausbilder";TRANSPOSE($BT$3:$CI$3) <> "";TRANSPOSE($BT'+Letzt_Zeile+':$CI'+Letzt_Zeile+') = "");"");"");IFERROR(IF($L'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CJ$3:$CY$3);TRANSPOSE($CJ$3:$CY$3) <> "Ausbilder";TRANSPOSE($CJ$3:$CY$3) <> "";TRANSPOSE($CJ'+Letzt_Zeile+':$CY'+Letzt_Zeile+') = "");"");"");IFERROR(IF($M'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CZ$3:$DN$3);TRANSPOSE($CZ$3:$DN$3) <> "Ausbilder";TRANSPOSE($CZ$3:$DN$3) <> "";TRANSPOSE($CZ'+Letzt_Zeile+':$DN'+Letzt_Zeile+') = "");"");"")};IF($C'+Letzt_Zeile+' >= 3;{IFERROR(IF($H'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($X$3:$AM$3);TRANSPOSE($X$3:$AM$3) <> "Ausbilder";TRANSPOSE($X$3:$AM$3) <> "";TRANSPOSE($X'+Letzt_Zeile+':$AM'+Letzt_Zeile+') = "");"");"");IFERROR(IF($I'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($AN$3:$BC$3);TRANSPOSE($AN$3:$BC$3) <> "Ausbilder";TRANSPOSE($AN$3:$BC$3) <> "";TRANSPOSE($AN'+Letzt_Zeile+':$BC'+Letzt_Zeile+') = "");"");"");IFERROR(IF($J'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BD$3:$BS$3);TRANSPOSE($BD$3:$BS$3) <> "Ausbilder";TRANSPOSE($BD$3:$BS$3) <> "";TRANSPOSE($BD'+Letzt_Zeile+':$BS'+Letzt_Zeile+') = "");"");"");IFERROR(IF($K'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BT$3:$CI$3);TRANSPOSE($BT$3:$CI$3) <> "Ausbilder";TRANSPOSE($BT$3:$CI$3) <> "";TRANSPOSE($BT'+Letzt_Zeile+':$CI'+Letzt_Zeile+') = "");"");"");IFERROR(IF($L'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CJ$3:$CY$3);TRANSPOSE($CJ$3:$CY$3) <> "Ausbilder";TRANSPOSE($CJ$3:$CY$3) <> "";TRANSPOSE($CJ'+Letzt_Zeile+':$CY'+Letzt_Zeile+') = "");"");"");IFERROR(IF($M'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CZ$3:$DN$3);TRANSPOSE($CZ$3:$DN$3) <> "Ausbilder";TRANSPOSE($CZ$3:$DN$3) <> "";TRANSPOSE($CZ'+Letzt_Zeile+':$DN'+Letzt_Zeile+') = "");"");"")};"")));"Select * where Col1 is not null";0)))))','=IF(EW'+Letzt_Zeile+'="";"0";COUNTA(SPLIT(EW'+Letzt_Zeile+';CHAR(10);1;0)))','=IF($B'+Letzt_Zeile+'="";"";MAX(FILTER($X'+Letzt_Zeile+':$EU'+Letzt_Zeile+';MOD(COLUMN($X'+Letzt_Zeile+':$EU'+Letzt_Zeile+');2) = 0)))']]);

  Sheet_Training.getRange("B4:EY").sort([{column: 3, ascending: false},{column: 2, ascending: true}]);

  if(Werte.Wiedereinstellung == "Ja")
  {
    var Sheet_Training_Wiedereinstellung = SpreadsheetApp.openById(LSPD.ID_Leitung_Training).getSheetByName("Wiedereinstellungen");
    var Sheet_Direction_Entlassung = SpreadsheetApp.openById(ID_Direction).getSheetByName("Entlassungen Archiv");

    var Array_Ausbildungen_Datum;
    var Array_Ausbildungen_Namen;

    var Letzte_Zeile_W = Sheet_Training_Wiedereinstellung.getLastRow() + 1;

    var Array_Direction_Namen = Sheet_Direction_Entlassung.getRange("C4:C").getValues();

    for(var y = 0; y < Array_Direction_Namen.length; y++)
    {
      if(Array_Direction_Namen[y][0] == Name)
      {
        Array_Ausbildungen_Datum = Sheet_Direction_Entlassung.getRange(y + 4,Spalte_in_Index("X"),1,Sheet_Direction_Entlassung.getLastColumn() - 21).getValues();
        Array_Ausbildungen_Namen = Sheet_Direction_Entlassung.getRange(y + 4,Spalte_in_Index("X"),1,Sheet_Direction_Entlassung.getLastColumn() - 21).getNotes();

        Sheet_Training_Wiedereinstellung.insertRowAfter(Letzte_Zeile_W - 1);
        Sheet_Training_Wiedereinstellung.getRange("B" + Letzte_Zeile_W).setValue(Name);
        Sheet_Training_Wiedereinstellung.getRange("C" + Letzte_Zeile_W).setFormula("=IF(B"+Letzte_Zeile_W+"=\"\";\"\";VLOOKUP(B"+Letzte_Zeile_W+";'Import Personaltabelle'!$A$4:$B;2;FALSE))");

        Sheet_Training_Wiedereinstellung.getRange(Letzte_Zeile_W, 4, 1,Array_Ausbildungen_Datum[0].length).setValues(Array_Ausbildungen_Datum);
        Sheet_Training_Wiedereinstellung.getRange(Letzte_Zeile_W, 4, 1,Array_Ausbildungen_Namen[0].length).setNotes(Array_Ausbildungen_Namen);

        break;
      }
    }
  }
}

function Training_Entlassung(e)
{
  var Werte = e.namedValues;

  var Name = Werte.Name;
  
  var Sheet_Training = SpreadsheetApp.openById(ID_Training).getSheetByName("Ausbildungsblatt");
  var Sheet_Direction_Entlassung = SpreadsheetApp.openById(ID_Direction).getSheetByName("Entlassungen Archiv");

  var Array_Ausbildungen_Namen = Sheet_Training.getRange("E3:V3").getValues();
  var Array_Ausbildungen_Daten;
  var ID;

  var Array_Name = Sheet_Training.getRange("B4:V" + Sheet_Training.getLastRow()).getValues();

  var Array_Direction_Namen = Sheet_Direction_Entlassung.getRange("C4:C").getValues();

  for(var y = 0; y < Array_Name.length; y++)
  {
    if(Array_Name[y][0] == Name)
    {
      var Formel = Sheet_Training.getRange("E" + (y + 4)).getFormula();

      ID = Formel.substring(Formel.indexOf("\"") + 1,Formel.indexOf("\"",Formel.indexOf("\"") + 1));

      Array_Ausbildungen_Daten = Array_Name[y];

      Sheet_Training.deleteRow(y + 4)

      Array_Ausbildungen_Daten.shift();
      Array_Ausbildungen_Daten.shift();
      Array_Ausbildungen_Daten.shift();

      break;
    }
  }

  var Datum = Utilities.formatDate(new Date(),"GTM+2","dd-MM-yyyy");

  var URL = DriveApp.getFileById(ID).moveTo( DriveApp.getFolderById("1KbFnGpZ62lcb0vcgBSTwLvN8MmepWcnw")).setName(SpreadsheetApp.openById(ID).getName() + " " + Datum + " Archiv").getUrl();

  Array_Ausbildungen_Daten[0] = URL;

  for(var y = 0; y < Array_Direction_Namen.length; y++)
  {
    if(Array_Direction_Namen[y][0] == Name)
    {
      Sheet_Direction_Entlassung.getRange(y + 4,Spalte_in_Index("X"),1,Array_Ausbildungen_Daten.length).setValues([Array_Ausbildungen_Daten]);
      Sheet_Direction_Entlassung.getRange(y + 4,Spalte_in_Index("X"),1,Array_Ausbildungen_Namen[0].length).setNotes(Array_Ausbildungen_Namen);

      break;
    }
  }
}

function Training_Manuell(Name = "Karl Dickens-Sobiak",Wiedereinstellung  = "Ja")
{

  Logger.log("Einstellung Training " + Name);

  var Folder_Personalakten = DriveApp.getFolderById("1PRzIqgBH_fK1Dpffs1ZRjM7u0FoaFrbl");
  var File_Vorlage = DriveApp.getFileById("1eSxvhzib1P9YiLpcHbQYtg9ZzvWF5gw34d9T3DYvhlY");

  var Sheet_Training = SpreadsheetApp.openById(ID_Training).getSheetByName("Ausbildungsblatt");

  var New_ID = File_Vorlage.makeCopy("Ausbildungen " + Name,Folder_Personalakten).getId();

  var Sheet_Neu = SpreadsheetApp.openById(New_ID);

  var Schutz = Sheet_Neu.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  for(var i = 0; i < Schutz.length; i++)
  {
    if(Schutz[i].getDescription().search("Generalschutz") != undefined && Schutz[i].canEdit())
    {
      Schutz[i].addEditors(["gvmp.lspd.bot@gmail.com","niklaskerl2001@gmail.com"]);
    }
  }

  var Letzt_Zeile = Sheet_Training.getLastRow() + 1;

  Sheet_Training.insertRowAfter(Letzt_Zeile - 1);

  var Text_Stile = SpreadsheetApp.newTextStyle().setForegroundColor("#f3f3f3").setUnderline(false).build();

  Sheet_Training.getRange("B" + Letzt_Zeile).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(Name).setLinkUrl("https://docs.google.com/spreadsheets/d/"+New_ID).setTextStyle(Text_Stile).build());

  Sheet_Training.getRange("C" + Letzt_Zeile + ":E" + Letzt_Zeile).setFormulas([['=IF($B'+Letzt_Zeile+'="";"";VLOOKUP($B'+Letzt_Zeile+';\'Import Personaltabelle\'!$A$4:$I;2;FALSE))','=IF($B'+Letzt_Zeile+'="";"";VLOOKUP($B'+Letzt_Zeile+';\'Import Personaltabelle\'!$A$4:$I;9;FALSE))','=IMPORTRANGE("'+New_ID+'";"Auswertungsgedöns!B14:ES14")']]);

  Sheet_Training.getRange("F" + Letzt_Zeile + ":V" + Letzt_Zeile).setValue("");

  Sheet_Training.getRange("EW" + Letzt_Zeile + ":EY" + Letzt_Zeile).setFormulas([['=IF(ISERR($E'+Letzt_Zeile+');"";IF($B'+Letzt_Zeile+'="";"";IFERROR(JOIN(CHAR(10);QUERY(IF($C'+Letzt_Zeile+'=1;{IFERROR(IF($H'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($X$3:$AM$3);TRANSPOSE($X$3:$AM$3) <> "Ausbilder";TRANSPOSE($X$3:$AM$3) <> "";TRANSPOSE($X'+Letzt_Zeile+':$AM'+Letzt_Zeile+') = "");"");"");IFERROR(IF($I'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($AN$3:$BC$3);TRANSPOSE($AN$3:$BC$3) <> "Ausbilder";TRANSPOSE($AN$3:$BC$3) <> "";TRANSPOSE($AN'+Letzt_Zeile+':$BC'+Letzt_Zeile+') = ""););"");IFERROR(IF($J'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BD$3:$BS$3);TRANSPOSE($BD$3:$BS$3) <> "Ausbilder";TRANSPOSE($BD$3:$BS$3) <> "";TRANSPOSE($BD'+Letzt_Zeile+':$BS'+Letzt_Zeile+') = "");"");"")};IF($C'+Letzt_Zeile+'=2;{IFERROR(IF($H'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($X$3:$AM$3);TRANSPOSE($X$3:$AM$3) <> "Ausbilder";TRANSPOSE($X$3:$AM$3) <> "";TRANSPOSE($X'+Letzt_Zeile+':$AM'+Letzt_Zeile+') = "");"");"");IFERROR(IF($I'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($AN$3:$BC$3);TRANSPOSE($AN$3:$BC$3) <> "Ausbilder";TRANSPOSE($AN$3:$BC$3) <> "";TRANSPOSE($AN'+Letzt_Zeile+':$BC'+Letzt_Zeile+') = "");"");"");IFERROR(IF($J'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BD$3:$BS$3);TRANSPOSE($BD$3:$BS$3) <> "Ausbilder";TRANSPOSE($BD$3:$BS$3) <> "";TRANSPOSE($BD'+Letzt_Zeile+':$BS'+Letzt_Zeile+') = "");"");"");IFERROR(IF($K'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BT$3:$CI$3);TRANSPOSE($BT$3:$CI$3) <> "Ausbilder";TRANSPOSE($BT$3:$CI$3) <> "";TRANSPOSE($BT'+Letzt_Zeile+':$CI'+Letzt_Zeile+') = "");"");"");IFERROR(IF($L'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CJ$3:$CY$3);TRANSPOSE($CJ$3:$CY$3) <> "Ausbilder";TRANSPOSE($CJ$3:$CY$3) <> "";TRANSPOSE($CJ'+Letzt_Zeile+':$CY'+Letzt_Zeile+') = "");"");"");IFERROR(IF($M'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CZ$3:$DN$3);TRANSPOSE($CZ$3:$DN$3) <> "Ausbilder";TRANSPOSE($CZ$3:$DN$3) <> "";TRANSPOSE($CZ'+Letzt_Zeile+':$DN'+Letzt_Zeile+') = "");"");"")};IF($C'+Letzt_Zeile+' >= 3;{IFERROR(IF($H'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($X$3:$AM$3);TRANSPOSE($X$3:$AM$3) <> "Ausbilder";TRANSPOSE($X$3:$AM$3) <> "";TRANSPOSE($X'+Letzt_Zeile+':$AM'+Letzt_Zeile+') = "");"");"");IFERROR(IF($I'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($AN$3:$BC$3);TRANSPOSE($AN$3:$BC$3) <> "Ausbilder";TRANSPOSE($AN$3:$BC$3) <> "";TRANSPOSE($AN'+Letzt_Zeile+':$BC'+Letzt_Zeile+') = "");"");"");IFERROR(IF($J'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BD$3:$BS$3);TRANSPOSE($BD$3:$BS$3) <> "Ausbilder";TRANSPOSE($BD$3:$BS$3) <> "";TRANSPOSE($BD'+Letzt_Zeile+':$BS'+Letzt_Zeile+') = "");"");"");IFERROR(IF($K'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($BT$3:$CI$3);TRANSPOSE($BT$3:$CI$3) <> "Ausbilder";TRANSPOSE($BT$3:$CI$3) <> "";TRANSPOSE($BT'+Letzt_Zeile+':$CI'+Letzt_Zeile+') = "");"");"");IFERROR(IF($L'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CJ$3:$CY$3);TRANSPOSE($CJ$3:$CY$3) <> "Ausbilder";TRANSPOSE($CJ$3:$CY$3) <> "";TRANSPOSE($CJ'+Letzt_Zeile+':$CY'+Letzt_Zeile+') = "");"");"");IFERROR(IF($M'+Letzt_Zeile+' <> TRUE;FILTER(TRANSPOSE($CZ$3:$DN$3);TRANSPOSE($CZ$3:$DN$3) <> "Ausbilder";TRANSPOSE($CZ$3:$DN$3) <> "";TRANSPOSE($CZ'+Letzt_Zeile+':$DN'+Letzt_Zeile+') = "");"");"")};"")));"Select * where Col1 is not null";0)))))','=IF(EW'+Letzt_Zeile+'="";"0";COUNTA(SPLIT(EW'+Letzt_Zeile+';CHAR(10);1;0)))','=IF($B'+Letzt_Zeile+'="";"";MAX(FILTER($X'+Letzt_Zeile+':$EU'+Letzt_Zeile+';MOD(COLUMN($X'+Letzt_Zeile+':$EU'+Letzt_Zeile+');2) = 0)))']]);

  Sheet_Training.getRange("B4:EY").sort([{column: 3, ascending: false},{column: 2, ascending: true}]);

  if(Wiedereinstellung == "Ja")
  {
    var Sheet_Training_Wiedereinstellung = SpreadsheetApp.openById(ID_Training).getSheetByName("Wiedereinstellungen");
    var Sheet_Direction_Entlassung = SpreadsheetApp.openById(ID_Direction).getSheetByName("Entlassungen Archiv");

    var Array_Ausbildungen_Datum ;
    var Array_Ausbildungen_Namen;

    var Letzte_Zeile_W = Sheet_Training_Wiedereinstellung.getLastRow() + 1;

    var Array_Direction_Namen = Sheet_Direction_Entlassung.getRange("C4:C").getValues();

    for(var y = 0; y < Array_Direction_Namen.length; y++)
    {
      if(Array_Direction_Namen[y][0] == Name)
      {
        Array_Ausbildungen_Datum = Sheet_Direction_Entlassung.getRange(y + 4,Spalte_in_Index("X"),1,Sheet_Direction_Entlassung.getLastColumn() - 21).getValues();
        Array_Ausbildungen_Namen = Sheet_Direction_Entlassung.getRange(y + 4,Spalte_in_Index("X"),1,Sheet_Direction_Entlassung.getLastColumn() - 21).getNotes();

        Sheet_Training_Wiedereinstellung.insertRowAfter(Letzte_Zeile_W - 1);
        Sheet_Training_Wiedereinstellung.getRange("B" + Letzte_Zeile_W).setValue(Name);
        Sheet_Training_Wiedereinstellung.getRange("C" + Letzte_Zeile_W).setFormula("=IF(B"+Letzte_Zeile_W+"=\"\";\"\";VLOOKUP(B"+Letzte_Zeile_W+";'Import Personaltabelle'!$A$4:$B;2;FALSE))");

        Sheet_Training_Wiedereinstellung.getRange(Letzte_Zeile_W, 4, 1,Array_Ausbildungen_Datum[0].length).setValues(Array_Ausbildungen_Datum);
        Sheet_Training_Wiedereinstellung.getRange(Letzte_Zeile_W, 4, 1,Array_Ausbildungen_Namen[0].length).setNotes(Array_Ausbildungen_Namen);

        break;
      }
    }
  }
}