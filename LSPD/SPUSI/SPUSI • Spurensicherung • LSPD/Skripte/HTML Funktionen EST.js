function EST_Start()  //Starten des Benutzerinterface für EST Eingaben
{
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('EST UI').setWidth(1114).setHeight(200).setSandboxMode(HtmlService.SandboxMode.IFRAME),"Prüfungs Daten eingeben");
}

function RCT_Alle()   //Ausgabe aller RCT für EST Benutzerinterface
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Startseite");
  var Array = Sheet.getRange("P6:P" + Sheet.getRange("P3").getValue()).getValues();
  return Array;
}

function Pruefer_Alle()   //Ausgabe aller Prüfer für EST Benutzerinterface
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Startseite");
  var Array = Sheet.getRange("B6:B" + Sheet.getRange("B3").getValue()).getValues();
  Array.push(["Marvin Hooker"])
  return Array;
}

function Email()    //Ausgabe des Namen des Aktiven Benutzers für EST Benutzerinterface
{
  return LSPD.Umwandeln();
}

function Academy_Bewerber()   //Ausgabe des Bewerbers für Acedemy Benutzerinterface
{
  var Sheet = SpreadsheetApp.getActiveSheet();
  var Name = Sheet.getRange("B5").getValue();
  return Name;
}

function Academy_Datum()   //Ausgabe Academy Termine für Acedemy Benutzerinterface
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
  var Array = Sheet.getRange("L15:L19").getValues();
  return Array;
}

function Academy_Start()  //Starten des Benutzerinterface für Academy Eingaben
{
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Academy').setWidth(500).setHeight(200).setSandboxMode(HtmlService.SandboxMode.IFRAME),"Academy Datum Auswählen");
}

function Academy_Eintragen(Datum = "07.04.2022",Bewerber = "TestH")    //Erstellen und eintragen in Academy
{
  Logger.log("test" + " " + Datum + " " + Bewerber)
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet_Bewerbung = SpreadsheetApp.getActiveSheet();
  var Sheet_LSMC = SpreadsheetApp.openById("1jXAlBnena43fRmdxWwJemh4NNBQ5YASAOdpFl37SAF4").getSheetByName("Liste");
  var Letzte_Zeile = Sheet_LSMC.getRange("B3").getValue();

  if(SS.getSheetByName("Academy " + Datum) == null)
  {
    var Sheet_Vorlage = SpreadsheetApp.getActive().getSheetByName("Academy Vorlage");

    Sheet_Vorlage.copyTo(SS).setName("Academy " + Datum);

    SpreadsheetApp.getActive().getSheetByName("Academy " + Datum).getRange("B2").setValue("Recruitment Division Academy " + Datum);

    Sheets.Spreadsheets.Values.update({ 
    "values": [["=SORT(FILTER({{IFERROR(INDIRECT(\"Academy \" & TEXT('Auswertungsgedöns'!L15 - 7;\"dd.MM.yyyy\") & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L15 & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L16 & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L17 & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L18 & \"!F5:F\");\"\")} \\ {IFERROR(INDIRECT(\"Academy \" & TEXT('Auswertungsgedöns'!L15 - 7;\"dd.MM.yyyy\") & \"!E5:E\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L15 & \"!E5:E\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L16 & \"!E5:E\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L17 & \"!E5:E\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L18 & \"!E5:E\");\"\")} \\ {IFERROR(INDIRECT(\"Academy \" & TEXT('Auswertungsgedöns'!L15 - 7;\"dd.MM.yyyy\") & \"!N5:N\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L15 & \"!N5:N\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L16 & \"!N5:N\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L17 & \"!N5:N\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L18 & \"!N5:N\");\"\")} \\ {IFERROR(INDIRECT(\"Academy \" & TEXT('Auswertungsgedöns'!L15 - 7;\"dd.MM.yyyy\") & \"!M5:M\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L15 & \"!M5:M\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L16 & \"!M5:M\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L17 & \"!M5:M\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L18 & \"!M5:M\");\"\")}}; {IFERROR(INDIRECT(\"Academy \" & TEXT('Auswertungsgedöns'!L15 - 7;\"dd.MM.yyyy\") & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L15 & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L16 & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L17 & \"!F5:F\");\"\") ; IFERROR(INDIRECT(\"Academy \" & 'Auswertungsgedöns'!L18 & \"!F5:F\");\"\")} <> \"\");1;TRUE)"]]},"1dwowG4PVZZNdllgZtKDbG9i9Qr-wrw4qdrE3e-_VkPs","Export Bewerber!D2",{"valueInputOption": "USER_ENTERED"})
    }

  var Sheet_Academy = SpreadsheetApp.getActive().getSheetByName("Academy " + Datum);

  Sheet_Academy.getRange("F5:L").sort(6);

  var Zeile_Academy = Sheet_Academy.getRange("F3").getValue();

  Sheet_Academy.getRange("F" + Zeile_Academy).setValue(Bewerber);
  Sheet_LSMC.getRange("B" + Letzte_Zeile).setValue(Bewerber);

  Sheet_Academy.setActiveSelection("F" + Zeile_Academy);

  SS.deleteSheet(Sheet_Bewerbung);
}
