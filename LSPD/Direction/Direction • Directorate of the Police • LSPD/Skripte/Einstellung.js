function Einstellung()
{
  var Sheet_Bewerber = SpreadsheetApp.getActive().getSheetByName("Import Bewerber");

  var Datum = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd");
  var Name = LSPD.Umwandeln().toString().replace(" ","+");

  var Array_Bewerber = Sheet_Bewerber.getRange("B3:H" + Sheet_Bewerber.getLastRow()).getValues();

  var Bewerber = SpreadsheetApp.getUi().prompt("Name des Bewerbers (Ohne Unterstrich)").getResponseText();
  var Gefunden = false;

  if(Bewerber.includes("_"))
  {
    SpreadsheetApp.getUi().alert("Kein Unterstrich");
    return 0;
  }
  
  Logger.log("Bewerber: " + Bewerber);

  Suche_Rang_Blacklist = SpreadsheetApp.getActive().getSheetByName("Blacklist").getRange("B4:B").createTextFinder(Bewerber).findNext();

  if(Suche_Rang_Blacklist != null)
  {
    SpreadsheetApp.getUi().alert(Bewerber + " kann nicht Eingestellt werden da er einen Aktiven Blacklist eintrag hat!")

    SpreadsheetApp.getActive().getSheetByName("Blacklist").setActiveSelection(Suche_Rang_Blacklist);
    
    return 0;
  }

  for(var y = 0; y < Array_Bewerber.length; y++)
  {
    if(Array_Bewerber[y][0].toString().toUpperCase() == Bewerber.toUpperCase())
    {
      var Sheet_Wiedereinstellung = SpreadsheetApp.getActive().getSheetByName("Entlassungen Archiv");

      var Suche_Rang = Sheet_Wiedereinstellung.getRange("C4:C").createTextFinder(Bewerber).findNext();
      var Wiedereinstellung_Text = "Nein";

      if(Suche_Rang != null)
      {
        Wiedereinstellung_Text = "Ja"
      }

      var URL = "https://docs.google.com/forms/d/e/1FAIpQLSf-mazqw8fOh4irePEzW7_tD1pSq24wq5Uj7-uWr9TV3_8YxA/viewform?usp=pp_url&entry.1539975656=" + Array_Bewerber[y][0].toString().replace(" ","+") + "&entry.772200757=0&entry.386296139=" + Array_Bewerber[y][1].toString().replace(" ","+") + "&entry.1040159266=" + Array_Bewerber[y][2].toString().replace(" ","+") + "&entry.1333249206=" + Array_Bewerber[y][3].toString().replace(" ","+") + "&entry.1645847867=" + Array_Bewerber[y][4].toString().replace(" ","+") + "&entry.1613146581=" + Array_Bewerber[y][5].toString().replace(" ","+") + "&entry.1172357223=" + Datum + "&entry.1269169831=" + Array_Bewerber[y][6].toString().replace(" ","+") + "&entry.1219952620=" + Wiedereinstellung_Text + "&entry.1998558585=" + Name;

      Gefunden = true;
      break;
    }
  }

  if(Gefunden == false)
  {
    var Sheet_Wiedereinstellung = SpreadsheetApp.getActive().getSheetByName("Entlassungen Archiv");

    var Array_Entlassung = Sheet_Wiedereinstellung.getRange("B3:M").getValues();

    for(var i = 0; i < Array_Entlassung.length; i++)
    {
      if(Array_Entlassung[i][1] == Bewerber)
      {
        var Werte = [Array_Entlassung[i]];
        var Datum = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd");
        var Name = LSPD.Umwandeln().toString().replace(" ","+");

        var URL = "https://docs.google.com/forms/d/e/1FAIpQLSf-mazqw8fOh4irePEzW7_tD1pSq24wq5Uj7-uWr9TV3_8YxA/viewform?entry.1348382729="+Werte[0][0]+"&entry.1539975656="+Werte[0][1].replace(" ","+")+"&entry.772200757="+Werte[0][2]+"&entry.386296139="+Werte[0][3]+"+&entry.1040159266="+Werte[0][4]+"&entry.1333249206="+Werte[0][5]+"&entry.1645847867="+Werte[0][6]+"&entry.1172357223="+Datum+"&entry.1613146581="+Werte[0][10]+"&entry.1269169831="+Werte[0][11]+"&entry.846646230=&entry.1219952620=Ja&entry.1998558585="+Name;

        Gefunden = true;
        break;
      }
    }
  }

  if(Gefunden == false)
  {
    var URL = "https://docs.google.com/forms/d/e/1FAIpQLSf-mazqw8fOh4irePEzW7_tD1pSq24wq5Uj7-uWr9TV3_8YxA/viewform?entry.1539975656="+Bewerber.toString().replace(" ","+")+"&entry.1172357223="+Datum+"&entry.846646230=&entry.1219952620=Nein&entry.1998558585="+Name;
  }
  
  var html = "<script>window.open('" + URL + "');google.script.host.close();</script>";
  
  var userInterface = HtmlService.createHtmlOutput(html);
  
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Open Tab');
}


