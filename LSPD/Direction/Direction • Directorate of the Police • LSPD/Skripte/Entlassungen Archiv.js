function Entlassungen_Archiv(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  if(Spalte == Spalte_in_Index("S") && Zeile >= 4 && Value == "TRUE") // Berufssperre
  {
    var Sheet_Entlassung = SpreadsheetApp.getActive().getSheetByName("Entlassungen Archiv");

    Sheet_Entlassung.getRange(Zeile, Spalte).clearContent().clearDataValidations();
    Sheet_Entlassung.getRange(Zeile, Spalte).setFormula(`=HYPERLINK("https://docs.google.com/forms/d/e/1FAIpQLSenXmBoxAmHNflNy7N1J2HbQ5D2Uq7ArjjCL2Qa2Wu3-YtOyw/viewform"; "KLICK")`);

    SpreadsheetApp.flush();
    Utilities.sleep(15000);

    Sheet_Entlassung.getRange(Zeile, Spalte).setValue(true).insertCheckboxes();
  }
  else if(Spalte == Spalte_in_Index("T") && Zeile >= 4 && Value == "TRUE")  // Forum Kündigung
  {
    var Sheet_Forum = SpreadsheetApp.getActive().getSheetByName("Entlassung Forum");
    var Sheet_Entlassung = SpreadsheetApp.getActive().getSheetByName("Entlassungen Archiv");

    var Array_Entlassung = Sheet_Entlassung.getRange("B" + Zeile + ":Q" + Zeile).getValues();
    var Array_Forum = Sheet_Forum.getRange("B4:E" + Sheet_Forum.getLastRow()).getValues();


    var Name = Array_Entlassung[0][1];
    var Datum = Array_Entlassung[0][9];
    var Grund = Array_Entlassung[0][13];
    var Art = Array_Entlassung[0][14];

    for(var i = 0; i < Array_Forum.length; i++)
    {
      if(Array_Forum[i][1] == LSPD.Umwandeln())
      {
        var Rang = Array_Forum[i][0];
        var Beamter = Array_Forum[i][1];
        var Unterschrift = Array_Forum[i][3];
        break;
      }
    }

    if(Art == "Entlassung")
    {
      var HTML = 
      '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/yEN6een.png" alt="ZrllxIQ.png"></a></p>'+
      '<p>Los Santos Police Department</p>'+
      '<p>Directorate of the Police</p>'+
      '<p>Mission Row 1</p>'+
      '<p>Los Santos</p>'+
      '<p><br></p>'+
      '<p class="text-right">Los Santos, ' + Utilities.formatDate(Datum,SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy") + '</p>'+
      '<p><br></p>'+
      '<p>Sehr geehrter Herr /geehrte Frau ' + Name + ',</p>'+
      '<p><br></p>'+
      '<p>hiermit werden Sie mit sofortiger Wirkung fristlos gekündigt.</p>'+
      '<p>Begründung hierfür ist:</p>'+
      '<p><em> - ' + Grund + '</em></p>'+
      '<p><br></p>'+
      '<p>Bei weiteren Fragen können Sie uns jederzeit kontaktieren.</p>'+
      '<p><br></p>'+
      '<p>Mit freundlichen Grüßen</p>'+
      '<p><img src="' + Unterschrift + '"></p>'+
      '<p>' + Beamter + '</p>'+
      '<p>' + Rang +'</p>'+
      '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/SGJ3vnc.png" alt="SGJ3vnc.png"></a></p>';

      SpreadsheetApp.getUi().alert(HTML);
    }
    else if(Art == "Kündigung")
    {
      var HTML =
      '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/yEN6een.png" alt="ZrllxIQ.png"></a></p>'+
      '<p>Los Santos Police Department</p>'+
      '<p>Directorate of the Police</p>'+
      '<p>Mission Row 1</p>'+
      '<p>Los Santos</p>'+
      '<p><br></p>'+
      '<p class="text-right">Los Santos, ' + Utilities.formatDate(Datum,SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy") + '</p>'+
      '<p><br></p>'+
      '<p>Sehr geehrter Herr /geehrte Frau ' + Name + ',</p>'+
      '<p><br></p>'+
      '<p>hiermit bestätigen wir den Eingang ihrer selbständigen Kündigung.</p>'+
      '<p>Ihre Begründung war Folgende:</p>'+
      '<p><em> - ' + Grund + '</em></p>'+
      '<p><br></p>'+
      '<p>Bei weiteren Fragen können Sie uns jederzeit kontaktieren.</p>'+
      '<p><br></p>'+
      '<p>Mit freundlichen Grüßen</p>'+
      '<p><img src="' + Unterschrift + '"></p>'+
      '<p>' + Beamter + '</p>'+
      '<p>' + Rang +'</p>'+
      '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/SGJ3vnc.png" alt="SGJ3vnc.png"></a></p>';

      SpreadsheetApp.getUi().alert(HTML);
    }
  }
  else if(Spalte == Spalte_in_Index("V") && Zeile >= 4 && Value == "TRUE") // Wiedereinstellung
  {
    var Sheet = SpreadsheetApp.getActive().getSheetByName("Entlassungen Archiv")
    var Werte = Sheet.getRange("B" + Zeile + ":M" + Zeile).getValues();
    var Datum = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy-MM-dd");
    var Name = LSPD.Umwandeln().toString().replace(" ","+");

    var URL = "https://docs.google.com/forms/d/e/1FAIpQLSf-mazqw8fOh4irePEzW7_tD1pSq24wq5Uj7-uWr9TV3_8YxA/viewform?entry.1348382729="+Werte[0][0]+"&entry.1539975656="+Werte[0][1].replace(" ","+")+"&entry.772200757="+Werte[0][2]+"&entry.386296139="+Werte[0][3]+"+&entry.1040159266="+Werte[0][4]+"&entry.1333249206="+Werte[0][5]+"&entry.1645847867="+Werte[0][6]+"&entry.1172357223="+Datum+"&entry.1613146581="+Werte[0][10]+"&entry.1269169831="+Werte[0][11]+"&entry.846646230=&entry.1219952620=Ja&entry.1998558585="+Name;

   Sheet.getRange(Zeile,Spalte).setFormula("=HYPERLINK(\""+URL+"\";\"Link\")").setFontColor("#0000ff").removeCheckboxes();

   Utilities.sleep(10000);

   Sheet.getRange(Zeile,Spalte).setFontColor("#000000").insertCheckboxes().setValue(true);

  }
}

