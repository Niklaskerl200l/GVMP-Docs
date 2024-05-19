function Bewerbungen_Quellcode(e)
{
  var Zeile = e.range.getRow();
  var Spalte = e.range.getColumn();
  var Value = e.value;

  var Sheet_Forum = SpreadsheetApp.getActive().getSheetByName("Forenverwaltung NEU");

  var Anrede = Sheet_Forum.getRange("B" + Zeile).getValue();
  var Bewerber = Sheet_Forum.getRange("C" + Zeile).getValue();
  var Datum = Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yyyy");

  var Array_Bearbeiter = Sheet_Forum.getRange("O4:Q18").getValues();
  var Array_Liste = Sheet_Forum.getRange("J4:M18").getValues();

  var Unterschrift;
  var Name;
  var Rang;
  var Liste = "";

  if(Anrede == "Herr")
  {
    Anrede = "Sehr geehrter Herr ";
  }
  else if(Anrede == "Frau")
  {
    Anrede = "Sehr geehrte Frau ";
  }
  else if(Anrede == "" || Anrede == undefined)
  {
    Anrede = "Sehr geehrter Herr ";
  }

  for(var i = 0;i < Array_Bearbeiter.length; i++)
  {
    if(Array_Bearbeiter[i][0] == LSPD.Umwandeln())
    {
      Rang = Array_Bearbeiter[i][1];
      Name = Array_Bearbeiter[i][0];
      Unterschrift = Array_Bearbeiter[i][2];
      break;
    }
  }

  for(var i = 0; i < Array_Liste.length; i++)
  {
    if(Array_Liste[i][0] != "")
    {
      Liste = Liste + '<tr>' +
        '<td>' + Array_Liste[i][1] + ' [user=\'' + Array_Liste[i][0] + '\']' + Array_Liste[i][1] + '[/user]</td>' +
        '<td>' + Array_Liste[i][2] + '</td>' +
        '<td>' + Array_Liste[i][3] + '</td>' +
      '</tr>';
    }
  }

//----------------------------------- In Evaluation ------------------------------------------------------------------//

  if(Zeile >= 4 && Zeile <= 8 && Spalte == Spalte_in_Index("D") && Value == "TRUE")
  { 
    var Code = '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/N4mEmmJ.png"></a></p>' +
    '<p>Los Santos Police Department</p>' +
    '<p>Recruitment Division</p>' +
    '<p>Mission Row 1</p>' +
    '<p>Los Santos</p>' +
    '<p class="text-right"></p>' +
    '<p><br></p>' +
    '<p class="text-right">Los Santos, ' + Datum + '</p>' +
    '<p><br></p>' +
    '<p>' + Anrede + Bewerber + ',</p>' +
    '<p><br></p>' +
    '<p>hiermit bestätigen wir den Eingang Ihrer Bewerbungsunterlagen.</p>' +
    '<p>Bitte haben Sie Verständnis dass die Evaluation Ihrer Bewerbung einige Zeit in Anspruch nehmen kann.</p>' +
    '<p>Nach Prüfung Ihrer Bewerbungsunterlagen werden wir uns mit Ihnen in Verbindung setzen.</p>' +
    '<p><br></p>' +
    '<p>Bei weiteren Fragen können Sie uns jederzeit kontaktieren.</p>' +
    '<p><br></p>' +
    '<p>Mit freundlichen Grüßen</p>' +
    '<p><img src="' + Unterschrift + '"></p>' +
    '<p>' + Name + '</p>' +
    '<p>' + Rang + '</p>' +
    '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/SGJ3vnc.png"></a></p>';

    Sheet_Forum.getRange("B"+ Zeile + ":G" + Zeile).clearContent();

    SpreadsheetApp.getUi().alert(Code);
  }

//--------------------------------------- Eingeladen --------------------------------------------------------//

  if(Zeile >= 4 && Zeile <= 8 && Spalte == Spalte_in_Index("E") && Value == "TRUE")
  {
    var Code = '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/ZrllxIQ.png"></a></p>' +
    '<p>Los Santos Police Department</p>' +
    '<p>Recruitment Division</p>' +
    '<p>Mission Row 1</p>' +
    '<p>Los Santos</p>' +
    '<p class="text-right"></p>' +
    '<p><br></p>' +
    '<p class="text-right">Los Santos, ' + Datum + '</p>' +
    '<p><br></p>' +
    '<p>' + Anrede + Bewerber + ',</p>' +
    '<p><br></p>' +
    '<p>vielen Dank für Ihre Geduld.</p>' +
    '<p><br></p>' +
    '<p>Nach gründlicher Evaluation Ihrer Bewerbungsunterlagen sind wir zum Entschluss gekommen</p>' +
    '<p>Ihnen die Möglichkeit zu geben, uns in einem Bewerbungsgespräch zu überzeugen.</p>' +
    '<p><br></p>' +
    '<p>Aufgrund der Vielzahl an Bewerbungen bitten wir Sie einen Termin mit allen der aufgeführten Beamten per E-Mail zu vereinbaren.</p>' +
    '<p>Bitte stellen Sie sicher, dass die Mitarbeiter zu den gewünschten Terminen zur Verfügung stehen.</p>' +
    '<p><br></p>' +
    '<table>' +
    '<tbody>' +
    '<tr>' +
    '<td><span style="font-size: 14pt;"><strong>VOR- UND ZUNAME</strong></span></td>' +
    '<td><span style="font-size: 14pt;"><strong>DIENSTGRAD</strong></span></td>' +
    '<td><span style="font-size: 14pt;"><strong>GESPRÄCHSZEITEN</strong></span></td>' +
    '</tr>' +
    Liste +
    '</tbody>' +
    '</table>' +
    '<p><br></p>' +
    '<p>Bei weiteren Fragen können Sie uns jederzeit kontaktieren.</p>' +
    '<p><br></p>' +
    '<p>Mit freundlichen Grüßen</p>' +
    '<p><img src="' + Unterschrift + '"></p>' +
    '<p>' + Name + '</p>' +
    '<p>' + Rang + '</p>' +
    '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/SGJ3vnc.png" alt="SGJ3vnc.png"></a></p>';

    Sheet_Forum.getRange("B"+ Zeile + ":G" + Zeile).clearContent();

    SpreadsheetApp.getUi().alert(Code);
  }

//-------------------------------------------- Angenommen --------------------------------------------------//

  if(Zeile >= 4 && Zeile <= 8 && Spalte == Spalte_in_Index("F") && Value == "TRUE")
  {
    var Zeile_Academy = SpreadsheetApp.getActive().getSheetByName("Import Bewerber").getRange("D1").getValue();
    var Array_Academy = SpreadsheetApp.getActive().getSheetByName("Import Bewerber").getRange("D2:F" + Zeile_Academy).getValues();
    var Termin_Academy;
    var Termin_Gesund;
    var Uhrzeit_Academy = SpreadsheetApp.getActive().getSheetByName("Export Auswertungsgedöns").getRange("L20").getValue();
    var Uhrzeit_Gesund = SpreadsheetApp.getActive().getSheetByName("Export Auswertungsgedöns").getRange("O20").getValue();
    var Gefunden = false;

    for(var i = 0; i < Array_Academy.length; i++)
    {
      if(Array_Academy[i][0] == Bewerber)
      {
        Termin_Academy = Array_Academy[i][1];
        Termin_Gesund = Array_Academy[i][2];
        
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)
    {
      Termin_Academy = SpreadsheetApp.getUi().prompt("Academy Termin(DD.MM.YYYYY)?").getResponseText();
      Termin_Gesund = SpreadsheetApp.getUi().prompt("Gesundheitscheck Termin(DD.MM.YYYYY)?").getResponseText();
    }    

    var Code = '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/%22%3E"><img src="https://i.imgur.com/d3DvSog.png" alt="d3DvSog.png"></a></p>' +
    '<p>Los Santos Police Department</p>' +
    '<p>Recruitment Division</p>' +
    '<p>Mission Row 1</p>' +
    '<p>Los Santos</p>' +
    '<p class="text-right"></p>' +
    '<p><br></p>' +
    '<p class="text-right">Los Santos, ' + Datum + '</p>' +
    '<p><br></p>' +
    '<p>' + Anrede + Bewerber + ',</p>' +
    '<p><br></p>' +
    '<p>vielen Dank für Ihre Geduld.</p>' +
    '<p><br></p>' +
    '<p>Es freut uns Ihnen mitteilen zu können, dass Sie uns im Bewerbungsgespräch von sich überzeugen konnten.</p>' +
    '<p>Wir laden Sie hiermit herzlich zur Rekrutierung beim Los Santos Police Department ein.</p>' +
    '<p><br></p>' +
    '<p>Hierfür sehen wir folgenden Termin vor:</p>' +
    '<p><br></p>' +
    '<p class="text-center"><strong><span style="color:#00FF00;">' + Termin_Gesund + ' - ' + Uhrzeit_Gesund + ' Gesundheitscheck <br></span></strong></p>' +
    '<p class="text-center"><strong><span style="color:#00FF00;">Und<br></span></strong></p>' +
    '<p class="text-center"><strong><span style="color:#00FF00;">' + Termin_Academy + ' - ' + Uhrzeit_Academy + ' Academy</span></strong></p>' +
    '<p><br></p>' +
    '<p>Finden Sie sich hierfür vor dem Los Santos Police Department - Mission Row ein.</p>' +
    '<p><br></p>' +
    '<p>Bringen Sie bitte folgende Gegenstände zur Rekrutierung mit:</p>' +
    '<p><br></p>' +
    '<p><strong>- Smartphone</strong></p>' +
    '<p><strong>- Funkgerät</strong></p>' +
    '<p><strong>- Laptop</strong></p>' +
    '<p><strong>- Rucksack</strong></p>' +
    '<p><strong>- festen Wohnsitz</strong></p>' +
    '<p><strong><br></strong></p>' +
    '<p>Sollten Sie terminliche Schwierigkeiten haben, melden Sie sich bitte zeitnah per E-Mail bei unserer Recruitment Division.</p>' +
    '<p><br></p>' +
    '<p>Bei weiteren Fragen können Sie uns jederzeit kontaktieren.</p>' +
    '<p></p>' +
    '<p></p>' +
    '<p><br></p>' +
    '<p>Mit freundlichen Grüßen</p>' +
    '<p><img src="' + Unterschrift + '"></p>' +
    '<p>' + Name + '</p>' +
    '<p>' + Rang + '</p>' +
    '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/%22%3E"><img src="https://i.imgur.com/SGJ3vnc.png" alt="SGJ3vnc.png"></a></p>';

    SpreadsheetApp.getUi().alert(Code);
  }

//------------------------------------------- Abgelehnt -------------------------------------------------//

  if(Zeile >= 4 && Zeile <= 8 && Spalte == Spalte_in_Index("H") && Value == "TRUE")
  {
    var Grund = Sheet_Forum.getRange("G" + Zeile).getValue();

    var Code = '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/RfZgEuG.png"></a></p>' +
    '<p></p>' +
    '<p>Los Santos Police Department</p>' +
    '<p>Recruitment Division</p>' +
    '<p>Mission Row 1</p>' +
    '<p>Los Santos</p>' +
    '<p class="text-right"></p>' +
    '<p><br></p>' +
    '<p class="text-right">Los Santos, ' + Datum + '</p>' +
    '<p><br></p>' +
    '<p>' + Anrede + Bewerber + ',</p>' +
    '<p><br></p>' +
    '<p>vielen Dank für Ihre Bewerbung und dass damit gezeigte Interesse an unserer Behörde.</p>' +
    '<p><br></p>' +
    '<p>Bedauerlicherweise müssen wir Ihnen mitteilen, dass uns Ihre Bewerbung nicht überzeugt hat.</p>' +
    '<p>Gründe hierfür sind:</p>' +
    '<p>- ' + Grund + '</p>' +
    '<p><br></p>' +
    '<p>Nach 7 Tagen besteht die Möglichkeit auf eine erneute Bewerbung bei uns.</p>' +
    '<p><br></p>' +
    '<p>Bei weiteren Fragen können sie sich jederzeit an die Recruitment Leitung wenden.</p>' +
    '<p><br></p>' +
    '<p>Mit freundlichen Grüßen</p>' +
    '<p><img src="' + Unterschrift + '"></p>' +
    '<p>' + Name + '</p>' +
    '<p>' + Rang + '</p>' +
    '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?board/30-bewerbungen/"><img src="https://i.imgur.com/SGJ3vnc.png"></a></p>';

    Sheet_Forum.getRange("B"+ Zeile + ":H" + Zeile).clearContent();

    SpreadsheetApp.getUi().alert(Code);
  }
}
