function GS_DB_Vorladung_Menu(Benutzer = LSPD.Umwandeln(false, false))
{
  var UI = SpreadsheetApp.getUi();
  UI.showModalDialog(HtmlService.createHtmlOutputFromFile("HTML DB Vorladung").setWidth(500).setHeight(350).setSandboxMode(HtmlService.SandboxMode.IFRAME), "Detective Bureau · Vorladung");
}

function GS_DB_Vorladung_SetDaten()
{
  var Sheet = SpreadsheetApp.getActive().getSheetByName("Import Foren-IDs");
  var Daten = Sheet.getRange("B3:B202").getValues().filter(function(e){return e[0] != ""});

  return Daten;
}

function GS_DB_Vorladung_Finish(Anrede, TVO, Datum, Uhrzeit, Dauer, Art)
{
  if(!Anrede || !TVO || !Datum || !Uhrzeit || !Dauer || !Art)
  {
    return;
  }

  var Sheet = SpreadsheetApp.getActive().getSheetByName("Import Foren-IDs");
  var Daten = Sheet.getRange("B3:C202").getValues().filter(function(e){return e[0] != ""});

  var Gefunden = false;
  for(var i = 0; i < Daten.length; i++)
  {
    if(Daten[i][0] == TVO)
    {
      Gefunden = true;
      break;
    }
  }

  if(Gefunden)
  {
    Datum = new Date(Datum);
    Datum = Utilities.formatDate(Datum, "CET", "dd.MM.yyyy");

    var ForumID = Daten[i][1];

    var Unterschriften = SpreadsheetApp.getActive().getSheetByName("Unterschriften").getRange("B3:E27").getValues().filter(function(e){return e[0] != ""});
    var Gefunden_Unterschrift = false;

    var Benutzer = LSPD.Umwandeln(false, false);

    for(var i = 0; i < Unterschriften.length; i++)
    {
      if(Unterschriften[i][0] == Benutzer)
      {
        Gefunden_Unterschrift = true;
        break;
      }
    }

    var Unterschrift_Unterschrift = Benutzer;
    var Unterschrift_Detective = "Detective des Los Santos Police Department";

    if(Gefunden_Unterschrift)
    {
      if(Unterschriften[i][3] != "")
      {
        Unterschrift_Unterschrift = '<img src="' + Unterschriften[i][3] + '">';
      }
      
      Unterschrift_Detective = Unterschriften[i][2] + " & Detective des Los Santos Police Department";
    }

    var HTML;
    if(Art == "Vorladung")
    {
      HTML =
      `
      <p class="text-center"><img src="https://i.imgur.com/yEN6een.png"></p>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td class="text-center"><strong><u>Vorladung</u></strong></td>
          </tr>
        </tbody>
      </table>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td>Sehr ${Anrede} ${TVO},<br>hiermit laden wir Sie zu einer Vernehmung vor.<br><br>Die Vernehmung ist für den ${Datum}, um ${Uhrzeit} Uhr im Detective Bureau, Mission Row Police Station an der folgender Adresse angesetzt:<br><br>Los Santos Police Department<br>Mission Row 1<br>Los Santos, San Andreas<br><br>Wir bitten Sie höflichst, pünktlich zur erscheinen. Die Vernehmung wird voraussichtlich ${Dauer} (Std.) dauern. Bitte planen Sie dementsprechend ausreichend Zeit ein. Sollten Sie verhindert sein, bitten wir Sie umgehend um eine Benachrichtigung. Wir sind bemüht eine alternative Terminvereinbarung zu treffen.</td>
          </tr>
        </tbody>
      </table>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td>Mit freundlichen Grüßen<br><em>${Unterschrift_Unterschrift}<br></em>(${Unterschrift_Detective})</td>
          </tr>
        </tbody>
      </table> 
      `
    }
    else if(Art == "Personalgespräch")
    {
      HTML = 
      `
      <p class="text-center"><img src="https://i.imgur.com/yEN6een.png"></p>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td class="text-center"><strong><u>Vorladung (Personalgespräch)</u></strong></td>
          </tr>
        </tbody>
      </table>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td>Sehr ${Anrede} ${TVO},<br>hiermit laden wir Sie zu einem Personalgespräch vor.<br><br>Das Personalgespräch ist für den ${Datum}, um ${Uhrzeit} Uhr im Detective Bureau, Mission Row Police Station an der folgender Adresse angesetzt:<br><br>Los Santos Police Department<br>Mission Row 1<br>Los Santos, San Andreas<br><br>Wir bitten Sie höflichst, pünktlich zur erscheinen. Das Gespräch wird voraussichtlich ${Dauer} (Std.) dauern. Bitte planen Sie dementsprechend ausreichend Zeit ein. Sollten Sie verhindert sein, bitten wir Sie umgehend um eine Benachrichtigung. Wir sind bemüht eine alternative Terminvereinbarung zu treffen.</td>
          </tr>
        </tbody>
      </table>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td>Mit freundlichen Grüßen<br><em>${Unterschrift_Unterschrift}<br></em>(${Unterschrift_Detective})</td>
          </tr>
        </tbody>
      </table>  
      `
    }
    else if(Art == "Zeugenaussage")
    {
      HTML =
      `
      <p class="text-center"><img src="https://i.imgur.com/yEN6een.png"></p>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td class="text-center"><strong><u>Vorladung (Zeugenaussage)</u></strong></td>
          </tr>
        </tbody>
      </table>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td>Sehr ${Anrede} ${TVO},<br>hiermit laden wir Sie zu einer Zeugenaussage vor.<br><br>Die Aussage ist für den ${Datum}, um ${Uhrzeit} Uhr im Detective Bureau, Mission Row Police Station an der folgender Adresse angesetzt:<br><br>Los Santos Police Department<br>Mission Row 1<br>Los Santos, San Andreas<br><br>Wir bitten Sie höflichst, pünktlich zur erscheinen. Die Aussage wird voraussichtlich ${Dauer} (Std.) dauern. Bitte planen Sie dementsprechend ausreichend Zeit ein. Sollten Sie verhindert sein, bitten wir Sie umgehend um eine Benachrichtigung. Wir sind bemüht eine alternative Terminvereinbarung zu treffen.</td>
          </tr>
        </tbody>
      </table>
      <p class="text-center"><br></p>
      <table>
        <tbody>
          <tr>
            <td>Mit freundlichen Grüßen<br><em>${Unterschrift_Unterschrift}<br></em>(${Unterschrift_Detective})</td>
          </tr>
        </tbody>
      </table>
      `
    }

    SpreadsheetApp.getUi().alert(HTML);

    LSPD.Open_HTML("https://www.gvmp.de/index.php?conversation-add/&userID=" + ForumID);
  }
  else
  {
    SpreadsheetApp.getUi().alert("Fehler! Der Mitarbeiter wurde nicht in der Liste gefunden!\n\nBitte wenden Sie sich an den technischen Dienst des LSPD. (Ja, sofort is dringend.)");
  }
}