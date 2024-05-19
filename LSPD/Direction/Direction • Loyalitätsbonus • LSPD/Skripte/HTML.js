function HTML()
{
  var Sheet_Auszahlung = SpreadsheetApp.getActive().getSheetByName("Auszahlungstabelle LSPD");

  var Letzte_Zeile = Sheet_Auszahlung.getRange("Q3").getValue();

  var Array = Sheet_Auszahlung.getRange("Q5:T" + Letzte_Zeile).getValues();

  var Liste = "";
  
  for(var y = 0; y < Array.length; y++)
  {
    if(Array[y][3] != "")
    {
      Liste = Liste + '<tr><td>' + Array[y][0] + '[user=\'' + Array[y][3] + '\']' + Array[y][0] + '[/user]</td><td>für ' + Array[y][2] +' Monate ' + Array[y][1] +'$</td></tr>';
    }
    else
    {
      Liste = Liste + '<tr><td>' + Array[y][0] + '</td><td>für ' + Array[y][2] +' Monate ' + Array[y][1] +'$</td></tr>';
    }
    
  }
  
  var Code = '<p class="text-center"><img src="https://www.gvmp.de/index.php?image-proxy/&amp;key=7727cc4769257ba5b078f66d22c3e617be1caf86ac68b490076d6624b43dc560-aHR0cHM6Ly9pLmltZ3VyLmNvbS9GWWd5anZsLnBuZw%3D%3D" alt="FYgyjvl.png" data-valid="true"></p>' + 
'<p class="text-center"><br></p>' + 
'<p><br></p>' + 
'<p><br></p>' + 
'<p class="text-center"><span style="font-size: 36pt;"><span style="color:#00FF00;"><u><strong>- Loyalitätsbonus -</strong></u></span></span></p>' + 
'<p><br></p>' + 
'<p><br></p>' + 
'<p><span style="font-size: 14pt;">Sehr geehrte Beamtinnen und Beamte,</span></p>' + 
'<p><span style="font-size: 14pt;"><br></span></p>' + 
'<p><span style="font-size: 14pt;">folgende Personen können sich bei mir melden und Ihren Loyalitätsbonus abholen.</span></p>' + 
'<p><br></p>' + 
'<div class="messageTableOverflow">' +
	'<table>' +
		'<tbody>' +
			Liste +
		'</tbody>' +
	'</table>' +
'</div>' +
'<p><br></p>' + 
'<p><br></p>' + 
'<p><span style="color:#00FF00;"></span></p>' + 
'<p><span style="font-size: 14pt;"><br></span></p>' + 
'<p><span style="font-size: 14pt;">Im Namen der Regierung bedanken wir uns für den Einsatz, den ihr jeden Tag für uns leistet!<br></span></p>' +
'<p><span style="font-size: 14pt;"><br></span></p>' +
'<p><img src="https://i.imgur.com/NmB5pod.png"></p>' +
'<p><br></p>' +
'<p><br></p>' +
'<p></p>' +
'<p><span style="font-size: 14pt;">Mit freundlichen Grüßen</span></p>' +
'<p><span style="font-size: 14pt;">Chief of Police<br></span></p>' +
'<p><br></p>';

SpreadsheetApp.getUi().alert(Code);
}
