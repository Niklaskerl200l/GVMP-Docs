function Spalte_in_Index(Text)    // String in Spaltenindex
{
  Text = Text.toUpperCase();
  
  if(Text.length == 1)
  {
    return Switch_ABC(Text);
  }
  else if(Text.length == 2)
  {
    return (Switch_ABC(Text[0]) * 26) + Switch_ABC(Text[1]);
  }
}

function Switch_ABC(Text)
{
  switch(Text)
    {
      case "A" : return 1; break;
      case "B" : return 2; break;
      case "C" : return 3; break;
      case "D" : return 4; break;
      case "E" : return 5; break;
      case "F" : return 6; break;
      case "G" : return 7; break;
      case "H" : return 8; break;
      case "I" : return 9; break;
      case "J" : return 10; break;
      case "K" : return 11; break;
      case "L" : return 12; break;
      case "M" : return 13; break;
      case "N" : return 14; break;
      case "O" : return 15; break;
      case "P" : return 16; break;
      case "Q" : return 17; break;
      case "R" : return 18; break;
      case "S" : return 19; break;
      case "T" : return 20; break;
      case "U" : return 21; break;
      case "V" : return 22; break;
      case "W" : return 23; break;
      case "X" : return 24; break;
      case "Y" : return 25; break;
      case "Z" : return 26; break;
      default: return 0;
    }
}

function Kalenderwoche(datum = new Date())
{
  var Datum = new Date(datum.valueOf());
  var TagN = (datum.getDay() + 6) % 7;

  Datum.setDate(Datum.getDate() - TagN + 3);

  var Erster_Donnerstag = Datum.valueOf();

  Datum.setMonth(0, 1);

  if (Datum.getDay() !== 4) 
  {
    Datum.setMonth(0, 1 + ((4 - Datum.getDay()) + 7) % 7);
  }

  return 1 + Math.ceil((Erster_Donnerstag - Datum) / 604800000);  
}

function Umwandeln(Name = false)
{
  var Array_Personal = SpreadsheetApp.getActive().getSheetByName("Personaltabelle").getRange("B4:G").getValues();
  var UI = SpreadsheetApp.getUi();
  var Off_Name = 2;

  if(Name == false) Name = LSPD.Propertie_Lesen("LSPD_Name");

  if(Name == null)
  {
    Name = UI.prompt("Bestätige deine Identität","Gib hier deinen IC Namen ein (!!!OHNE UNTERSTRICH!!!)",UI.ButtonSet.OK).getResponseText();

    if(Name == "" || Name == undefined || Name == null)
    {
      Logger.log("Eingabe Name Leer");
      
      return Umwandeln(null);
    }

    Name = Umwandeln(Name.toString().replace("_"," "));

    LSPD.Propertie_Setzen("LSPD_Name", Name[Off_Name]);    

    SpreadsheetApp.getActive().toast("Erfolgreich Registriert","Erfolg",10);
    return Name;
  }
  else
  {
    if(Name.toString().includes("/Admin "))   // Admin Zugriff
    {
      Name = Name.replace("/Admin ","")
      LSPD.Propertie_Setzen("LSPD_Name", Name);

      Array_Personal[0][Off_Name] = Name;
      return Array_Personal[0];
    }
    
    for(var i = 0; i < Array_Personal.length; i++)
    {
      if(Array_Personal[i][Off_Name].toUpperCase() == Name.toUpperCase())
      {
        Name = Array_Personal[i][Off_Name];

        Logger.log("Name Gefunden: " + Name);

        return Array_Personal[i];

      }
    }

    Logger.log("Name: " + Name + " nicht gefunden");
    UI.alert("Fehler","Name: " + Name + " nicht in Personaltabelle gefunden",UI.ButtonSet.OK);
    
    return Umwandeln(null);
  }
}

function Umwandeln_Abteilung(Name = LSPD.Umwandeln(false,false))
{
  if(Name != null)
  {
    var Sheet_Abteilungen = SpreadsheetApp.getActive().getSheetByName("Abteilungen");
    var Sheet_Personal = SpreadsheetApp.getActive().getSheetByName("Personaltabelle");
    var Sheet_Import = SpreadsheetApp.getActive().getSheetByName("Import Dienstblatt LSPD");

    var Array_Abteilung = Sheet_Abteilungen.getRange(4,2,19,Sheet_Abteilungen.getLastColumn()).getValues();
    var Array_Personal = Sheet_Personal.getRange("B4:F").getValues();
    var Array_Sonder = Sheet_Import.getRange("AE5:AG").getValues();
    var Array_Ausgabe = new Array();

    var Rang = 0;

    for(var y = 0; y < Array_Personal.length; y++)
    {
      if(Array_Personal[y][2] == Name)
      {
        Array_Ausgabe.push(["LSPD", Array_Personal[y][3], Array_Personal[y][0]])

        Rang = Array_Personal[y][0]
      }
    }

    for(var x = 2; x < Array_Abteilung[0].length; x++)
    {
      for(var y = 3; y < Array_Abteilung.length; y++)
      {
        if(Array_Abteilung[y][x] == Name)
        {
          if(Array_Abteilung[y][x+1] == "")
          {
            Array_Abteilung[y][x+1] = "Mitglied";
          }
          
          Array_Ausgabe.push([Array_Abteilung[0][x-2].toString().replace(/[^0-9a-z ]/gi, '').trim(),Array_Abteilung[y][x+1],Rang]);
        }
      }
    }


    for(var y = 0; y < Array_Sonder.length; y++)
    {
      if(Array_Sonder[y][0] == Name)
      {
        Array_Ausgabe.push([Array_Sonder[y][1],Array_Sonder[y][2],Rang]);
      }
    }

    LSPD.Propertie_Setzen("LSPD_Abteilung",JSON.stringify(Array_Ausgabe));
  }
}

LSPD.Eingabe_Test();

function printStackTrace()
{
  const error = new Error();
  const stack = error.stack
    .split('\n')
    .slice(2)
    .map((line) => line.replace(/\s+at\s+/, ''))
    .join('\n');
  return "\n" + stack;
}

function Log_Zaehler(Aktion, Notitz)
{
  var Sheet_Log = SpreadsheetApp.getActive().getSheetByName("Log");

  var Lock = LockService.getDocumentLock();
  try
  {
    Lock.waitLock(28000);
  }
  catch(err)
  {
    throw Error("Log: Zeitüberschreitung!");
  }

  Sheet_Log.insertRowAfter(2);
  Sheet_Log.getRange("B3:G3").setValues([[LSPD.Umwandeln(), new Date(), Aktion, Notitz, Kalenderwoche(), (new Date().getMonth() + 1)]]);

  Lock.releaseLock();
}

function Log_Archivieren()
{
  var Sheet_Log = SpreadsheetApp.getActive().getSheetByName("Log");
  var Array_Log = Sheet_Log.getRange("B3:G").getValues();

  var Array_Ausgabe = [];
  for(var i = Array_Log.length - 1; i >= 0; i--)
  {
    if(Array_Log[i][0] != "")
    {
      Array_Ausgabe.push(Array_Log[i]);
      Logger.log(i + 3);
      Sheet_Log.deleteRow(i + 3);
    }
  }

  if(Array_Ausgabe.length > 0)
  {
    var SS_Archiv = SpreadsheetApp.openById(LSPD.ID_Archiv_Dienstblatt_Logs);
    var Sheet_Archiv = SS_Archiv.getSheetByName("Archiv Log");

    Sheet_Archiv.insertRowsAfter(Sheet_Archiv.getLastRow(), Array_Ausgabe.length);
    Sheet_Archiv.getRange(Sheet_Archiv.getLastRow() + 1, 2, Array_Ausgabe.length, Array_Ausgabe[0].length).setValues(Array_Ausgabe);
  }
}

function BeamtenlisteForum() // Niklas_Kerl®
{
  var ui = SpreadsheetApp.getUi();
  
  var sheet = SpreadsheetApp.getActive().getSheetByName("Dienstbelegung");
  
//------  Deklaration Import Dienstblatt Start  ------//
  
  var Rang0_Namen  = sheet.getRange("B6:B50").getValues();
  var Rang1_Namen  = sheet.getRange("C6:C50").getValues();
  var Rang2_Namen  = sheet.getRange("D6:D50").getValues();
  var Rang3_Namen  = sheet.getRange("E6:E50").getValues();
  var Rang4_Namen  = sheet.getRange("F6:F50").getValues();
  var Rang5_Namen  = sheet.getRange("G6:G50").getValues();
  var Rang6_Namen  = sheet.getRange("H6:H50").getValues();
  var Rang7_Namen  = sheet.getRange("I6:I50").getValues();
  var Rang8_Namen  = sheet.getRange("J6:J50").getValues();
  var Rang9_Namen  = sheet.getRange("K6:K50").getValues();
  var Rang10_Namen = sheet.getRange("L6:L50").getValues();
  var Rang11_Namen = sheet.getRange("M6:M50").getValues();
  var Rang12_Namen = sheet.getRange("N6:N50").getValues();
  
  var Rang0_Titel  = "Recruit";
  var Rang1_Titel  = "Officer in Training";
  var Rang2_Titel  = "Probationary Officer";
  var Rang3_Titel  = "Officer";
  var Rang4_Titel  = "Senior Officer";
  var Rang5_Titel  = "Corporal";
  var Rang6_Titel  = "Senior Corporal";
  var Rang7_Titel  = "Sergeant";
  var Rang8_Titel  = "Lieutenant";
  var Rang9_Titel  = "Captain";
  var Rang10_Titel = "Deputy Chief";
  var Rang11_Titel = "Assistant Chief";
  var Rang12_Titel = "Chief of Police";
  
  var Beamten_Anzahl = sheet.getRange("C2").getValue();
  
  var Kategorie1 = sheet.getRange("B4").getValue();
  var Kategorie3 = sheet.getRange("L4").getValue();
  var Datum = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd.MM.yyyy");
  
//------  Deklaration Import Dienstblatt Ende  ------//
  
  
  
//------------  Definition Manuell Start ------------//                                                         
                                                                                                                
  var Rang0_Bild  = "https://i.imgur.com/YL2FTH1.png";     https://i.imgur.com/YL2FTH1.png      
  var Rang1_Bild  = "https://i.imgur.com/DppbBJe.png";     https://i.imgur.com/O7xxIQa.png
  var Rang2_Bild  = "https://i.imgur.com/O7xxIQa.png";     https://i.imgur.com/O7xxIQa.png 
  var Rang3_Bild  = "https://i.imgur.com/oVkjE2d.png";         
  var Rang4_Bild  = "https://i.imgur.com/lxzlJhH.png";         
  var Rang5_Bild  = "https://i.imgur.com/wNqxmvj.png";         
  var Rang6_Bild  = "https://i.imgur.com/8xedNH2.png";           
  var Rang7_Bild  = "https://i.imgur.com/LjPLYdb.png";      
  var Rang8_Bild  = "https://i.imgur.com/kY2wUdu.png";       
  var Rang9_Bild  = "https://i.imgur.com/VrgVll4.png";       
  var Rang10_Bild = "https://i.imgur.com/ACcofRT.png";      //alt https://i.imgur.com/2ohK4to.png
  var Rang11_Bild = "https://i.imgur.com/9NJlVMy.png";   
  var Rang12_Bild = "https://i.imgur.com/8MVISFv.png";     

var Banner_Bild = "https://i.imgur.com/yEN6een.png";
var Footer_Bild = "https://i.imgur.com/SGJ3vnc.png";

//https://www.badgecreator.com/vb01.php

//Marke	Finish			  Line1 					      Line2		    Seal	  Line3
//S649	Silver			  Police Recruit			  Los Santos	C980M   Police
//S649	Silver / Gold	Police Cadet			    Los Santos	C980M	  Police
//M260	Silver			  Probationary Officer	Los Santos	C980M 	Police
//M260	Silver / Gold	Officer					      Los Santos	C980M	  Police
//M260	Gold			    Senior Officer			  Los Santos	C980M 	Police
//S642	Silver			  Corporal			      	Los Santos	C980M	  Police
//S642	Silver / Gold	Senior Corporal		  	Los Santos	C980M 	Police
//S642	Gold			    Sergeant			      	Los Santos	C980M	  Police
//S510	Silver / Gold	Lieutenant	    			Los Santos	C980M	  Police
//S510	Gold			    Captain			      		Los Santos	C980M 	Police
//S503	Silver			  Chief of Staff	  		Los Santos	C980M	  Police
//S503	Silver / Gold	Deputy Chief	    		Los Santos	C980M 	Police
//S503	Gold			    Chief of Police		   	Los Santos	C980M 	Police

  
  var Rang0_Abk  = "RCT";
  var Rang1_Abk  = "OiT";
  var Rang2_Abk  = "PBO";
  var Rang3_Abk  = "OFC";
  var Rang4_Abk  = "SO";
  var Rang5_Abk  = "CPL";
  var Rang6_Abk  = "SC";
  var Rang7_Abk  = "SGT";
  var Rang8_Abk  = "LT";
  var Rang9_Abk  = "CPT";
  var Rang10_Abk = "DC";
  var Rang11_Abk = "AC";
  var Rang12_Abk = "CoP";
  
//------------  Definition Manuell Ende ------------//  
  
//------------  Umrechnung in HTML Start ------------// 
  
  var Rang0_Namen_HTML  = '';
  var Rang1_Namen_HTML  = '';
  var Rang2_Namen_HTML  = '';
  var Rang3_Namen_HTML  = '';
  var Rang4_Namen_HTML  = '';
  var Rang5_Namen_HTML  = '';
  var Rang6_Namen_HTML  = '';
  var Rang7_Namen_HTML  = '';
  var Rang8_Namen_HTML  = '';
  var Rang9_Namen_HTML  = '';
  var Rang10_Namen_HTML = '';
  var Rang11_Namen_HTML = '';
  var Rang12_Namen_HTML = '';

  for(var i = 0; Rang0_Namen[i] != ''; i++) { Rang0_Namen_HTML = Rang0_Namen_HTML + Rang0_Namen[i] + '<br>'; }
  for(var i = 0; Rang1_Namen[i] != ''; i++) { Rang1_Namen_HTML = Rang1_Namen_HTML + Rang1_Namen[i] + '<br>'; }
  for(var i = 0; Rang2_Namen[i] != ''; i++) { Rang2_Namen_HTML = Rang2_Namen_HTML + Rang2_Namen[i] + '<br>'; }
  for(var i = 0; Rang3_Namen[i] != ''; i++) { Rang3_Namen_HTML = Rang3_Namen_HTML + Rang3_Namen[i] + '<br>'; }
  for(var i = 0; Rang4_Namen[i] != ''; i++) { Rang4_Namen_HTML = Rang4_Namen_HTML + Rang4_Namen[i] + '<br>'; }
  for(var i = 0; Rang5_Namen[i] != ''; i++) { Rang5_Namen_HTML = Rang5_Namen_HTML + Rang5_Namen[i] + '<br>'; }
  for(var i = 0; Rang6_Namen[i] != ''; i++) { Rang6_Namen_HTML = Rang6_Namen_HTML + Rang6_Namen[i] + '<br>'; }
  for(var i = 0; Rang7_Namen[i] != ''; i++) { Rang7_Namen_HTML = Rang7_Namen_HTML + Rang7_Namen[i] + '<br>'; }
  for(var i = 0; Rang8_Namen[i] != ''; i++) { Rang8_Namen_HTML = Rang8_Namen_HTML + Rang8_Namen[i] + '<br>'; }
  for(var i = 0; Rang9_Namen[i] != ''; i++) { Rang9_Namen_HTML = Rang9_Namen_HTML + Rang9_Namen[i] + '<br>'; }
  for(var i = 0; Rang10_Namen[i] != ''; i++) { Rang10_Namen_HTML = Rang10_Namen_HTML + Rang10_Namen[i] + '<br>'; }
  for(var i = 0; Rang11_Namen[i] != ''; i++) { Rang11_Namen_HTML = Rang11_Namen_HTML + Rang11_Namen[i] + '<br>'; }
  for(var i = 0; Rang12_Namen[i] != ''; i++) { Rang12_Namen_HTML = Rang12_Namen_HTML + Rang12_Namen[i] + '<br>'; }
  
//------------  Umrechnung in HTML Ende ------------//  
  
//------------ HTML Programmierung Start ------------//
  var i = 0;

  var Code = '<p class="text-center"><a href="https://www.gvmp.de/forum/index.php?thread/80691-lspd-beamtenliste/"><img src="'+ Banner_Bild + '"></a></p>' +
'<p><br></p>' +  
'<p class="text-center"><strong><span style="font-size: 12pt;">Beamtenliste</span></strong><span style="font-size: 12pt;"><strong></strong></span></p>' +
'<p class="text-center"><em><span style="font-size: 10pt;">( '+ Beamten_Anzahl + ' Beamte )</span></em></p>' +
'<p class="text-center"><em><span style="font-size: 10pt;"><br></span></em></p>' +
'<p class="text-right"><span style="font-size: 10pt;">Zuletzt aktualisiert: ' + Datum + '</span></p>' +
'<p><strong><u><br></u></strong></p>' +
'<p class="text-center"><strong><u></u></strong></p>' +
'<table>' +
	'<tbody>' +
		'<tr>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><img src="'+ Rang12_Bild +'" ><br></td>' +
			'<td class="text-center"><img src="'+ Rang11_Bild +'" ><br></td>' +
			'<td class="text-center"><img src="'+ Rang10_Bild +'" ></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><strong><span style="font-size: 12pt;"><span style="color:#B22222;">' + Rang12_Titel + ' ( Rang 12 )</span></span></strong><span style="color:#D3D3D3;"><span style="font-size: 12pt;"><br></span><span style="font-size: 10pt;"><strong>Abk.: '+ Rang12_Abk + '</strong></span></span><strong><br></strong></td>' +
			'<td class="text-center"><span style="font-size: 12pt;"><span style="color:#B22222;"><strong>' + Rang11_Titel + ' ( Rang 11 )<br></strong></span></span><span style="color:#D3D3D3;"><span style="font-size: 10pt;"><strong>Abk.: ' + Rang11_Abk + '</strong></span></span><br></td>' +
			'<td class="text-center"><span style="font-size: 12pt;"><strong><span style="color:#008000;">'+ Rang10_Titel + ' ( Rang 10 )</span></strong></span><span style="color:#B22222;"><span style="font-size: 12pt;"><strong></strong><br></span></span><span style="color:#D3D3D3;"><span style="font-size: 10pt;"><strong>Abk.: '+ Rang10_Abk +'</strong></span></span><strong></strong><br></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center">'+ Rang12_Namen_HTML+'</td>' +
			'<td class="text-center">'+ Rang11_Namen_HTML+'</td>' +
			'<td class="text-center">'+ Rang10_Namen_HTML+'' +
		'</tr>' +
	'</tbody>' +
'</table>' +
'<p class="text-center"><strong><u></u></strong></p>' +
'<table>' +
	'<tbody>' +
		'<tr>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><img src="'+ Rang9_Bild +'" ><br></td>' +
			'<td class="text-center"><img src="'+ Rang8_Bild +'" ></td>' +
		'</tr>' + 
		'<tr>' +
			'<td class="text-center"><span style="color:#008000;"><strong><span style="font-size: 12pt;">'+ Rang9_Titel + ' ( Rang 9 )</span></strong></span><span style="font-size: 12pt;"><strong><span style="color:#4B0082;"><br></span></strong></span><span style="font-size: 10pt;"><span style="color:#D3D3D3;"><strong>Abk.: '+ Rang9_Abk +'</strong></span></span><strong></strong></td>' +
			'<td class="text-center"><span style="color:#008000;"><strong><span style="font-size: 12pt;">'+ Rang8_Titel + ' ( Rang 8 ) </span></strong></span><span style="font-size: 12pt;"><strong><span style="color:#008000;"></span><span style="color:#4B0082;"><br></span></strong></span><strong><span style="font-size: 10pt;"><span style="color:#D3D3D3;">Abk.: '+ Rang8_Abk + '</span></span></strong><strong><span style="color:#4B0082;"><br></span></strong></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center">'+ Rang9_Namen_HTML + '</td>' +
			'<td class="text-center">'+ Rang8_Namen_HTML + '</td>' +
		'</tr>' +
	'</tbody>' +
'</table>' +
'<p class="text-center"><strong><u></u></strong></p>' +
'<table>' +
	'<tbody>' +
		'<tr>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><img src="'+ Rang7_Bild +'"><br></td>' +
			'<td class="text-center"><img src="'+ Rang6_Bild +'"><br></td>' +
			'<td class="text-center"><img src="'+ Rang5_Bild +'"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><span style="color:#FF8C00;"><strong><span style="font-size: 12pt;">'+ Rang7_Titel +' ( Rang 7 )</span></strong></span><span style="color:#FF8C00;"><br></span><strong><span style="color:#D3D3D3;"><span style="font-size: 10pt;">Abk.: '+ Rang7_Abk +'</span></span></strong></td>' +
			'<td class="text-center"><span style="color:#FF8C00;"><strong><span style="font-size: 12pt;">'+ Rang6_Titel +' ( Rang 6 )</span></strong><br></span><strong><span style="color:#D3D3D3;"><span style="font-size: 10pt;">Abk.: '+ Rang6_Abk +'</span></span></strong><br></td>' +
			'<td class="text-center"><strong></strong><span style="color:#FF8C00;"><span style="font-size: 12pt;"><strong>'+ Rang5_Titel +' ( Rang 5 )</strong></span><strong><br></strong></span><strong><span style="color:#D3D3D3;"><span style="font-size: 10pt;">Abk.: '+ Rang5_Abk +'</span></span></strong><br></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center">'+ Rang7_Namen_HTML + '</td>' +
			'<td class="text-center">'+ Rang6_Namen_HTML + '</td>' +
			'<td class="text-center">'+ Rang5_Namen_HTML + '</td>' +
		'</tr>' +
	'</tbody>' +
'</table>' +
'<p class="text-center"><strong><u></u></strong></p>' +
'<table>' +
	'<tbody>' +
		'<tr>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><br><img src="'+ Rang4_Bild +'"><br></td>' +
			'<td class="text-center"><img src="'+ Rang3_Bild +'"><br></td>' +
			'<td class="text-center"><img src="'+ Rang2_Bild +'"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><strong><span style="color:#0000CD;"><span style="font-size: 12pt;">'+ Rang4_Titel +' ( Rang 4 )</span><span style="font-size: 10pt;"><br></span></span></strong><span style="color:#D3D3D3;"><span style="font-size: 10pt;"><strong>Abk.: '+ Rang4_Abk +'</strong></span></span><br></td>' +
			'<td class="text-center"><span style="font-size: 12pt;"><span style="color:#0000CD;"><strong>'+ Rang3_Titel +' ( Rang 3 )<br></strong></span></span><strong><span style="color:#D3D3D3;"><span style="font-size: 10pt;">Abk.: '+ Rang3_Abk +'</span></span></strong><br></td>' +
			'<td class="text-center"><span style="font-size: 12pt;"><span style="color:#0000CD;"><strong> '+ Rang2_Titel +' (Rang 2)</strong><br></span></span><strong><span style="color:#D3D3D3;"><span style="font-size: 10pt;">Abk.: '+ Rang2_Abk +'</span></span></strong><strong></strong><br></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center">'+ Rang4_Namen_HTML + '</td>' +
			'<td class="text-center">'+ Rang3_Namen_HTML + '</td>' +
			'<td class="text-center">'+ Rang2_Namen_HTML + '</td>' +
		'</tr>' +
	'</tbody>' +
'</table>' +
'<p class="text-center"><strong><u></u></strong></p>' +
'<table>' +
	'<tbody>' +
		'<tr>' +
			'<td class="text-center"></td>' +
			'<td class="text-center"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><img src="'+ Rang1_Bild +'"><br></td>' +
			'<td class="text-center"><img src="'+ Rang0_Bild +'"></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center"><span style="font-size: 12pt;"><strong>'+ Rang1_Titel +' ( Rang 1 )  </strong></span><span style="font-size: 10pt;"><strong><br></strong><strong>Abk.: '+ Rang1_Abk +'</strong></span><br></td>' +
			'<td class="text-center"><span style="font-size: 12pt;"><strong>'+ Rang0_Titel +' ( Rang 0 )<br></strong></span><strong><span style="font-size: 10pt;">Abk.: '+ Rang0_Abk +'</span></strong><br></td>' +
		'</tr>' +
		'<tr>' +
			'<td class="text-center">'+ Rang1_Namen_HTML + '</td>' +
			'<td class="text-center">'+ Rang0_Namen_HTML + '</td>' +
		'</tr>' +
	'</tbody>' +
'</table>' +
'<p><br></p>' +
'<p><a href="https://www.gvmp.de/forum/index.php?thread/80691-lspd-beamtenliste/"><img src="'+ Footer_Bild + '"></a></p>'
  
  
  var HTML_CODE = '';
  
  HTML_CODE = HTML_CODE + Code;
  
  ui.alert(HTML_CODE);
  
//------------ HTML Programmierung Ende ------------//
  
}