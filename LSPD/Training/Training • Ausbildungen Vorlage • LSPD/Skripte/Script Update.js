function Update_Scripte()
{
  var Array_ID = SpreadsheetApp.openById(LSPD.ID_Training).getSheetByName("Ausbildungsblatt").getRange("EV4:EV").getValues();

  Logger.log("ID Liste länge = " + Array_ID.length);

  var Token = ScriptApp.getOAuthToken();

  var Inhalt = Get_Script_Content(ScriptApp.getScriptId());

  if(Inhalt == 0)
  {
    Logger.log("Fehler");

    return 0;
  }
  
  var Optionen = 
  {
    'method' : 'PUT',
    'headers' : {'Authorization' : 'Bearer '+  Token,'Content-Type' : 'application/vnd.google-apps.script+json'},
    'payload' : Inhalt
  };

  for(var i = 0; i < Array_ID.length; i++)
  {
    if(Array_ID[i][0] != "" && Array_ID[i][0] != ScriptApp.getScriptId())
    {
      Logger.log("Update Script für ID: " + Array_ID[i][0]);

      UrlFetchApp.fetch("https://script.googleapis.com/v1/projects/" + Array_ID[i][0] + "/content",Optionen);
    }
  }
}

function Get_Script_Content(Script_ID = ScriptApp.getScriptId())
{
  var Token = ScriptApp.getOAuthToken();
  
  var Optionen = 
  {
    'method' : 'GET',
    'headers' : {'Authorization' : 'Bearer '+  Token}
  };

  Logger.log("Get Content from Script: " + Script_ID);

  try
  {
    return UrlFetchApp.fetch("https://script.googleapis.com/v1/projects/" + Script_ID + "/content",Optionen);
  }
  catch(err)
  {
    Logger.log(err);

    return 0;
  }
}