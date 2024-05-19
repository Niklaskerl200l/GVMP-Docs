function Update_Scripte()
{
  var ID = "1VT3CT42Z31MwrWcpkeXZmhyqmtebkQLg9N_1hUawWU9nZ2gyCQv1EnY8";

  var Token = ScriptApp.getOAuthToken();

  var Inhalt_JSON = JSON.parse(Get_Script_Content(ScriptApp.getScriptId()));
  var Inhalt_Live_JSON = JSON.parse(Get_Script_Content(ID));

  if(Inhalt_JSON == 0)
  {
    Logger.log("Fehler");

    return 0;
  }

  for(var i = 0; i < Inhalt_JSON.files.length; i++)
  {
    if(Inhalt_JSON.files[i].name == "Update_Script")
    {
      delete Inhalt_JSON.files[i]
    }
    else if (Inhalt_JSON.files[i].name == "appsscript")
    {
      var appsscript = JSON.parse(Inhalt_JSON.files[i].source);

      delete appsscript.oauthScopes;

      Inhalt_JSON.files[i].source = JSON.stringify(appsscript)
    }
  }

  if(Inhalt_JSON == Inhalt_Live_JSON)
  {
    Logger.log("True")
  }
  else
  {
    Logger.log("False")
  }

  //Logger.log(Inhalt_JSON)


  var Optionen = 
  {
    'method' : 'PUT',
    'headers' : {'Authorization' : 'Bearer '+  Token,'Content-Type' : 'application/vnd.google-apps.script+json'},
    'payload' :  JSON.stringify(Inhalt_JSON)
  };

  UrlFetchApp.fetch("https://script.googleapis.com/v1/projects/" + ID + "/content",Optionen);
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