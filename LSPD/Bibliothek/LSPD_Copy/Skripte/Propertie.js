function Propertie_Setzen(Key = "LSPD_Name", Wert, Speicher_Art = "User")
{
  Logger.log("Setze " + Speicher_Art + " Propertie " + Key + ": " + Wert);

  if(Speicher_Art == "User")
  {
     PropertiesService.getUserProperties().setProperty(Key,Wert);
  }
  else if(Speicher_Art == "Script")
  {
    PropertiesService.getScriptProperties().setProperty(Key,Wert);
  }
}

function Propertie_Lesen(Key = "LSPD_Name", Speicher_Art = "User")
{
  var Wert;

  if(Speicher_Art == "User")
  {
    Wert = PropertiesService.getUserProperties().getProperty(Key);
  }
  else if(Speicher_Art == "Script")
  {
    Wert = PropertiesService.getScriptProperties().getProperty(Key);
  }
  
  Logger.log("Lese " + Speicher_Art + " Propertie " + Key + ": " + Wert);

  return Wert;
}

function Properties_Delete(Key = "LSPD_Name", Speicher_Art = "User")
{
  Logger.log("Lösche " + Speicher_Art + " Propertie: " + Key)

  if(Speicher_Art == "User")
  {
     PropertiesService.getUserProperties().deleteProperty(Key);
  }
  else if(Speicher_Art == "Script")
  {
    PropertiesService.getScriptProperties().deleteProperty(Key);
  }
}

function Properties_Get_Keys()
{
  Logger.log("Alle Propertie Keys: " + JSON.stringify(PropertiesService.getUserProperties().getProperties()));
}

function Properties_Get_All_Keys(Speicher_Art = "Script")
{
  if(Speicher_Art == "User")
  {
    return Object.entries(PropertiesService.getUserProperties().getProperties());
  }
  else if(Speicher_Art == "Script")
  {
    return Object.entries(PropertiesService.getScriptProperties().getProperties());
  }
}