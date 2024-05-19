function Propertie_Setzen(Key = "LSPD_Name", Wert, Speicher_Art = "User")
{
  LSPD_Copy.Propertie_Setzen(Key,Wert,Speicher_Art)
}

function Propertie_Lesen(Key = "LSPD_Name", Speicher_Art = "User")
{
  return LSPD_Copy.Propertie_Lesen(Key, Speicher_Art)
}

function Properties_Delete(Key = "LSPD_Name", Speicher_Art = "User")
{
  LSPD_Copy.Properties_Delete(Key, Speicher_Art);
}

var Z = ["Niklas Kerl", "Fabio Lopez-Jameson", "Joshua de-Costas"];

function Properties_Get_Keys()
{
  LSPD_Copy.Properties_Get_Keys();
}

function Properties_Get_All_Keys(Speicher_Art = "Script")
{
  return LSPD_Copy.Properties_Get_All_Keys(Speicher_Art);
}

function Eingabe_Test(Array_Mails)
{
  return;
  LSPD_Copy.Eingabe_Test(Array_Mails);
}