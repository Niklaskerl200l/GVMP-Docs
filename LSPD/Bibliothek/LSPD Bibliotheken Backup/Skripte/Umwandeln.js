function Umwandeln(Exit_Fehler, Use_UI)
{
  var Name = Propertie_Lesen("LSPD_Name");

  if(Name == null)
  {
    if(Use_UI == true)
    {
      SpreadsheetApp.getUi().alert("Bestätige deine Identität","Bitte Stempel dich im Dienstblatt neu ein (Login)",SpreadsheetApp.getUi().ButtonSet.OK);
    }

    if(Exit_Fehler == true)
    {
      Identität_Bestätigen;
    }
    
    return "Unbekannt";
  }
  else
  {
    return Name;
  }
}