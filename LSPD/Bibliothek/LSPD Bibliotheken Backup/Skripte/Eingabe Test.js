function Eingabe_Test(Array_Mails)
{  
  var Array_Mails = ["gvmp.lspd.bot@gmail.com","niklaskerl2001@gmail.com","bmfaith31@gmail.com"];
  
  try
  {
    var SS_Sheet = SpreadsheetApp.getActiveSpreadsheet();

    if(Array_Mails.includes(SS_Sheet.getOwner().getEmail()) == false)
    {
      if(SS_Sheet.getName != "Kopierschutz")
      {
        var Zeit = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"yyyy.MM.dd HH:mm:ss")
        PropertiesService.getScriptProperties().setProperty(Zeit,SS_Sheet.getName() + "|#|" + Umwandeln(false,false) + "|#|" + SS_Sheet.getOwner().getEmail() + "|#|" + "https://docs.google.com/spreadsheets/d/" + SpreadsheetApp.getActiveSpreadsheet().getId());

        SS_Sheet.rename("Kopierschutz")
      }
     
      var Sheets = SS_Sheet.getSheets();
      
      try
      {
        Sheets[0].setActiveSelection("A1");
      }
      catch(err)
      {
        Logger.log(err.stack);
      }

      for(var i = 1; i < Sheets.length; i++)
      {
        try
        {
          SS_Sheet.deleteSheet(Sheets[i]);
        }
        catch(err)
        {
          Logger.log(err.stack);
        }
      }

      try
      {
        SS_Sheet.deleteSheet(Sheets[0]);
      }
      catch(err)
      {
        Logger.log(err.stack);
      }

      SpreadsheetApp.flush();

      Sheets = SS_Sheet.getSheets();

      for(var i = 0; i < Sheets.length; i++)
      {
        var Sheet = Sheets[i];
        
        try{Sheet.hideColumns(2,Sheet.getLastColumn() - 1);}catch(err){}
        try{Sheet.hideRows(2,Sheet.getLastRow() - 1);}catch(err){}
        try{Sheet.deleteColumns(2,Sheet.getLastColumn() - 1);}catch(err){}
        try{Sheet.deleteRows(2,Sheet.getLastRow() - 1);}catch(err){}
        try{SS_Sheet.deleteSheet(Sheet);}catch(err){}
      }
    }
  }
  catch(err)
  {
    Logger.log(err.stack);

    try
    {
      var SS_Doc = DocumentApp.getActiveDocument()

      var Gefunden = true;
      var Array_Editors = SS_Doc.getEditors();

      for(var i = 0; i < Array_Mails.length; i++)
      {
        for(var x = 0; x < Array_Editors.length; x++)
        {
          if(Array_Mails[i] == Array_Editors[x].getEmail())
          {
            Gefunden = false;
            i = Array_Mails.length;
            break;
          }
        }
      }
   
      if(Gefunden)
      {
        if(SS_Doc.getName != "Kopierschutz")
        {
          var Zeit = Utilities.formatDate(new Date(),"GMT+1","yyyy.MM.dd HH:mm:ss")
          PropertiesService.getScriptProperties().setProperty(Zeit,SS_Doc.getName() + "|#|" + Umwandeln(false,false) + "|#|" + Array_Editors[0].getEmail() + "|#|" + " https://docs.google.com/document/d/" + SS_Doc.getId());

          SS_Doc.setName("Kopierschutz")
        }

        SS_Doc.getBody().clear();
        SS_Doc.getFooter().clear();
        SS_Doc.getHeader().clear();
      }
    }
    catch(err)
    {
      Logger.log(err.stack);

      try
      {
        var SS_Slides = SlidesApp.getActivePresentation()

        if(Array_Mails.includes(SS_Slides.getOwner().getEmail()) == false)
        {
          if(SS_Slides.getName != "Kopierschutz")
          {
            var Zeit = Utilities.formatDate(new Date(),"GMT+1","yyyy.MM.dd HH:mm:ss")
            PropertiesService.getScriptProperties().setProperty(Zeit + " " + SS_Slides.getName() + " https://docs.google.com/presentation/d/"  + SS_Slides.getId(), Umwandeln(false,false));

            SS_Slides.rename("Kopierschutz")
          }
        }
      }
      catch(err)
      {
        Logger.log(err.stack);
      }
    }
  }
}