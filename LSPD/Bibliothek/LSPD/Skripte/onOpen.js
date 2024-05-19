for(var x = 0; x < Array_URLs.length; x++)
{
  eval("function Funktion_" + Array_URLs[x][0] + "(){Open_HTML(\'" + Array_URLs[x][3] + Array_URLs[x][1] + "\');}");
}

function Open_HTML(URL)
{
  var UI;

  try
  {
    UI = SpreadsheetApp.getUi();
  }
  catch(err)
  {
    UI = DocumentApp.getUi();
  }

  UI.showModalDialog(HtmlService.createHtmlOutput("<script>window.open(\'" + URL + "\');google.script.host.close();</script>"), "Open Tab")
}

function LSPD_MAP()
{
  Open_HTML("https://map.lspd-gvmp.eu/");
}

function onOpen(UI = true)
{
  if(UI == true)
  {
    try
    {
      UI = SpreadsheetApp.getUi();
    }
    catch(err)
    {
      UI = DocumentApp.getUi();
    }

    var Menu = UI.createMenu('LSPD Dokumente');
    var Array_Abteilung = JSON.parse(Propertie_Lesen("LSPD_Abteilung"));

    Logger.log(Array_Abteilung);

    if(Array_Abteilung != null)
    {
      var Abteilungen;

      //------------------------------- LSPD -------------------------------//
      //Array_Abteilung.forEach(y => { if(y[0] == "IT" || y[0].includes("Directorate of the Police") || y[0] == "LSPD"){Abteilungen = y; return 0;}});
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0] == "LSPD")
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0] == "LSPD" || Abteilungen[0].includes("Directorate of the Police") || Abteilungen[0] == "IT"))
      {
        var Sub_LSPD = UI.createMenu("LSPD");

        Sub_LSPD.addItem("Dienstblatt","LSPD.Funktion_ID_Dienstblatt");
        Sub_LSPD.addSeparator();

        Sub_LSPD.addItem("Tickets melden","LSPD.Funktion_ID_Formular_Ticket");
        Sub_LSPD.addItem("Ortskunde Selbststudium","LSPD.Funktion_ID_Ortskunde_Selbststudium");
        Sub_LSPD.addItem("Interaktive Ortskunde Karte","LSPD.LSPD_MAP");

        if(Abteilungen[2] >= 7)
        {
          Sub_LSPD.addSeparator();
          Sub_LSPD.addItem("Ride Along","LSPD.Funktion_ID_Ride_Along");
        }

        if(Abteilungen[2] >= 9) Sub_LSPD.addItem("Geschwindigkeits Tickets","LSPD.Funktion_ID_Geschwindigkeits_Tickets");
        Sub_LSPD.addItem("Vorlage Doc LSPD","LSPD.Funktion_ID_Vorlage_Doc");

        Menu.addSubMenu(Sub_LSPD);
      }

      //------------------------------- IT -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT")
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && Abteilungen[0] == "IT")
      {
        var Sub_IT = UI.createMenu("IT");

        Sub_IT.addItem("LSPD Master","LSPD.Funktion_ID_LSPD_Master");
        Sub_IT.addItem("Zeitsystem","LSPD.Funktion_ID_Zeitsystem");
        Sub_IT.addItem("Formulare","LSPD.Funktion_ID_Formulare");
        Sub_IT.addItem("Frak Import","LSPD.Funktion_ID_Frak_Import");
        Sub_IT.addItem("LSPD Quelle","LSPD.Funktion_ID_LSPD_Quelle");
        Sub_IT.addItem("LSPD IDs","LSPD.Funktion_ID_LSPD_IDs");
        Sub_IT.addItem("Einsatz","LSPD.Funktion_ID_Einsatz");
        Sub_IT.addItem("Entlassungen","LSPD.Funktion_ID_Entlassungen");
        Sub_IT.addItem("Export Dienstblatt","LSPD.Funktion_ID_Export_Dienstblatt");
        Sub_IT.addItem("Export Personaltabelle","LSPD.Funktion_ID_Export_Personaltabelle");
        Sub_IT.addItem("Export Abteilungen","LSPD.Funktion_ID_Export_Abteilungen");
        Sub_IT.addItem("Export Ausbildungen","LSPD.Funktion_ID_Export_Ausbildungen");
        Sub_IT.addItem("Export ARMY","LSPD.Funktion_ID_Export_ARMY");
        Sub_IT.addItem("Export FIB","LSPD.Funktion_ID_Export_FIB");
        Sub_IT.addItem("Export DOC","LSPD.Funktion_ID_Export_DOC");
        Sub_IT.addItem("Export GOV","LSPD.Funktion_ID_Export_GOV");
        Sub_IT.addItem("Export WN","LSPD.Funktion_ID_Export_WN");
        Sub_IT.addItem("Export DPOS","LSPD.Funktion_ID_Export_Wartung");
        Sub_IT.addItem("Archiv Stempeluhr","LSPD.Funktion_ID_Archiv_Stempeluhr");
        Sub_IT.addItem("Archiv Leitstelle","LSPD.Funktion_ID_Archiv_Leitstelle");
        Sub_IT.addItem("Archiv Dienstblatt Logs","LSPD.Funktion_ID_Archiv_Dienstblatt_Logs");
        Sub_IT.addItem("Archiv Einsatz Logs","LSPD.Funktion_ID_Archiv_Einsatz_Logs");
        Sub_IT.addItem("Archiv Beamtenzahl","LSPD.Funktion_ID_Archiv_Beamtenzahl");
        Sub_IT.addItem("Archiv Parkkrallen","LSPD.Funktion_ID_Archiv_Parkkrallen");
        Sub_IT.addItem("Archiv Beschlagnahmungen","LSPD.Funktion_ID_Archiv_Beschlagnahmungen");
        Sub_IT.addItem("Archiv Aktenvergabe", "LSPD.Funktion_ID_Archiv_Aktenvergabe");
        Sub_IT.addItem("Archiv Einsätze", "LSPD.Funktion_ID_Archiv_Einsaetze");
        Sub_IT.addItem("Namensänderung","LSPD.Funktion_ID_Formular_Namensänderung");
        Sub_IT.addItem("Telefonnummer Änderung","LSPD.Funktion_ID_Formular_Telefonnummer_Änderung");
        Sub_IT.addItem("Entlassungen Forms","LSPD.Funktion_ID_Formular_Entlassungen");
        Sub_IT.addItem("Einstellung","LSPD.Funktion_ID_Formular_Einstellung");
        Sub_IT.addItem("Archiv Kapazität","LSPD.Funktion_ID_Archiv_Kapazität");
        Sub_IT.addItem("LST Suche Archiv", "LSPD.Funktion_ID_Archiv_LSTSuche");

        Menu.addSubMenu(Sub_IT);
      }

      //------------------------------- Direction -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      
      }
      if(Abteilungen != undefined && (Abteilungen[0].includes("Directorate of the Police")|| Abteilungen[0] == "IT"))
      {
        var Sub_Direction = UI.createMenu("Direction");

        Sub_Direction.addItem("Direction","LSPD.Funktion_ID_Direction");
        Sub_Direction.addSeparator();

        Sub_Direction.addItem("Sondergenehmigungen","LSPD.Funktion_ID_Sondergenehmigungen");
        Sub_Direction.addItem("Loyalitätsbonus","LSPD.Funktion_ID_Loyalitätsbonus");
        Sub_Direction.addItem("OOC Beschwerden","LSPD.Funktion_ID_OOC_Beschwerden");
        Sub_Direction.addItem("Inaktivitätsmeldungen","LSPD.Funktion_ID_Inaktivitätsmeldungen");

        Menu.addSubMenu(Sub_Direction);
      }


      //------------------------------- Detective -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0].includes("Detective Bureau"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0].includes("Detective Bureau") || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police")))
      {
        var Sub_Detective = UI.createMenu("Detective");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Detective.addItem("Leitung Detective","LSPD.Funktion_ID_Leitung_Detective");
          Sub_Detective.addSeparator();
        }
        Sub_Detective.addItem("Detective","LSPD.Funktion_ID_Detective");
        Sub_Detective.addSeparator();

        Sub_Detective.addItem("Geschwindigkeits Tickets","LSPD.Funktion_ID_Geschwindigkeits_Tickets");
        Sub_Detective.addItem("Archiv Besprechungen","LSPD.Funktion_ID_Archiv_Detective_Besprechungen");

        Menu.addSubMenu(Sub_Detective);
      }


      //------------------------------- Recruitment -------------------------------//

      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0].includes("Recruitment Division"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0].includes("Recruitment Division") || Abteilungen[0].includes("Directorate of the Police")|| Abteilungen[0] == "IT"))
      {
        var Sub_Recruitment = UI.createMenu("Recruitment");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Recruitment.addItem("Leitung Recruitment","LSPD.Funktion_ID_Leitung_Recruitment");
          Sub_Recruitment.addSeparator();
        }

        Sub_Recruitment.addItem("Recruitment","LSPD.Funktion_ID_Recruitment");
        Sub_Recruitment.addSeparator();

        Sub_Recruitment.addItem("Academy Präsentation","LSPD.Funktion_ID_Academy_Präsentation");
        Sub_Recruitment.addItem("LSPD 1x1","LSPD.Funktion_ID_1x1");
        
        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Recruitment.addSeparator();
          Sub_Recruitment.addItem("EST Antrag","LSPD.Funktion_ID_EST_Antrag");
          Sub_Recruitment.addItem("Archiv EST","LSPD.Funktion_ID_Archiv_EST");
        }
        
        Sub_Recruitment.addSeparator();
        Sub_Recruitment.addItem("Archiv Academy","LSPD.Funktion_ID_Archiv_Academy");
        Sub_Recruitment.addItem("Archiv Bewerbungsgespräche","LSPD.Funktion_ID_Archiv_Bewerbungsgespräche");
        Sub_Recruitment.addItem("Archiv Besprechungen","LSPD.Funktion_ID_Archiv_Recruitment_Besprechungen");

        Menu.addSubMenu(Sub_Recruitment);
      }


      //------------------------------- Training -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0].includes("Training Division"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0].includes("Training Division") || Abteilungen[0].includes("Directorate of the Police")|| Abteilungen[0] == "IT"))
      {
        var Sub_Training = UI.createMenu("Training");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Training.addItem("Leitung Training","LSPD.Funktion_ID_Leitung_Training");
          Sub_Training.addSeparator();
        }

        Sub_Training.addItem("Training","LSPD.Funktion_ID_Training");

        if(Abteilungen[1].includes("Officer Prüfung") || Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Training.addSeparator();
          Sub_Training.addItem("Officer Prüfung","LSPD.Funktion_ID_Officer_Prüfung");
          Sub_Training.addItem("Archiv Officer Prüfung","LSPD.Funktion_ID_Archiv_Officer_Prüfung");
        }

        if(Abteilungen[1].includes("Lieutenant Prüfung") || Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Training.addSeparator();
          Sub_Training.addItem("Lieutenant Prüfung","LSPD.Funktion_ID_Lieutenant_Prüfung");
          Sub_Training.addItem("Archiv Lieutenant Prüfung","LSPD.Funktion_ID_Archiv_Lieutenant_Prüfung");
        }

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Training.addSeparator();
          Sub_Training.addItem("Ausbildungen Vorlage","LSPD.Funktion_ID_Ausbildungen_Vorlage");
        }
        
        Sub_Training.addItem("Archiv Training Besprechungen","LSPD.Funktion_ID_Archiv_Training_Besprechungen");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_Training.addSeparator();
          Sub_Training.addItem("Ausbildungen Dokumente Ordner","LSPD.Funktion_ID_Ausbildungen_Dokumente");
        }

        Sub_Training.addSeparator();
        Sub_Training.addItem("Personal Ausbildungsakten Ordner","LSPD.Funktion_ID_Personal_Ausbildungsakten");
        Sub_Training.addItem("Training Dokumente Ordner","LSPD.Funktion_ID_Training_Dokumente");

        Menu.addSubMenu(Sub_Training);
      }


      //------------------------------- SOC -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0].includes("Special Operation Command"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0].includes("Special Operation Command") || Abteilungen[0].includes("Directorate of the Police")|| Abteilungen[0] == "IT"))
      {
        var Sub_SOC = UI.createMenu("SOC");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_SOC.addItem("Leitung SOC","LSPD.Funktion_ID_Leitung_SOC");
          Sub_SOC.addSeparator();
        }

        Sub_SOC.addItem("SOC","LSPD.Funktion_ID_SOC");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_SOC.addSeparator();
          Sub_SOC.addItem("Modulprüfung","LSPD.Funktion_ID_Modulprüfung");
        }

        Sub_SOC.addSeparator();
        Sub_SOC.addItem("Archiv Besprechungen","LSPD.Funktion_ID_Archiv_SOC_Besprechungen");

        Menu.addSubMenu(Sub_SOC);
      }


      //------------------------------- PRU -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0].includes("Public Relations Unit"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0].includes("Public Relations Unit") || Abteilungen[0].includes("Directorate of the Police")|| Abteilungen[0] == "IT"))
      {
        var Sub_PRU = UI.createMenu("PRU");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_PRU.addItem("Leitung PRU","LSPD.Funktion_ID_Leitung_PRU");
        }
        Sub_PRU.addItem("PRU","LSPD.Funktion_ID_PRU");
        Sub_PRU.addItem("Archiv Besprechnungen","LSPD.Funktion_ID_Archiv_PRU_Besprechnungen");
        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_PRU.addItem("Kummerkasten","LSPD.Funktion_ID_Kummerkasten");
        }
        Sub_PRU.addItem("PRU Vorlagen","LSPD.Funktion_ID_PRU_Vorlagen");

        Menu.addSubMenu(Sub_PRU);
      }


      //------------------------------- GTF -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0].includes("Gang Task Force"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0].includes("Gang Task Force") || Abteilungen[0].includes("Directorate of the Police")|| Abteilungen[0] == "IT"))
      {
        var Sub_GTF = UI.createMenu("GTF");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_GTF.addItem("Leitung GTF","LSPD.Funktion_ID_Leitung_GTF");
          Sub_GTF.addSeparator();
        }

        Sub_GTF.addItem("GTF","LSPD.Funktion_ID_GTF");
        
        Sub_GTF.addSeparator();
        Sub_GTF.addItem("Archiv Besprechungen","LSPD.Funktion_ID_Archiv_GTF_Besprechungen");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_GTF.addSeparator();
          Sub_GTF.addItem("Akteneinträge","LSPD.Funktion_ID_Akteneinträge");
        }

        Menu.addSubMenu(Sub_GTF);
      }

      //------------------------------- WLD -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0].includes("Weapons License Division"))
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0].includes("Weapons License Division") || Abteilungen[0].includes("Directorate of the Police")|| Abteilungen[0] == "IT"))
      {
        var Sub_WLD = UI.createMenu("WLD");

        if(Abteilungen[1] == "Leitung" || Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police"))
        {
          Sub_WLD.addItem("Leitung WLD","LSPD.Funktion_ID_Leitung_WLD");
          Sub_WLD.addSeparator();
        }

        Sub_WLD.addItem("WLD","LSPD.Funktion_ID_WLD");

        Sub_WLD.addSeparator();
        Sub_WLD.addItem("Archiv Besprechungsprotokoll","LSPD.Funktion_ID_Archiv_WLD_Besprechnungen");
        Sub_WLD.addItem("Archiv Kurse","LSPD.Funktion_ID_Archiv_WLD_Kurse");

        Menu.addSubMenu(Sub_WLD);
      }

      //------------------------------- Vertrauensperson -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0] == "Vertrauensperson")
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0] == "Vertrauensperson" ||  Abteilungen[0] == "IT"))
      {
        var Sub_Vertrauensperson = UI.createMenu("Vertrauensperson");

        Sub_Vertrauensperson.addItem("Vertrauensperson","LSPD.Funktion_ID_Vertrauensperson");

        Menu.addSubMenu(Sub_Vertrauensperson);
      }

      //------------------------------- Akteneinträge -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "Akteneinträge")
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && Abteilungen[0] == "Akteneinträge")
      {
        var Sub_Akteneinträge = UI.createMenu("Akteneinträge");

        Sub_Akteneinträge.addItem("Akteneinträge","LSPD.Funktion_ID_Akteneinträge");

        Menu.addSubMenu(Sub_Akteneinträge);
      }

      //------------------------------- Fahrzeugwartung -------------------------------//
      for(var i = 0; i < Array_Abteilung.length; i++)
      {
        if(Array_Abteilung[i][0] == "IT" || Array_Abteilung[i][0].includes("Directorate of the Police") || Array_Abteilung[i][0] == "Fahrzeugwartung")
        {
          Abteilungen = Array_Abteilung[i];
          break;
        }
      }

      if(Abteilungen != undefined && (Abteilungen[0] == "IT" || Abteilungen[0].includes("Directorate of the Police") || Abteilungen[0] == "Fahrzeugwartung"))
      {
        var Sub_Fahrzeugwartung = UI.createMenu("Fahrzeugwartung");

        Sub_Fahrzeugwartung.addItem("Fahrzeugwartung", "LSPD.Funktion_ID_Fahrzeugwartung");

        Menu.addSubMenu(Sub_Fahrzeugwartung);
      }

      if(Abteilungen != undefined)
      {
        Menu.addToUi()
      }

      onOpen_Meldungen(true);
    }

    if(Propertie_Lesen("Popup Copy Aktiv","Script") == "true")
    {
      var Popup_User = Propertie_Lesen("Popup Copy User","Script");

      var Array_User = Popup_User.toString().split("|#|");

      for(var i = 0; i < Array_User.length; i++)
      {
        if(Array_User[i] == Umwandeln(false,false))
        {
          UI.alert("Neue Kopie im Copy Master bitte Prüfen!")
          break;
        }
      }
    }
  }
}

function onOpen_Meldungen(useUI = false)
{
  if(useUI == true)
  {
    try
    {
      var UI = SpreadsheetApp.getUi();
    }
    catch(err)
    {
      try
      {
        var UI = SlidesApp.getUi();
      }
      catch(err)
      {
        try
        {
          DocumentApp.getUi();
        }
        catch(err)
        {
          Logger.log(err.stack);
        }
      }
    }

    try
    {
      var Array_Abteilung = JSON.parse(Propertie_Lesen("LSPD_Abteilung"));
      if(Array_Abteilung != null)
      {
        var Abteilungen;
        Array_Abteilung.forEach(y => { if(y[0] == "IT" || y[0] == "LSPD"){Abteilungen = y;}});

        if((Abteilungen[0] == "LSPD" || Abteilungen[0] == "IT") && Abteilungen[2] >= 8)
        {
          var Meldungen_UI = UI.createMenu("GOV Meldungen");

          for(var i = 0; i < Array_Meldungen.length; i++)
          {
            Meldungen_UI.addItem(`${Array_Meldungen[i][1].toString()}`, `LSPD.Meldung_${Array_Meldungen[i][2].toString()}`);
          }

          Meldungen_UI.addSeparator();
          Meldungen_UI.addItem("Manuelle Sperrzone ausrufen", "LSPD.Meldung_Manuell");

          Meldungen_UI.addSeparator();
          Meldungen_UI.addItem("Eingriffmeldung ausrufen", "LSPD.Schlichtung_Manuell");

          Meldungen_UI.addToUi();
        }
      }
    }
    catch(err)
    {
      Logger.log(err.stack);
    }
  }
}