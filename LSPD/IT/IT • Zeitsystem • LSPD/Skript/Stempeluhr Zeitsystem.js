
//---------------- Stempeluhr Zeitsystem  ©Niklas Kerl -----------------------//

var Glob_Array_Zeitsystem_Namen;
var Glob_Array_Zeitsystem_Zeit;
var Glob_Array_Zeitsystem_Anwesenheit;
var Glob_Array_Log_Stp = [];
var Glob_Jetzt = new Date();

function Stempeluhr_Zeitsystem()
{
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
  var Sheet_Zeitsystem = SpreadsheetApp.getActive().getSheetByName("Zeitsystem");
  var Sheet_Log_Stp = SpreadsheetApp.getActive().getSheetByName("Log Stempeluhr");

  var Letzte_Zeile_Import_Stp = Sheet_Auswertung.getRange("B1").getValue();
  var Letzte_Zeile_Speicher_Stp = Sheet_Auswertung.getRange("E1").getValue();
  var Letzte_Zeile_Zeitsystem = Sheet_Zeitsystem.getRange("B1").getValue();
  var Letzte_Zeile_Log_Stp = Sheet_Log_Stp.getRange("B1").getValue();

  Glob_Array_Zeitsystem_Namen = Sheet_Zeitsystem.getRange("B4:B" + Letzte_Zeile_Zeitsystem).getValues();
  Glob_Array_Zeitsystem_Zeit = Sheet_Zeitsystem.getRange("F4:F" + Letzte_Zeile_Zeitsystem).getValues();
  Glob_Array_Zeitsystem_Anwesenheit = Sheet_Zeitsystem.getRange("R4:R" + Letzte_Zeile_Zeitsystem).getValues();

  var Schnellsuche;
  var Gefunden;

  if(Letzte_Zeile_Import_Stp == 3)  // Wenn Import leer
  {
    Letzte_Zeile_Import_Stp = 4;
  }
  
  var Array_Import = Sheet_Auswertung.getRange("B4:C" + Letzte_Zeile_Import_Stp).getValues();


  if(Letzte_Zeile_Speicher_Stp == 3)  // Wenn Speicher leer
  {
    Letzte_Zeile_Speicher_Stp = 4;
  }
  
  var Array_Speicher = Sheet_Auswertung.getRange("E4:F" + Letzte_Zeile_Speicher_Stp).getValues();


//----------- Suche neuer Loggin -----------------//

  for(var y1 = 0; y1 < Array_Import.length; y1++)
  {
    Schnellsuche = false;

    try
    {
      Schnellsuche = Array_Import[y1][0] == Array_Speicher[y1][0] && Array_Import[y1][1].toString() == Array_Speicher[y1][1].toString();
    }
    catch(e){}

    if(Schnellsuche)  //Schnellsuche (gleiche Position) Minute + 1
    {
      Minute_Up_Stp(Array_Import[y1][0]);
    }
    else
    {
      if(y1 <= (Array_Import.length / 2))   // Suche kleiner als bei Hälfte (Performance)
      {
        for(var y2 = 0; y2 < Array_Speicher.length; y2++)  //Suche im Speicher nach Person
        {
          if(Array_Import[y1][0] == Array_Speicher[y2][0])  //Person gefunden
          {
            if(Array_Import[y1][1].toString() != Array_Speicher[y2][1].toString())  //Umgetragen Abteilung (Logout)
            { 
              Logout_Log_Stp(Array_Import[y1][0],Array_Speicher[y2][1]);

              break;
            }
            break;
          }
        }
      }
      else                                                      // Suche von unten nach oben ab Hälfte Import (Performance)
      {
        for(var y2 = Array_Speicher.length - 1; y2 >= 0; y2--)  //Suche im Speicher nach Person
        {
          if(Array_Import[y1][0] == Array_Speicher[y2][0])  //Person gefunden
          {
            if(Array_Import[y1][1].toString() != Array_Speicher[y2][1].toString())  //Umgetragen Abteilung (Logout )
            {
              Logout_Log_Stp(Array_Import[y1][0],Array_Speicher[y2][1]);

              break;
            }
            break;
          }
        }
      }
      
      Minute_Up_Stp(Array_Import[y1][0]);     // Setzt die Zeiten der Eingeloggt + 1 Minute (ausgenomen schnellsuche)
    }
  }
  
//--------------- Suche Logout ----------------//

  for(var y2 = 0; y2 < Array_Speicher.length; y2++)
  {
    Gefunden = false;
    Schnellsuche = false;

    try
    {
      var Schnellsuche = Array_Speicher[y2][0] == Array_Import[y2][0];
    }
    catch(e){}

    if(Schnellsuche)  //Schnellsuche (gleiche Position)
    {
      Gefunden = true;
    }
    else
    {
      if(y1 <= (Array_Speicher.length / 2))     // Suche kleiner als bei Hälfte (Performance)
      {
        for(var y1 = 0; y1 < Array_Import.length; y1++)
        {
          if(Array_Speicher[y2][0] == Array_Import[y1][0])
          {
            Gefunden = true;
            break;
          }
        }
      }
      else                                    // Suche von unten nach oben ab Hälfte Import (Performance)
      {
        for(var y1 = Array_Import.length - 1; y1 >= 0; y1--)
        {
          if(Array_Speicher[y2][0] == Array_Import[y1][0])
          {
            Gefunden = true;
            break;
          }
        }
      }
    }

    if(Gefunden == false)   //Person nicht gefunden Logout
    {
      Logout_Log_Stp(Array_Speicher[y2][0],Array_Speicher[y2][1]);
    }
  }

  if(Glob_Array_Log_Stp.length != 0)
  {
    Sheet_Log_Stp.getRange("B" + Letzte_Zeile_Log_Stp + ":D" + (Letzte_Zeile_Log_Stp + Glob_Array_Log_Stp.length - 1)).setValues(Glob_Array_Log_Stp);   // Setzt die Log einträge ins Log
  }

  Sheet_Zeitsystem.getRange("F4:F" + Letzte_Zeile_Zeitsystem).setValues(Glob_Array_Zeitsystem_Zeit);          // Setzt die Zeiten ins blatt

  Sheet_Zeitsystem.getRange("R4:R" + Letzte_Zeile_Zeitsystem).setValues(Glob_Array_Zeitsystem_Anwesenheit);    // Setzt die Letzte Anwesenheit ins Blatt

  Sheet_Auswertung.getRange("E4:F" + Letzte_Zeile_Speicher_Stp).setValue("");                                 // Speicher leeren

  Sheet_Auswertung.getRange("E4:F" + Letzte_Zeile_Import_Stp).setValues(Array_Import);                        // Werte aus Import in den Speicher schreiben
}

function Minute_Up_Stp(Name)
{
  for(var y = 0; y < Glob_Array_Zeitsystem_Namen.length; y++)
  {
    if(Glob_Array_Zeitsystem_Namen[y][0] == Name)   //Suche nach Namen
    {
      Glob_Array_Zeitsystem_Zeit[y][0].setMinutes(Glob_Array_Zeitsystem_Zeit[y][0].getMinutes() + 1);  // Setzt Heute + 1 min

      Glob_Array_Zeitsystem_Anwesenheit[y][0] = Glob_Jetzt;                                            // Setzt Letzte Aktivität
    }
  }
}

function Logout_Log_Stp(Name, Beginn)   //Zeit ins Log schreiben
{
  Glob_Array_Log_Stp.push([Name,Beginn,Glob_Jetzt]);
}