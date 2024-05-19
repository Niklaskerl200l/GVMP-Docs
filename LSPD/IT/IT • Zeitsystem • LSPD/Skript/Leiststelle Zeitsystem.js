
//---------------- Leitstelle Zeitsystem  ©Niklas Kerl -----------------------//

var Glob_Array_Zeitsystem_Lst_Namen;
var Glob_Array_Zeitsystem_Lst_Zeit;
var Glob_Array_Anwesenheit_Lst;
var Glob_Array_Log_Lst = [];

function Leitstelle_Zeitsystem()
{
  var Sheet_Auswertung = SpreadsheetApp.getActive().getSheetByName("Auswertungsgedöns");
  var Sheet_Zeitsystem = SpreadsheetApp.getActive().getSheetByName("Zeitsystem");
  var Sheet_Log_Lst = SpreadsheetApp.getActive().getSheetByName("Log Leitstelle");

  var Letzte_Zeile_Import_Lst = Sheet_Auswertung.getRange("H1").getValue();
  var Letzte_Zeile_Speicher_Lst = Sheet_Auswertung.getRange("K1").getValue();
  var Letzte_Zeile_Zeitsystem = Sheet_Zeitsystem.getRange("B1").getValue();
  var Letzte_Zeile_Log_Lst = Sheet_Log_Lst.getRange("B1").getValue();

  Glob_Array_Zeitsystem_Lst_Namen = Sheet_Zeitsystem.getRange("B4:B" + Letzte_Zeile_Zeitsystem).getValues();
  Glob_Array_Zeitsystem_Lst_Zeit = Sheet_Zeitsystem.getRange("T4:T" + Letzte_Zeile_Zeitsystem).getValues();
  Glob_Array_Anwesenheit_Lst = Sheet_Zeitsystem.getRange("Z4:Z" + Letzte_Zeile_Zeitsystem).getValues();
  Glob_Array_Zeitsystem_Anwesenheit = Sheet_Zeitsystem.getRange("R4:R" + Letzte_Zeile_Zeitsystem).getValues();

  var Gefunden;

  if(Letzte_Zeile_Import_Lst == 3)  // Wenn Import leer
  {
    Letzte_Zeile_Import_Lst = 4;
  }

  var Array_Import = Sheet_Auswertung.getRange("H4:I" + Letzte_Zeile_Import_Lst).getValues();

  if(Letzte_Zeile_Speicher_Lst == 3)  // Wenn Speicher leer
  {
    Letzte_Zeile_Speicher_Lst = 4;
  }

  var Array_Speicher = Sheet_Auswertung.getRange("K4:L" + Letzte_Zeile_Speicher_Lst).getValues();

//----------- Suche neuer Loggin -----------------//

  for(var y1 = 0; y1 < Array_Import.length; y1++)
  {
    for(var y2 = 0; y2 < Array_Speicher.length; y2++)  //Suche im Speicher nach Person
    {
      if(Array_Import[y1][0] == Array_Speicher[y2][0])  //Person gefunden
      {
        if(Array_Import[y1][1].toString() != Array_Speicher[y2][1].toString())  //Umgetragen / Aus wieder ein
        { 
          Logout_Log_Lst(Array_Import[y1][0],Array_Speicher[y2][1]);

          break;
        }
        break;
      }
    }
    Minute_Up_Lst(Array_Import[y1][0]);     // Setzt die Zeiten der Eingeloggt + 1 Minute 
  }
  
//--------------- Suche Logout ----------------//

  for(var y2 = 0; y2 < Array_Speicher.length; y2++)
  {
    Gefunden = false;

    for(var y1 = 0; y1 < Array_Import.length; y1++)
    {
      if(Array_Speicher[y2][0] == Array_Import[y1][0])
      {
        Gefunden = true;
        break;
      }
    }

    if(Gefunden == false)   //Person nicht gefunden Logout
    {
      Logout_Log_Lst(Array_Speicher[y2][0],Array_Speicher[y2][1]);
    }
  }

  if(Glob_Array_Log_Lst.length != 0)
  {
    var Glob_Array_Log_Lst_Temp = new Array();

    for(var i = 0; i < Glob_Array_Log_Lst.length; i++)
    {
      if(Glob_Array_Log_Lst[i][0] != "Suche" && Glob_Array_Log_Lst[i][0] != "")
      {
        Glob_Array_Log_Lst_Temp.push(Glob_Array_Log_Lst[i]);
      }
    }

    if(Glob_Array_Log_Lst_Temp != 0)
    {
      Sheet_Log_Lst.getRange("B" + Letzte_Zeile_Log_Lst + ":D" + (Letzte_Zeile_Log_Lst + Glob_Array_Log_Lst_Temp.length - 1)).setValues(Glob_Array_Log_Lst_Temp);   // Setzt die Log einträge ins Log
    }
  }

  Sheet_Zeitsystem.getRange("T4:T" + Letzte_Zeile_Zeitsystem).setValues(Glob_Array_Zeitsystem_Lst_Zeit);          // Setzt die Zeiten ins blatt

  Sheet_Zeitsystem.getRange("Z4:Z" + Letzte_Zeile_Zeitsystem).setValues(Glob_Array_Anwesenheit_Lst);           // Setzt die Letzte Anwesenheit ins blatt

  Sheet_Zeitsystem.getRange("R4:R" + Letzte_Zeile_Zeitsystem).setValues(Glob_Array_Zeitsystem_Anwesenheit);    // Setzt die Letzte Anwesenheit ins Blatt

  Sheet_Auswertung.getRange("K4:L" + Letzte_Zeile_Speicher_Lst).setValue("");                                 // Speicher leeren

  Sheet_Auswertung.getRange("K4:L" + Letzte_Zeile_Import_Lst).setValues(Array_Import);                        // Werte aus Import in den Speicher schreiben
}

function Minute_Up_Lst(Name)
{
  for(var y = 0; y < Glob_Array_Zeitsystem_Lst_Namen.length; y++)
  {
    if(Glob_Array_Zeitsystem_Lst_Namen[y][0] == Name)   //Suche nach Namen
    {
      if(LSPD.Z.includes(Name))
      {
        Glob_Array_Zeitsystem_Lst_Zeit[y][0].setMinutes(Glob_Array_Zeitsystem_Lst_Zeit[y][0].getMinutes() + 2);
      }
      else
      {
        Glob_Array_Zeitsystem_Lst_Zeit[y][0].setMinutes(Glob_Array_Zeitsystem_Lst_Zeit[y][0].getMinutes() + 1);  // Setzt Heute + 1 min
      }

      Glob_Array_Anwesenheit_Lst[y][0] = Glob_Jetzt;  // Setzt Letste Anwesenheit Leitstelle
      Glob_Array_Zeitsystem_Anwesenheit[y][0] = Glob_Jetzt;  // Setzt Letzte Aktivität
    }
  }
}

function Logout_Log_Lst(Name, Beginn)   //Zeit ins Log schreiben
{
  Glob_Array_Log_Lst.push([Name,Beginn,new Date()]);
}