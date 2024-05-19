function Formular_Trigger()
{
  try
  {
    Freigaben();
  }
  catch(err)
  {
    Logger.log(err.stack);
  }

  try
  {
    Freigaben_Fehler();
  }
  catch(err)
  {
    Logger.log(err.stack);
  }

  try
  {
    Lese_Rechte();
  }
  catch(err)
  {
    Logger.log(err.stack);
  }

  try
  {
    Schutz();
  }
  catch(err)
  {
    Logger.log(err.stack);
  }
}