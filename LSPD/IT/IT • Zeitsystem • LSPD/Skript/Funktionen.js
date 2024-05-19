function Zeit_Dauer(Zeit)
{
  Zeit = new Date(Zeit);
  var Stunden;
  var Minuten;
  var Text;

  

  var Dummy = Zeit;
  Dummy.setDate(Dummy.getDate() + 25569)
  Dummy = ((Dummy.getTime() / 1000 / 60 / 60 + 1).toFixed(2)).toString();

  if(Dummy.length == 4)
  {
    Stunden = Dummy[0];
  }
  else if(Dummy.length == 5)
  {
    Stunden = Dummy[0] + Dummy[1];
  }
  else if(Dummy.length == 6)
  {
    Stunden = Dummy[0] + Dummy[1] + Dummy[2];
  }
  else if(Dummy.length == 7)
  {
    Stunden = Dummy[0] + Dummy[1] + Dummy[2] + Dummy[3];
  }

  Minuten = Zeit.getMinutes().toString();

  if(Minuten.length == 1)
  {
    Text = Stunden + ":0" + Minuten;
  }
  else
  {
    Text = Stunden + ":" + Minuten;
  }

  return [Stunden,Minuten,Text];
 
}

LSPD.Eingabe_Test();
