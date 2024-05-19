function Spalte_in_Index(Text)    // String in Spaltenindex
{
  Text = Text.toUpperCase();
  
  if(Text.length == 1)
  {
    return Switch_ABC(Text);
  }
  else if(Text.length == 2)
  {
    return (Switch_ABC(Text[0]) * 26) + Switch_ABC(Text[1]);
  }

}

function Switch_ABC(Text)
{
  switch(Text)
    {
      case "A" : return 1; break;
      case "B" : return 2; break;
      case "C" : return 3; break;
      case "D" : return 4; break;
      case "E" : return 5; break;
      case "F" : return 6; break;
      case "G" : return 7; break;
      case "H" : return 8; break;
      case "I" : return 9; break;
      case "J" : return 10; break;
      case "K" : return 11; break;
      case "L" : return 12; break;
      case "M" : return 13; break;
      case "N" : return 14; break;
      case "O" : return 15; break;
      case "P" : return 16; break;
      case "Q" : return 17; break;
      case "R" : return 18; break;
      case "S" : return 19; break;
      case "T" : return 20; break;
      case "U" : return 21; break;
      case "V" : return 22; break;
      case "W" : return 23; break;
      case "X" : return 24; break;
      case "Y" : return 25; break;
      case "Z" : return 26; break;
      default: return 0;
    }
}

LSPD.Eingabe_Test();

function Abteilung_Sort(a,b)
{
  var Spalte_Name = 1;
  var Spalte_Rang = 0;
  var Spalte_Position = 4;

  if(a[Spalte_Position] == "Leitung" && a[Spalte_Name] != "")
  {
    if(b[Spalte_Position] == "" && b[Spalte_Name] != "")
    {
      return -1;
    }
    else if(b[Spalte_Position] == "Anwärter" && b[Spalte_Name] != "")
    {
      return -1;
    }
    else if(b[Spalte_Position] == "Leitung" && b[Spalte_Name] != "")
    {
      if(a[Spalte_Rang] > b[Spalte_Rang])
      {
        return -1;
      }
      else if(a[Spalte_Rang] < b[Spalte_Rang])
      {
        return 1;
      }
      else if(a[Spalte_Rang] == b[Spalte_Rang])
      {
        if(a[Spalte_Name].toLowerCase() < b[Spalte_Name].toLowerCase())
        {
          return -1;
        }
        else if(a[Spalte_Name].toLowerCase() > b[Spalte_Name].toLowerCase())
        {
          return 1;
        }
      }
    }
  }
  else if(a[Spalte_Position] == "" && a[Spalte_Name] != "")
  {
    if(b[Spalte_Position] == "Anwärter" && b[Spalte_Name] != "")
    {
      return -1;
    }
    else if(b[Spalte_Position] == "" && b[Spalte_Name] != "")
    {
      if(a[Spalte_Rang] > b[Spalte_Rang])
      {
        return -1;
      }
      else if(a[Spalte_Rang] < b[Spalte_Rang])
      {
        return 1;
      }
      else if(a[Spalte_Rang] == b[Spalte_Rang])
      {
        if(a[Spalte_Name].toLowerCase() < b[Spalte_Name].toLowerCase())
        {
          return -1;
        }
        else if(a[Spalte_Name].toLowerCase() > b[Spalte_Name].toLowerCase())
        {
          return 1;
        }
      }
    }
  }
  else if(a[Spalte_Position] == "Anwärter" && a[Spalte_Name] != "")
  {
    if(b[Spalte_Position] == "Anwärter" && b[Spalte_Name] != "")
    {
      if(a[Spalte_Rang] > b[Spalte_Rang])
      {
        return -1;
      }
      else if(a[Spalte_Rang] < b[Spalte_Rang])
      {
        return 1;
      }
      else if(a[Spalte_Rang] == b[Spalte_Rang])
      {
        if(a[Spalte_Name].toLowerCase() < b[Spalte_Name].toLowerCase())
        {
          return -1;
        }
        else if(a[Spalte_Name].toLowerCase() > b[Spalte_Name].toLowerCase())
        {
          return 1;
        }
      }
    }
  }
  
  if(a[Spalte_Name] != "" && b[Spalte_Name] == "")
  {
    return -1;
  }
  return 0;
}