function Test()
{
  var Sheet_Fragen = SpreadsheetApp.getActive().getSheetByName("Forumlar");

  var Form = FormApp.openById("1Gvt8QlBw9rq19MA87nm2tUgd6vAp7yfUcAXgYq2f9Bc");

  Form.setIsQuiz(true)

  var Items = Form.getItems();

  var Array_Fragen = Sheet_Fragen.getRange("B2:N" + Sheet_Fragen.getLastRow()).getValues();

  for(var i = 0; i < Items.length; i++)
  {
    if(i >= 1)
    {
      Form.deleteItem(Items[i]);
    }
  }

  for(var i = 2; i < Array_Fragen.length; i++)
  {
    if(Array_Fragen[i][2] == "Checkbox")
    {
      var Item = Form.addCheckboxItem();
      var Array_Antworten = new Array();

      

      for(var x = 3; x < 8; x++)
      {
        if(Array_Fragen[i][x] != "")
        {
          Array_Antworten.push(Item.createChoice(Array_Fragen[i][x], Array_Fragen[i][x+5]))
        }
      }

      if(Array_Antworten.length > 0)
      {
        Item.setChoices(Array_Antworten);
      }

      Item.setTitle(Array_Fragen[i-2][0])
      Item.setPoints(Array_Fragen[i][1]);

    }
    else if(Array_Fragen[i][2] == "Lang Text")
    {
      var Item2 = Form.addParagraphTextItem();

      Item2.setValidation(FormApp.createTextValidation().requireTextEqualTo(Array_Fragen[i][3]).build());

      Item2.setTitle(Array_Fragen[i-2][0])
      Item2.setPoints(Array_Fragen[i][1]);
    }
  }
}
