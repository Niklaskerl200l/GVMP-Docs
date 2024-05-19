function createTimeTriggerSpecifcDate() {
 //Create a trigger to run a function called "runScript"
 ScriptApp.newTrigger("onSubmit")
   //The trigger is time based
   .forSpreadsheet(SpreadsheetApp.getActive())
   .onFormSubmit()
   //Configure when to run the trigger
   //Finally create the trigger
   .create();
}