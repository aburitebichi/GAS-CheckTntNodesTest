function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Node check", functionName: "myFunction"}];
  ss.addMenu("script", menuEntries);

  var range= SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 1);
  // write sheet head
  var titles = ["Address", "Verify TIME", "IP TEST", "TIME TEST", "CALENDAR TEST", "CREDIT TEST", "TNT VOL"];
  for(var i=0; i<7; i++) {
    range.offset(0, i).setValue(titles[i]);
  }
}

function setTrigger() {
  // run every hours
  ScriptApp.newTrigger('myFunction').timeBased().everyHours(1).create();
}

function U2Gtime(unixtime) {
  var newDate = new Date( );
  newDate.setTime( unixtime );
  dateString = newDate.toUTCString( );
  return dateString
}

function myFunction() {

  var range= SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(1, 1);
  var row= SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getLastRow();

  for(var n=1; n<(row); n++) {
    // read address to write
    var address = range.offset(n, 0).getValue();

    // clear sheet
    range.offset(n, 1, n, 6).clearContent();

    // call API
    var response = UrlFetchApp.fetch("https://c.chainpoint.org/nodes/"+ address +"",{ muteHttpExceptions:true });
    if(response.getResponseCode() != 200) {
      range.offset(n, 1).setValue("404 not found");
      continue;
    }
    // parse API result
    var testResult = JSON.parse(response.getContentText());

    // write sheet
    //for(var i=0; i<testResult.recent_audits.length; i++) {
    for(var i=0; i<1; i++) {
      var event = testResult.recent_audits[i];

      // time of verify
      range.offset(n, 1).setValue(U2Gtime(event.time));

      //public ip
      range.offset(n, 2).setValue(event.public_ip_test);

      // time
      range.offset(n, 3).setValue(event.time_test);

      // calendar
      range.offset(n, 4).setValue(event.calendar_state_test);

      // credit
      range.offset(n, 5).setValue(event.minimum_credits_test);
    }

    // call Ethplorer API
    var response = UrlFetchApp.fetch("api.ethplorer.io/getAddressInfo/"+ address +"?apiKey=freekey",{ muteHttpExceptions:true });
    if(response.getResponseCode() != 200) {
      range.offset(n, 1).setValue("404 not found");
      continue;
    }
    // parse API result
    var testResult = JSON.parse(response.getContentText());
    //range.offset(n+1, 7).setValue(testResult.tokens[0].tokenInfo.symbol);
    range.offset(n, 6).setValue(testResult.tokens[0].balance/100000000);

    Utilities.sleep(5000);
  }
}
