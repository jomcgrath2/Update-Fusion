var tableIDFusion = '3462201' // Add the table ID of the fusion table here
var rangeName = 'updateFusion' //the name of the range used in the program
    
//create button
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Update Fusion Table", functionName: "updateFusion"}, 
                      {name: "Change Email Information", functionName: "fixEmail"},
                      {name: "Change Range of Data to be Sent (Include Headers)", functionName: "setRangeFusion"}];
  ss.addMenu("Update Fusion", menuEntries);
}



//main function
function updateFusion() {
  var email = UserProperties.getProperty('email'); //gets the user property 'email' out of project properties 
  var password = UserProperties.getProperty('password'); //gets the user property 'password' out of project properties
  //if either email or password is not saved in project properties this will store them there
  if (email === null || password === null) {
    email = Browser.inputBox('Enter email'); //browser box to input email
    password = Browser.inputBox('Enter password'); //browser box to input password
    UserProperties.setProperty('email',email); //sets email as a user property called 'email'
    UserProperties.setProperty('password', password); //sets password as a user property called 'password'
  }
  var authToken = getGAauthenticationToken(email,password); //call getGAauthenticationToken send email and password infomation
  deleteData(authToken, tableIDFusion); //calls delete data sends authToken and tableID
  updateData(authToken, tableIDFusion); //call updateData send authentication token and the tableID
  SpreadsheetApp.getActiveSpreadsheet().toast(Logger.getLog(), "Fusion Tables Update", 10) //browserbox confirmation that info has been sent 
}



//Google Authentication API this is taken directly from the google fusion api website
function getGAauthenticationToken(email, password) {
  password = encodeURIComponent(password);
  var response = UrlFetchApp.fetch("https://www.google.com/accounts/ClientLogin", {
    method: "post",
    payload: "accountType=GOOGLE&Email=" + email + "&Passwd=" + password + "&service=fusiontables&Source=testing"
  });
  var responseStr = response.getContentText();
  responseStr = responseStr.slice(responseStr.search("Auth=") + 5, responseStr.length);
  responseStr = responseStr.replace(/\n/g, "");
  return responseStr;
}



//query fusion API post
function queryFusionTables(authToken, query) {
  var URL = "http://www.google.com/fusiontables/api/query"; //location to send the infomation to
  
  try {
    //sends the the authentication and the query in url format
    var response = UrlFetchApp.fetch(URL, {
      method: "post",
      headers: {
        "Authorization": "GoogleLogin auth=" + authToken,
      },
      payload: "sql=" + query
    });
  } catch(err) {
    if (err.message.search('401') != -1) {
      // If the auth failed, get a new token
      token = getGAauthenticationToken();
      if (!token) {
        return -2;
      }
      return -1;
    } else if (err.message.search('500') != -1) {
      // If there were too many requests being sent, sleep for a bit
      Utilities.sleep(3000);
      return -1;
    } else {
      if (err.message.search('Bad column reference') != -1) {
        Logger.log('Column names in sheet do not match column names in the table.');
      }
      Logger.log(err.message);
      return -2;
    }
  }
  
  
  response = response.getContentText();
  return response;
}



//delete old data in fusion table
function deleteData(authToken, tableID) {
  var query = encodeURIComponent("DELETE FROM " + tableID); //encodes delete info into a url format
  return queryFusionTables(authToken, query); //returns it to queryFusion Tables
}



//this puts all the current information in the spreadsheet into a query
function updateData(authToken, tableID) {
  //find sheets with ranges that will be sent
  var ss = SpreadsheetApp.getActiveSpreadsheet();  //gets the active spreadsheet
  var range = ss.getRangeByName(rangeName);
  var numColumns = range.getNumColumns();
  var optimalSend = Math.round(7500/numColumns) - 1;
  var sendNum = Math.min(500,optimalSend);
  var data = range.getValues();

  /*
  //format date (REMOVE COMMENTS TO ALLOW CODE TO EXECUTE)
  for( var i in data ) {
    for( var j in data[i] ) {
      if (Object.prototype.toString.call(data[i][j]) === '[object Date]') {
        data[i][j] = Utilities.formatDate(data[i][j], "GMT", "MM/dd/yyyy HH:mm:ss");
      }
    }
  }
  */

  //change array to string and remove apostrophe
  var dataString = Utilities.jsonStringify(data);  //turns the data array into a string
  dataString = dataString.replace(/'/g, "\\'"); //removes apostrophes from the stings and replaces them with \' which wont break the code
  data = Utilities.jsonParse(dataString); //turns the data string into an array again
  
  //define headers and prepend and create query
  var headers = data[0]; //defines the headers as the first row in the data array
  var queryPrepend = "INSERT INTO " + tableID + " (" + "\'" +headers.join("\',\'") + "\'" + ") VALUES ('"; //creates the prepend for the query
  var query = ""; //creates the string where the query will be saved
  var count = 0;
  var updatedRowsCount = 0;
  var response;
  
  
  for (var i = 1; i < data.length; ++i) {
    if (data[i][0] != "") {  //if the first column is empty continue will jump to the next i value instead going further
      query += queryPrepend + data[i].join("','") + "'); "; //combines the data with the prepends and adds it to the query string;
      count++
    }
    
    if (count == sendNum || i == data.length - 1) {
      response = queryFusionTables(authToken, encodeURIComponent(query));
      
      // If the query failed with a 401 or 500 error, try again one more time.
      if (response == -1) {
        response = queryFusionTables(authToken, encodeURIComponent(query));
      }
    
      // If the query failed again, or failed for some other reason, return.
      if (response == -1 || response == -2) {
        return;
      }
      
      updatedRowsCount += response.split(/\n/).length - 2;
      query = "";
      count = 0;
    }
  }
  Logger.log("Updated " + updatedRowsCount + " rows in the Fusion Table");
  return;
}



//change email if needed
function fixEmail() {
   var decision = Browser.msgBox("WARNING", "Are you sure you want to change your email?", Browser.Buttons.YES_NO);  //confirmation that you want to change email
   if (decision == 'yes'){
     var email = Browser.inputBox('Enter email');  //input new email
     var password = Browser.inputBox('Enter password');  //input new password
     UserProperties.setProperty('email',email);  //set new email in user properties in the Project Properties
     UserProperties.setProperty('password', password); //set new password in user properties in the Project Properties
   }
}



//set range
function setRangeFusion() {
   var decision = Browser.msgBox("WARNING", "Are you sure you want to change the Update Fusion Range?", Browser.Buttons.YES_NO);
   if (decision == 'yes'){
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     var check = ss.getRangeByName(rangeName)
         if (check != null) {
           ss.removeNamedRange(rangeName);
         }
     var range = SpreadsheetApp.getActiveRange()
     ss.setNamedRange(rangeName, range);
     Browser.msgBox("WARNING", "The range \'" + rangeName + "\' used to send data to Fusion has been changed.", Browser.Buttons.OK);
   }
}