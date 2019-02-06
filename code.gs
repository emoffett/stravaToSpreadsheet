// V2 - started 01/02/19 to address change in Strava API authentication

// TODO:
// good to start from blank
// error messages at top of spreadsheet
// how setup to run automatically?
// look at publishing - GitHub, Google, Strava
// generic version for just grabbing Strava information and dumping it in a spreadsheet
// 'unlocked' version gets token from spreadsheet?
// 'locked' version does not have stravaAccessToken, gets data from webserver (with own token to access webserver)
// getting started sheet filled in
// dashboard sheet
// weekly/monthly summary sheets?


/*
OAuth2 for Apps Script library to authenticate with Strava:
https://github.com/gsuitedevs/apps-script-oauth2
OAuth2 for Apps Script is a library for Google Apps Script that provides the ability to create and authorize OAuth2 tokens as well as refresh them when they expire. This library uses Apps Script's StateTokenBuilder and /usercallback endpoint to handle the redirects.

Setup
This library is already published as an Apps Script, making it easy to include in your project. To add it to your script, do the following in the Apps Script code editor:

Click on the menu item "Resources > Libraries..."
In the "Find a Library" text box, enter the script ID 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF and click the "Select" button.
Choose a version in the dropdown box (usually best to pick the latest version).
Click the "Save" button.
Alternatively, you can copy and paste the files in the /dist directory directly into your script project.

If you are setting explicit scopes in your manifest file, ensure that the following scope is included:

https://www.googleapis.com/auth/script.external_request
*/

var CLIENT_ID = 'YOUR STRAVA CLIENT ID HERE - KEEP THE QUOTES';
var CLIENT_SECRET = 'YOUR SECRET HERE - KEEP THE QUOTES';
/*
On your strava api page, set your application's Authorization Callback Domain to be script.google.com
*/

/**
 * Authorizes and makes a request to the Strava API.
 */
function run() {
  var service = getService_();
  if (service.hasAccess()) {
    Logger.log('have access');
    var url = 'https://www.strava.com/api/v3/activities';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
    Logger.log('starting getStrava2');
    getStrava2(service.getAccessToken());
    Logger.log('ending getStrava2');
  } else {
    Logger.log('do not have access');
    showSidebar();
/*    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);*/
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getService_();
  service.reset();
}

/**
 * Configures the service.
 * Three required and optional parameters are not specified
 * because the library creates the authorization URL with them
 * automatically: `redirect_url`, `response_type`, and
 * `state`.
 */
function getService_() {
  return OAuth2.createService('Strava')
      // Set the endpoint URLs.
      .setAuthorizationBaseUrl('https://www.strava.com/oauth/authorize')
      .setTokenUrl('https://www.strava.com/oauth/token')

      // Set the client ID and secret.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function that should be invoked to
      // complete the OAuth flow.
      .setCallbackFunction('authCallback_')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
}

/**
 * Handles the OAuth callback.
 */
function authCallback_(request) {
  var service = getService_();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Logs the redict URI to register.
 */
function logRedirectUri() {
  var service = getService_();
  Logger.log(OAuth2.getRedirectUri());
}

function showSidebar() {
  var driveService = getService_();
  if (!driveService.hasAccess()) {
    var authorizationUrl = driveService.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    Logger.log('showSidebar has access');
    getStrava2(service.getAccessToken());
  }
}

function getStrava2(stravaAccessToken) {
  // pick the most recent spreadsheet you looked at....
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get the right athlete, based on their stravaAccessToken, as listed in the spreadsheet
  var tab = ss.getSheetByName('Setup');
  ss.setActiveSheet(tab);

  // find the last cell with data
  var tab = ss.getSheetByName('Run');
  ss.setActiveSheet(tab); 
  var bottomCell = tab.getRange("C3:C").getValues().filter(String).length;
  var startCol = 1;
  var startRow = 4;
  var nextRow = tab.getRange(startRow,startCol,bottomCell).getLastRow();
  var lastDate = tab.getRange(nextRow-1,startCol+1).getValue();
  var lastTime = Math.round(lastDate.getTime()/1000);
  var requestData = 'after=' + lastTime;
  Logger.log(requestData);

  // create the request from the strava APR
  var url = 'https://www.strava.com/api/v3/athlete/activities?' + requestData;
  // for more info on Stava setup, e.g. the names of other activity types, see http://strava.github.io/api/v3/activities/
  var activityType = "Run";
  var headers = {
    'Authorization' : 'Bearer ' + stravaAccessToken
  };
  // Make a GET request and log the returned content.
  var options = {
    'method' : 'get',
    'headers' : headers
  };
  Logger.log(url);
  Logger.log(options);
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response);
  var data = JSON.parse(response.getContentText());
  //Logger.log(data);
  
  var rows = [];
  // date, datetime, distance (km), time, average heartrate  
  var working, date, datetime, distance, moving;
  for (i=0; i<data.length; i++){
    working = data[i];
    if (working.type == activityType){
      // PDT is picked so that runs after 17:00 aren't moved forward a date in the date conversion....
      // is this due to where the Google server is located?!
      // will it still work during/after the clocks change? (5th Nov for PDT, 29th Oct for BST)
      date = Utilities.formatDate(new Date(working.start_date_local),"PDT" ,"dd/MM/YYYY");
      Logger.log("start_date_local = " + working.start_date_local + " date = " + date);
      datetime = new Date(working.start_date);
      distance = (working.distance/1000).toFixed(1);
      moving = Math.floor(working.moving_time/3600) + ':' + Math.floor((working.moving_time % 3600) / 60) + ':' + Math.floor(working.moving_time % 60);
      rows.push([date, datetime, distance, moving, working.average_heartrate]);
    }
  }

  if (rows.length > 0){
    tab.getRange(nextRow,startCol+2,rows.length,1).setNumberFormat('HH:mm:ss');
    tab.getRange(nextRow,startCol,rows.length,5).setValues(rows);
  }
}
