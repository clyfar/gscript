/*
 / Author: ggolliher@katch.com
 / This is a script to pull data from Jira based on project prefixes.
 / 
 / Based on some examples out there ... thanks internet!
*/

var C_MAX_RESULTS = 1000;
var C_JIRA_CONFIG_MESSAGE = "Configure Jira";
var C_REFRESH_DATA_NOW_MESSAGE = "Refresh Data Now";
var C_SCHEDULE_AUTO_REFRESH_MESSAGE = "Schedule 1 Hourly Automatic Refresh";
var C_STOP_AUTO_REFRESH_MESSAGE = "Stop Automatic Refresh";
var C_JIRA_MENU_TITLE = "Jira";
var C_PROJECT_PREFIX_MESSAGE = "Enter the prefix for your Jira Project. e.g. IMPHEALTH or a list e.g. IMPHEALTH,IMPVSUPP";
var C_PREFIX_FIELD = "prefix";
var C_ENTER_PROJECT_ALERT = "You actually need to enter a project.";
var C_HOST_FIELD_PROMPT = "Enter the host name of your on demand instance e.g. katch-com.atlassian.net";
var C_DEFAULT_HOST = "katch-com.atlassian.net";
var C_HOST_FIELD = "host";
var C_USER_PASSWORD_PROMPT = "Enter your Jira User id and Password in the form User:Password.";
var C_NO_PASSWORD_ALERT = "There is no auth set. boo!";
var C_ISSUE_TYPES_PROPERTY = "issueTypes";
var C_ISSUE_TYPES_DEFAULT = "epic,story,task,bug";
var C_JIRA_CONFIG_SAVED_MESSAGE = "Jira configuration saved successfully.";
var C_NO_LONGER_REFRESH_MESSAGE = "Spreadsheet will no longer refresh automatically.";
var C_YES_REFRESH_MESSAGE = "Spreadsheet will refresh automatically every 1 hours.";
var C_AUTH_TYPE = "Basic ";
var C_JIRA_BACKLOG_SUCCESS_MESSAGE = "Jira backlog successfully imported";
var C_JIRA_PULL_ERROR = "Error pulling data from Jira - aborting now.";

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
      {name: C_JIRA_CONFIG_MESSAGE, functionName: "jiraConfigure"},
      {name: C_REFRESH_DATA_NOW_MESSAGE, functionName: "jiraPullManual"},
      {name: C_SCHEDULE_AUTO_REFRESH_MESSAGE, functionName: "scheduleRefresh"},
      {name: C_STOP_AUTO_REFRESH_MESSAGE, functionName: "removeTriggers"}];
   ss.addMenu(C_JIRA_MENU_TITLE, menuEntries);
 }

function jiraConfigure() {
  
  var ui = SpreadsheetApp.getUi();
  var prefix_result = ui.prompt(C_PROJECT_PREFIX_MESSAGE, ui.ButtonSet.OK_CANCEL);
  
  if (prefix_result.getSelectedButton() == ui.Button.OK) {
    var prefix = result.getResponseText();
    PropertiesService.getUserProperties().setProperty(C_PREFIX_FIELD, prefix.toUpperCase());
  } else if (PropertiesService.getUserProperties().getProperty("digest") == undefined) {
    ui.alert(C_ENTER_PROJECT_ALERT, "project", ui.ButtonSet.OK);
  }
  
  var host_result = ui.prompt(C_HOST_FIELD_PROMPT, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() == ui.Button.OK) {
    var host = result.getResponseText();
    PropertiesService.getUserProperties().setProperty(C_HOST_FIELD, host);
  } else if (PropertiesService.getUserProperties().getProperty(C_HOST_FIELD) == undefined) {
    var host = C_DEFAULT_HOST;
  }
  
  var auth_result = ui.prompt(C_USER_PASSWORD_PROMPT, ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() == ui.Button.OK) {
    var auth = result.getResponseText();
    var x = Utilities.base64Encode(userAndPassword);
    PropertiesService.getUserProperties().setProperty("digest", C_AUTH_TYPE + x);
  } else if (PropertiesService.getUserProperties().getProperty("digest") == undefined) {
    ui.alert(C_NO_PASSWORD_ALERT);
  }
  
  PropertiesService.getUserProperties().setProperty(C_ISSUE_TYPES_PROPERTY, C_ISSUE_TYPES_DEFAULT);
  Browser.msgBox(C_JIRA_CONFIG_SAVED_MESSAGE);
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Browser.msgBox(C_NO_LONGER_REFRESH_MESSAGE);
}

function scheduleRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("jiraPull").timeBased().everyHours(1).create();
  Browser.msgBox(C_YES_REFRESH_MESSAGE);
}

function jiraPullManual() {
  jiraPull();
  Browser.msgBox(C_JIRA_BACKLOG_SUCCESS_MESSAGE);
}

function getFields() {
  return JSON.parse(pullJiraData("field"));
}

function getTickets() {
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  var project_query = "project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix");
  
  while (data.startAt + data.maxResults < data.total) {
    Logger.log("Making request for %s entries", C_MAX_RESULTS);
    if (PropertiesService.getUserProperties().getProperty("prefix").indexOf(',') >= 0) {
      var pa = PropertiesService.getUserProperties().getProperty("prefix").split(',');
      project_query = "(project%20%3D%20" + pa[0];
      for (var i in pa) {
        if (i != 0) {
          project_query = project_query + "%20OR%20project%20%3D%20%22" + pa[i] + "%22";
        }
      }
      project_query = project_query + ")";
    }
    data =  JSON.parse(
        pullJiraData("search?jql=" + project_query + 
                     "%20and%20status%20!%3D%20resolved%20and%20status%20!%3D%20done%20and%20type%20in%20("+ 
                     PropertiesService.getUserProperties().getProperty("issueTypes") + ")%20order%20by%20rank%20&maxResults=" + 
                     C_MAX_RESULTS + "&startAt=" + startAt));
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  }
  return allData;
}

function pullJiraData(path) {
  var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
  var digestfull = PropertiesService.getUserProperties().getProperty("digest");
  
  var headers = { "Accept":"application/json",
                  "Content-Type":"application/json",
                  "method": "GET",
                  "headers": {"Authorization": digestfull},
                  "muteHttpExceptions": true
                };
  
  var resp = UrlFetchApp.fetch(url,headers );
  if (resp.getResponseCode() != 200) {
    Browser.msgBox("Error retrieving data for url" + url + ":" + resp.getContentText());
    return "";
  } else {
    return resp.getContentText();
  }
}

function jiraPull() {
  var allFields = getAllFields();
  var data = getTickets();
  
  if (allFields === "" || data === "") {
    Browser.msgBox(C_JIRA_PULL_ERROR);
    return;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backlog");
  var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
  var y = new Array();
  for (i=0;i<data.issues.length;i++) {
    var d=data.issues[i];
    y.push(getTicket(d,headings,allFields));
  }
  
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backlog");
  var last = ss.getLastRow();
  if (last >= 2) {
    ss.getRange(2, 1, ss.getLastRow()-1,ss.getLastColumn()).clearContent();
  }
  
  if (y.length > 0) {
    ss.getRange(2, 1, data.issues.length,y[0].length).setValues(y);
  }
}

function getAllFields() {
  
  var theFields = getFields();
  var allFields = new Object();
  allFields.ids = new Array();
  allFields.names = new Array();
  
  for (var i = 0; i < theFields.length; i++) {
    allFields.ids.push(theFields[i].id);
    allFields.names.push(theFields[i].name.toLowerCase());
  }
  return allFields;
}

function getTicket(data,headings,fields) {
  var story = [];
  for (var i = 0;i < headings.length;i++) {
    if (headings[i] !== "") {
      story.push(getDataForHeading(data,headings[i].toLowerCase(),fields));
    }
  }
  return story;
}

function getDataForHeading(data, heading, fields) {
  
  if (data.hasOwnProperty(heading)) {
    if (heading == "status" && data.hasOwnProperty(heading)) {
      return data[heading].name;
    }
    if (heading == "assignee" && data.hasOwnProperty(heading)) {
      if (data[heading] != null) {
        return data[heading].name;
      }
    }
    return data[heading];
  } else if (data.fields.hasOwnProperty(heading)) {
    if (heading == "status" && data.fields.hasOwnProperty(heading)) {
      return data.fields[heading].name;
    }
    if (heading == "assignee" && data.fields.hasOwnProperty(heading)) {
      if (data.fields[heading] != null) {
        return data.fields[heading].name;
      }
    }
    return data.fields[heading];
  }
  
  var fieldName = getFieldName(heading,fields);
  
  if (fieldName !== "") {
    if (data.hasOwnProperty(fieldName)) {
      return data[fieldName];
    } else if (data.fields.hasOwnProperty(fieldName)) {
      return data.fields[fieldName];
    }
  }
  
  var splitName = heading.split(" ");
  
  if (splitName.length == 2) {
    if (data.fields.hasOwnProperty(splitName[0]) ) {
      if (data.fields[splitName[0]] && data.fields[splitName[0]].hasOwnProperty(splitName[1])) {
        return data.fields[splitName[0]][splitName[1]];
      }
      return "";
    }
  }
  return "Could not find value for " + heading;
}

function getFieldName(heading,fields) {
  var index = fields.names.indexOf(heading);
  if ( index > -1) {
     return fields.ids[index];
  }
  return "";
}
