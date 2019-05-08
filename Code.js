var scriptName = "Purge Labels";
var userProperties = PropertiesService.getUserProperties();
var labels_to_purge = userProperties.getProperty("labels_to_purge") || '';
var labels_never_purge = userProperties.getProperty("labels_never_purge") || '';
var purge_checkFrequency_HOUR = userProperties.getProperty("purge_checkFrequency_HOUR") || 1;
var status = userProperties.getProperty("status") || "disabled";
var user_email = Session.getEffectiveUser().getEmail();

global.doGet = doGet;
global.test = test;
global.getSettings = getSettings;
global.deleteAllTriggers = deleteAllTriggers;
global.purgeGmail = purgeGmail;
global.labelPurge = labelPurge;

function getSettings() {
  Logger.log(userProperties.getProperty("labels_to_purge"));
  Logger.log(userProperties.getProperty("purge_checkFrequency_HOUR"));
  Logger.log(userProperties.getProperty("status"));
}

function doGet(e) {
  if (e.parameter.setup) { // SETUP
    deleteAllTriggers();
    if (status === "enabled") {
      ScriptApp.newTrigger("purgeGmail").timeBased().atHour(purge_checkFrequency_HOUR).everyDays(1).create();
    }

    var content = "<p>" + scriptName + " has been installed on your email " + user_email + ". "
      + "It is currently set to your specified labels in settings every 1 AM.</p>"
      + '<p>You can change these settings by clicking the HOPLA Tools extension icon or HOPLA Tools Settings on gmail.</p>';

    return HtmlService.createHtmlOutput(content);
  } else if (e.parameter.savesettings) { // SET VARIABLES
    var settings = JSON.parse(e.parameter.savesettings);
    userProperties.setProperty("labels_to_purge", JSON.stringify(settings.labels_to_purge));
    userProperties.setProperty("labels_never_purge", JSON.stringify(settings.labels_never_purge));
    userProperties.setProperty("purge_checkFrequency_HOUR", settings.purge_checkFrequency_HOUR || 1);
    userProperties.setProperty("status", settings.status);

    labels_to_purge = userProperties.getProperty("labels_to_purge");
    labels_never_purge = userProperties.getProperty("labels_never_purge");
    purge_checkFrequency_HOUR = userProperties.getProperty("purge_checkFrequency_HOUR");
    status = userProperties.getProperty("status") || "disabled";

    deleteAllTriggers();
    if (status === "enabled") ScriptApp.newTrigger("purgeGmail").timeBased().atHour(purge_checkFrequency_HOUR).everyDays(1).create();


    return ContentService.createTextOutput("Purge settings has been saved.");
  } else if (e.parameter.purge_trigger) { // DO IT NOW
    var msg = purgeGmail() || "Purge has finished.";
    return ContentService.createTextOutput(msg);
  } else if (e.parameter.purge_enable) { // ENABLE
    userProperties.setProperty("status", "enabled");
    deleteAllTriggers();
    ScriptApp.newTrigger("purgeGmail").timeBased().atHour(purge_checkFrequency_HOUR).everyDays(1).create();

    return ContentService.createTextOutput("Triggers has been enabled.");
  } else if (e.parameter.purge_disable) { // DISABLE
    userProperties.setProperty("status", "enabled");
    deleteAllTriggers();
    return ContentService.createTextOutput("Triggers has been disabled.");
  } else if (e.parameter.purge_getVariables) { // GET VARIABLES
    var triggers = ScriptApp.getProjectTriggers();
    var status = triggers.length > 0 ? 'enabled' : 'disabled';
    var purge_checkFrequency_HOUR = userProperties.getProperty("purge_checkFrequency_HOUR") || 1;
    purge_checkFrequency_HOUR = parseInt(purge_checkFrequency_HOUR, 10);
    var labels_to_purge = _JSONPARSE(userProperties.getProperty("labels_to_purge"));
    var labels_never_purge = _JSONPARSE(userProperties.getProperty("labels_never_purge"));
    var resjson = {
      labels_to_purge,
      labels_never_purge,
      purge_checkFrequency_HOUR,
      status
    };
    return ContentService.createTextOutput(JSON.stringify(resjson));
  } else { // NO PARAMETERS
    // use an externally hosted stylesheet
    var style = '<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">';


    var greeting = scriptName;
    var name = "has been installed";
    var heading = HtmlService.createTemplate('<h1><?= greeting ?> <?= name ?>!</h1>');
    heading.greeting = greeting;
    heading.name = name;

    deleteAllTriggers();
    if (status === 'enabled') {
      ScriptApp.newTrigger("purgeGmail").timeBased().atHour(purge_checkFrequency_HOUR).everyDays(1).create();
    }

    var content = "<p>" + scriptName + " has been installed on your email " + user_email + ". "
      + "It is currently set to your specified labels in settings every 1 AM.</p>"
      + '<p>You can change these settings by clicking the HOPLA Tools extension icon or HOPLA Tools Settings on gmail.</p>';

    var HTMLOutput = HtmlService.createHtmlOutput();
    HTMLOutput.append(style);
    HTMLOutput.append(heading.evaluate().getContent());
    HTMLOutput.append(content);

    return HTMLOutput;
  }
}

function _JSONPARSE(o) { // RETURNS EMPTY TXT WHEN NOT JSON
  let ret = '';
  try {
    ret = JSON.parse(o);
  } catch (e) {
    ret = '';
  }
  return ret;
}

function deleteAllTriggers() {
  // DELETE ALL TRIGGERS
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  // DELETE ALL TRIGGERS***
}

function purgeGmail() {
  var deletedThreads = 0;
  var processedLabels = 0;

  try {
    var universal_never_purge = _JSONPARSE(userProperties.getProperty("labels_never_purge"));
    var labels_to_purge = _JSONPARSE(userProperties.getProperty("labels_to_purge"));
    if (labels_to_purge) {
      for (var i in labels_to_purge) {
        var oPurge = labels_to_purge[i];
        Logger.log("Purging [" + oPurge.purge_label + "]");
        var deleted = labelPurge(oPurge, universal_never_purge);
        if (deleted) deletedThreads += deleted;
        processedLabels += 1;
      }

      return "Processed " + processedLabels + " label(s). Deleted " + deletedThreads + " thread(s).";
    }
  } catch (e) {
    Logger.log(e);
  }
}

function test() {
  var deletedThreads = 0;
  var processedLabels = 0;

  try {
    var all_purge = JSON.parse('{"labels_to_purge":[{"purge_label":"HOPLA Emails","older_than":"1","labels_never_purge":["ignore"],"purge_important":true,"purge_starred":false},{"purge_label":"Spam","older_than":"2","labels_never_purge":["ignore"],"purge_important":false,"purge_starred":true}],"labels_never_purge":["Auto Replies"],"purge_checkFrequency_HOUR":5,"status":"disabled"}');
    var labels_to_purge = all_purge.labels_to_purge;
    if (labels_to_purge && labels_to_purge.length) {
      for (var i in labels_to_purge) {
        var oPurge = labels_to_purge[i];
        Logger.log("Purging [" + oPurge.purge_label + "]");
        var deleted = labelPurge(oPurge);
        if (deleted) deletedThreads += deleted;
        processedLabels += 1;
      }
      // for (var key in labels) {
      //   Logger.log("Purging [" + key + "]");
      //   var deleted = labelPurge(key, labels[key]);
      //   if (deleted) deletedThreads += deleted;
      //   processedLabels += 1;
      // }

      return "Processed " + processedLabels + " label(s). Deleted " + deletedThreads + " thread(s).";
    }
  } catch (e) {
    Logger.log(e);
  }
}

// function labelPurge(label, purgeafter) {
function labelPurge(oPurge, aNeverPurge) {
  var label = oPurge.purge_label; // target label
  var purgeafter = oPurge.older_than; // days
  if (label === "") {
    return 1;
  }
  var age = new Date();
  age.setDate(age.getDate() - purgeafter);

  var purge = Utilities.formatDate(age, Session.getTimeZone(), "yyyy-MM-dd");


  var search = "label:" + label + " before:" + purge;

  if (oPurge.labels_never_purge && oPurge.labels_never_purge.length) {
    oPurge.labels_never_purge.forEach(function (value) {
      search += ' -label:' + value;
    });
  }

  if (aNeverPurge && aNeverPurge.length) {
    aNeverPurge.forEach(function (value) {
      search += ' -label:' + value;
    });
  }

  if (!oPurge.purge_important) search += ' -label:important';
  if (!oPurge.purge_starred) search += ' -label:starred';

  Logger.log("Your query: " + search);
  // return;
  // This will create a simple Gmail search
  // query like label:Newsletters before:10/12/2012
  // eslint-disable-next-line no-unreachable
  var deletedThreads = 0;
  try {
    // We are processing 100 messages in a batch to prevent script errors.
    // Else it may throw Exceed Maximum Execution Time exception in Apps Script

    var threads = GmailApp.search(search, 0, 100);


    // For large batches, create another time-based trigger that will
    // activate the auto-purge process after 'n' minutes.
    Logger.log("Found: " + threads.length);
    if (threads.length === 100) {
      ScriptApp.newTrigger("purgeGmail")
        .timeBased()
        .at(new Date((new Date()).getTime() + 1000 * 60 * 10))
        .create();
    }

    // An email thread may have multiple messages and the timestamp of
    // individual messages can be different.

    for (var i = 0; i < threads.length; i++) {
      var messages = GmailApp.getMessagesForThread(threads[i]);
      for (var j = 0; j < messages.length; j++) {
        var email = messages[j];
        if (email.getDate() < age) {
          Logger.log("Delete [" + email.getSubject() + "]");
          email.moveToTrash();
          deletedThreads += 1;
        } else {
          Logger.log("Skip [" + email.getSubject() + "]");
        }
      }
    }

    // If the script fails for some reason or catches an exception,
    // it will simply defer auto-purge until the next day.
  } catch (e) {
    Logger.log("Error = " + e);
  }

  return deletedThreads;
}