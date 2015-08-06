/**
 * organizeEmail 
 * 
 * specifications:
 * ***************
 *
 * considers only mails having a label under "ML/"
 * - archive inbox threads with a label "ML/*" older thant 2d
 * - remove archived threads with a label "ML/*" older thant 60d and having not "keep" label
 * - send a report of the execution 
 *   "!" suffix in a report line means that the action hasn't been perform (dry run to test behavior)
 *
 * How to use it:
 * **************
 * - copy this script onto your https://script.google.com/
 * - adapt first parameters to fit your needs
 * - schedule the script each night
 * - set "archiveMail" or "deleteMail" to true if and only if you have tested before.
 *
 **/
function organizeEmail____________MainEntryPoint() {
  // general parameters
  var labelKeep = "keep";
  var labelML = "ML/";
  var minArchiveAge = '2d';
  var minDeleteAge  = '60d';
  var maxThreads = 500;
  var reportRecipient = 'you.email@yoyo.fr';
  var reportTitle = '#GGScript organizeEmail report ';

  // dry run  : set to false // set it to true only after tests :)
  var archiveMail = true; 
  // dry run  : set to false // set it to true only after tests :)
  var deleteMail  = false;

  
  var organizeReport = [];

  // get threads to archive
  var toArchiveMLThreads = getToArchiveMLThreads(minArchiveAge, maxThreads, labelML);
  // archive them
  archiveMLThreads(organizeReport, archiveMail, toArchiveMLThreads);
  // get threads to delete
  var toDeleteMLThreads = getToDeleteMLThreads(minDeleteAge, maxThreads, labelML, labelKeep);
  // delete them
  deleteMLThreads(organizeReport, deleteMail, toDeleteMLThreads);
  // send report by email
  sendReport(reportRecipient, reportTitle, organizeReport);
}

function isInML(labels, labelML) {
    var i = labels.length;
    while (i--) {
       if (labels[i].getName().indexOf(labelML) === 0) {
           return true;
       }
    }
    return false;
}

function toString(thread) {
  return  thread.getLastMessageDate() + " | " + thread.getFirstMessageSubject() + " [" + thread.getMessageCount() + "]";
}

function getToArchiveMLThreads(minArchiveAge, maxThreads, labelML) {
  var oldInboxMail = "label:inbox older_than:" + minArchiveAge;
  var threads = GmailApp.search(oldInboxMail, 0, maxThreads);
  var rez = [];
  for (var i=0; i < threads.length ; i++ ) {
    var t = threads[i];
    var labels = t.getLabels();
    if (isInML(labels, labelML)) {
      rez.push(t);
    }
  }
  Logger.log("\ngetToArchiveMLThreads filter:" + oldInboxMail + " result:" + rez.length);
  return rez;
}

function archiveMLThreads(organizeReport, archiveMail, toArchiveMLThreads) {
  if ((toArchiveMLThreads == undefined) || (toArchiveMLThreads.length < 1)) {
    return;
  }
  organizeReport.push("archiveMLThreads " + toArchiveMLThreads.length);
  for (var i=0; i < toArchiveMLThreads.length ; i++ ) {
    var t = toArchiveMLThreads[i];
    if (archiveMail) {
      organizeReport.push(" ARCHIVED >>> " + toString(t));
      t.moveToArchive();
    } else {
      organizeReport.push("!ARCHIVED >>> " + toString(t));
    }
  }
}

function getToDeleteMLThreads(minDeleteAge, maxThreads, labelML, labelKeep) {
  var oldArchivedMail = "!label:" + labelKeep+ " !label:inbox older_than:" + minDeleteAge;
  var threads = GmailApp.search(oldArchivedMail, 0, maxThreads);
  var rez = [];
  for (var i=0; i < threads.length ; i++ ) {
    var t = threads[i];
    var labels = t.getLabels();
    if (isInML(labels, labelML)) {
      rez.push(t);
    }
  }
  Logger.log("\ngetToDeleteMLThreads filter:" + oldArchivedMail + " result:" + rez.length);
  return rez;
}

function deleteMLThreads(organizeReport, deleteMail, toDeleteMLThreads) {
    if ((toDeleteMLThreads == undefined) || (toDeleteMLThreads.length < 1)) {
      return;
    }
    organizeReport.push("deleteMLThreads " + toDeleteMLThreads.length);
    for (var i=0; i < toDeleteMLThreads.length ; i++ ) {
    var t = toDeleteMLThreads[i];
    if (deleteMail) {
      organizeReport.push(" REMOVED >>> " + toString(t));
      t.moveToTrash();
    } else {
      organizeReport.push("!REMOVED >>> " + toString(t));
    }
  }
}

function sendReport(reportRecipient, reportTitle, organizeReport) {
  if ((organizeReport == undefined) || (organizeReport.length < 1)) {
    return;
  }
  var reportLogs = "Execution Script's logs :\n" + Logger.getLog();
  var reportText = "Organize Email Report\n\n";
  var reportHtml = "Organize Email Report\n\n<ul>";
  for (var j=0; j < organizeReport.length ; j++ ) {
    var line = organizeReport[j];
    reportText += line + "\n";
    reportHtml += "<li>" + line + "</li>\n";
  }
  reportHtml += "</ul>";
  
  if (reportLogs) {
    reportText += "\n-------------------------\n\n"+reportLogs;
    reportHtml += "<hr/>\n<pre>"+reportLogs + "</pre>";
  }
  
  GmailApp.sendEmail(reportRecipient, reportTitle, reportText, {
     htmlBody: reportHtml
 });
}