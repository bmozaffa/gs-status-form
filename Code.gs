const documentLinks = getDocumentLinks();
const globalLinks = getGlobalLinks();
const kerberosMap = getKerberosMap(globalLinks.roverSheetId);

function getDocumentLinks() {
  let docsLinks = new Map()
  let spreadsheet = SpreadsheetApp.openById("1RKF97_z2ruAgUJvxcoArlVt3JEtoFjfLB-spycO3GHw");
  let sheetNames = ["Initiatives", "Managers", "Templates"];
  for (let sheetName of sheetNames) {
    let map = new Map();
    let sheet = spreadsheet.getSheetByName(sheetName);
    for (let row = 2; row <= sheet.getLastRow(); row++) {
      let key = sheet.getRange(row, 1, 1, 1).getValue();
      if (key.length > 0) {
        map.set(key, sheet.getRange(row, 2, 1, 1).getValue());
      }
      docsLinks.set(sheetName, map);
    }
  }
  return docsLinks;
}

function getGlobalLinks() {
  let links = {};
  links.statusFormId = "14g3I22FRYFV9kbE9csTdlbrKK3XbnHXo8rsMDg-Z_HI";
  links.roverSheetId = "1i7y_tFpeO68SetmsU2t-C6LsFETuZtkJGY5AVZ2PHW8";
  links.statusEmailsId = "1pke_nZSAwVFL9iIx-HKgaaZU4lMmr5aParHgN9wGdXE";
  links.rosterSheetId = "1ARSzzTSBtiOhPfo8agZe9TvI1tM4WQFEvENzZsD3feU";
  return links;
}

const htmlFormatting = {
  lineBreak: "<br/>",
  paragraphBreak: "<p/>",
  italicStart: "<i>",
  italicEnd: "</i>",
  boldStart: "<b>",
  boldEnd: "</b>",
};

const textFormatting = {
  lineBreak: "\n",
  paragraphBreak: "\n\n\n",
  italicStart: "",
  italicEnd: "",
  boldStart: "**",
  boldEnd: "**",
};

function doGet(e) {
  let command = e.parameter.command;
  Logger.log("Received a web request with the following parameters" + e.parameters)
  let response;
  if (command === "missing") {
    response = "";
    let missingHierarchy = getMissingStatusReport();
    missingHierarchy.forEach((associates, manager) => {
      response += "<p>" + manager + ": ";
      response += associates.join(", ");
      response += " has/have not entered status";
    });
    if (response.length === 0) {
      response = "No missing status entries this week";
    }
  } else if (command === "generate") {
    response = compileStatus(e.parameter.docId);
  } else if (command === "archive") {
    let archived = archiveReports(e.parameter.days);
    response = "Successfully archived " + archived + " form submissions";
  } else if (command === "send-drafts") {
    let sentMessages = sendDraftEmails();
    response = "Successfully sent " + sentMessages + " emails from saved drafts";
  } else if (command === "mismatch") {
    response = getMismatchResponse(e.parameter.manager, htmlFormatting);
  } else {
    response = "Provide the command (missing, generate, etc) as a request parameter: https://script.google.com/..../exec?command=generate";
  }
  return HtmlService.createHtmlOutput(response);
}

function getMismatchResponse(mgrKerberos, format) {
  if (mgrKerberos) {
    return getMismatchAssignmentByManager(mgrKerberos, format);
  } else {
    let response = "";
    for (const manager of documentLinks.get("Managers").keys()) {
      response += getAssociateName(manager) + ":" + format.lineBreak;
      response += getMismatchAssignmentByManager(manager, format);
      response += format.paragraphBreak;
    }
    return response;
  }
}

function getMismatchAssignmentByManager(mgrKerberos, format) {
  const fullMismatch = compareAssignments();
  const filteredMismatch = [new Map(), new Map(), new Map()];
  for (let i = 0; i < 3; i++) {
    for (const kerberos of fullMismatch[i].keys()) {
      if (getAssociateManager(kerberos) === mgrKerberos) {
        filteredMismatch[i].set(kerberos, fullMismatch[i].get(kerberos));
      }
    }
  }

  let response = "";
  for (const kerberos of filteredMismatch[0].keys()) {
    let associate = getAssociateName(kerberos);
    let activities = filteredMismatch[0].get(kerberos).join(" / ");
    response += format.italicStart + associate + format.italicEnd + " reported working on " + format.boldStart + activities + format.boldEnd + " this week. The roster does not contain this assignment. This is unusual and likely reflects a missing roster assignment, although it may be ignored if it's temporary and not a significant contribution." + format.lineBreak;
  }
  if (filteredMismatch[0].size > 0) {
    response += format.paragraphBreak;
  }
  for (const kerberos of filteredMismatch[1].keys()) {
    let associate = getAssociateName(kerberos);
    response += format.boldStart + associate + format.boldEnd + " is not assigned any work and did not report any activity.This is unsual and likely reflects a missing roster assignment, as well as a lack of activity report." + format.lineBreak;
  }
  if (filteredMismatch[1].size > 0) {
    response += format.paragraphBreak;
  }
  if (filteredMismatch[2].size > 0) {
    response += "The following associates did not report any activity in the specified assignments this week. This is not unusual, especially if they are assigned multiple projects, but please review and investigate if this recurs!" + format.lineBreak;
  }
  for (const kerberos of filteredMismatch[2].keys()) {
    let associate = getAssociateName(kerberos);
    response += format.italicStart + associate + format.italicEnd + ": " + filteredMismatch[2].get(kerberos).join(", ") + format.lineBreak;
  }
  if (filteredMismatch[2].size > 0) {
    response += format.paragraphBreak;
  }
  return response;
}

function notifyMismatchAssignment() {
  for (const manager of documentLinks.get("Managers").keys()) {
    notifyMismatchAssignmentByManager(manager);
  }
}

function notifyMismatchAssignmentByManager(manager) {
  if (isPaused('notifyMismatchAssignment')) {
    return;
  }
  const report = getMismatchAssignmentByManager(manager, htmlFormatting);
  if (report.length === 0) {
    //No mismatch, so don't send an email
    return;
  }
  const me = "Babak Mozaffari <Babak@redhat.com>";
  const to = getAssociateName(manager) + " <" + getAssociateEmail(manager) + ">";
  const subject = "Mismatch between your associates' assignments and reported activity";
  const bodyBase = "This is an automated message to notify you of mismatches detected for your team<br/><br/>";
  const signature = getSignature();
  const htmlBody = bodyBase + report + signature;
  const options = {
    cc: me,
    from: me,
    htmlBody: htmlBody
  };
  GmailApp.sendEmail(to, subject, "", options);
}

function getMissingStatusReport() {
  let allStatusDocumentIds = Array.from(documentLinks.get("Templates").keys())
  let missing = getMissingStatus(readResponseObjects(globalLinks.statusFormId), allStatusDocumentIds);
  let missingHierarchy = new Map();
  missing.statusRequired.forEach(kerberos => {
    let associateName = getAssociateName(kerberos).split(" ")[0];
    let managerUID = getAssociateManager(kerberos);
    let managerName = getAssociateName(managerUID).split(" ")[0];
    let associates = getMapArray(missingHierarchy, managerName);
    associates.push(associateName);
    Logger.log(">> %s", associateName);
  });
  return missingHierarchy;
}

function notifyMissingStatus() {
  if (isPaused('notifyMissingStatus')) {
    return;
  }
  let allStatusDocumentIds = Array.from(documentLinks.get("Templates").keys())
  let missing = getMissingStatus(readResponseObjects(globalLinks.statusFormId), allStatusDocumentIds);

  let subjectBase = "Missing weekly status report for ";
  let body = "This is an automated message to notify you that a weekly status report was not detected for the week.\n\nPlease submit your status report promptly. Kindly ignore this message if you are off work and a status entry is not expected.";
  missing.statusRequired.forEach(kerberos => {
    let managerId = getAssociateManager(kerberos);
    let to = getAssociateEmail(kerberos);
    let cc = getAssociateEmail(managerId);
    let subject = subjectBase + getAssociateName(kerberos).split(" ")[0];
    GmailApp.sendEmail(to, subject, body, {cc: cc});
    Logger.log("%s who reports to %s is missing status, sent them an email!", getAssociateName(kerberos), getAssociateName(managerId));
    Logger.log("Sending an email to %s with subject line [%s] and body [%s] and copying %s", to, subject, body, cc);
  });
}

function getMissingStatus(responseObjects, statusDocIds) {
  //Figure out who need to submit status report
  let statusRequired = new Set();
  const excludedStatus = ["Director, Software Engineering_Global", "Senior Manager, Software Engineering_Global", "Manager, Software Engineering", "Associate Manager, Software Engineering", "Senior Manager, SAP Alliance Technology Team"];
  for (let kerberos of getAllUsers()) {
    const title = getAssociateTitle(kerberos);
    if (excludedStatus.includes(title)) {
      //Exclude managers from status entry
    } else if (title.includes("Director")) {
      //Exclude directors from missing entry
    } else if (title.includes("Intern")) {
      //Exclude interns from missing entry
    } else if (kerberos) {
      //Add entry only if non-empty
      statusRequired.add(kerberos);
    }
  }

  //Remove from the list those who have reported status
  for (let responseObj of responseObjects) {
    statusRequired.delete(responseObj.kerberos);
  }

  //Only report missing status for associates relevant to provided status documents
  let userAssignmentMap = getUserAssignmentMap();
  nextUser: for (let kerberos of statusRequired) {
    //Check if user's manager owns this status
    let managerUID = getAssociateManager(kerberos);
    if (statusDocIds.includes(documentLinks.get("Managers").get(managerUID))) {
      continue;
    }

    //Check if user's assignments belong in this status doc
    if (userAssignmentMap.has(kerberos)) {
      for (let assignment of userAssignmentMap.get(kerberos)) {
        let initiative = assignment.split('\.')[0]
        if (statusDocIds.includes(documentLinks.get("Initiatives").get(initiative))) {
          continue nextUser;
        }
      }
    }
    //If there was no match based on manager or initiative, the user status is not required for THIS status doc
    statusRequired.delete(kerberos);
  }
  //Remove from the list those who are on PTO as per their personal calendars
  for (let kerberos of statusRequired) {
    if (isOnPTO(kerberos.concat("@redhat.com"))) {
      Logger.log("The personal calendar of %s shows as OOO, so won't mark them as missing status", getAssociateName(kerberos));
      statusRequired.delete(kerberos);
    }
  }

  Logger.log("Missing status for %s people", statusRequired.size);
  let missingStatus = {};
  missingStatus.statusRequired = statusRequired;
  return missingStatus;
}

function archiveReports(days) {
  let form = FormApp.openById(globalLinks.statusFormId);
  let cutoff = new Date();
  if (!days) {
    days = 7;
  }
  cutoff.setDate(cutoff.getDate() - days);
  Logger.log("Cutoff is " + cutoff);
  responses = form.getResponses();
  let count = 0;
  let total = responses.length;
  responses.forEach(response => {
    if (response.getTimestamp() < cutoff) {
      form.deleteResponse(response.getId());
      count++;
    } else {
    }
  });
  Logger.log("Archived " + count + " responses out of " + total);
  return count;
}

function logStatus() {
  let responseObjects = readResponseObjects(globalLinks.statusFormId);
  let map = getStatusMap(responseObjects);
  map.forEach((statusList, category) => {
    if ( hasNoStatus(category) ) {
      statusList.forEach(responseObject => {
        Logger.log(responseObject.kerberos + " has PTO / Learning / No Status");
      });
    } else {
      statusList.forEach(responseObject => {
        let status = getStatusText(responseObject).map(status => {
          return status.text;
        }).join("");
        Logger.log(status);
      });
    }
  });
}

function testCompileStatus() {
  // let templateId = "1I871rNsYejPBqrdX3eCsBFso8zZC34NzxeJrOKWYwNI";
  let templateId = "1KNrtZgdzApR6-NfxY0VwrzGk3u4WRmB2JsTAE3LvgaU";
  let statusDocId = "1ke5E5Qe0CVGX9-DI9O8q_ZmcEhEfpxWm1q5lDClOGLU";
  copyTemplate(templateId, statusDocId);

  let statusMap = new Map();
  {
    let statusArray = getMapArray(statusMap, "FSI.Cross cutting concerns");
    let responseObj = {};
    responseObj.timestamp = "Right now";
    responseObj.kerberos = "bmozaffa";
    responseObj.initiative = "FSI";
    responseObj.effort = "Cross cutting concerns";
    responseObj.epic = "";
    responseObj.status = "Did some generic process stuff";
    statusArray.push(responseObj);
  }
  {
    let statusArray = getMapArray(statusMap, "FSI.Temenos");
    let responseObj = {};
    responseObj.timestamp = "Right now";
    responseObj.kerberos = "bmozaffa";
    responseObj.initiative = "FSI";
    responseObj.effort = "Temenos";
    responseObj.epic = "";
    responseObj.status = "Did some Temenos work";
    statusArray.push(responseObj);
  }
  insertStatus(statusDocId, statusMap, 1);
}

function compileStatus(docId) {
  if (isPaused('compileStatus')) {
    return;
  }
  let statusDocIds;
  if (docId) {
    statusDocIds = [docId];
  } else {
    statusDocIds = documentLinks.get("Templates").keys();
  }
  let statusDocs = new Set();
  let lastStatusTimestamp = getLastStatusTimestamp(globalLinks.statusFormId);
  for (let statusDocId of statusDocIds) {
    if (needsUpdate(lastStatusTimestamp, statusDocId)) {
      copyTemplate(documentLinks.get("Templates").get(statusDocId), statusDocId);
      statusDocs.add(statusDocId);
    }
  }

  if (statusDocs.size === 0) {
    return "There is no status entry since document was generated";
  } else {
    let allResponseObjects = readResponseObjects(globalLinks.statusFormId);
    for (let statusDocId of statusDocs) {
      let responseObjects = [];
      for (let responseObj of allResponseObjects) {
        if (matchesStatusDoc(responseObj, statusDocId)) {
          responseObjects.push(responseObj);
        }
      }
      let statusMap = getStatusMap(responseObjects);
      insertStatus(statusDocId, statusMap, responseObjects.length);
    }
    let form = FormApp.openById(globalLinks.statusFormId);
    return "Successfully generated status reports based on " + form.getResponses().length + " form submissions";
  }
}

function matchesStatusDoc(responseObject, statusDocId) {
  let initiativeDocId = documentLinks.get("Initiatives").get(responseObject.initiative);
  if (initiativeDocId) {
    return initiativeDocId === statusDocId;
  } else {
    let managerUID = getAssociateManager(responseObject.kerberos);
    return documentLinks.get("Managers").get(managerUID) === statusDocId;
  }
}

function getLastStatusTimestamp(formId) {
  let form = FormApp.openById(formId);
  let lastStatusTimestamp;
  form.getResponses().forEach(response => {
    if (lastStatusTimestamp) {
      if (response.getTimestamp() > lastStatusTimestamp) {
        lastStatusTimestamp = response.getTimestamp();
      }
    } else {
      lastStatusTimestamp = response.getTimestamp();
    }
  });
  Logger.log("Got latest status entry dated " + lastStatusTimestamp);
  return lastStatusTimestamp;
}

function needsUpdate(lastStatusTimestamp, statusDocId) {
  if (!lastStatusTimestamp) {
    return false;
  }

  let body = DocumentApp.openById(statusDocId).getBody();
  let lastUpdateMessage = "This document is no longer auto-generated. It was last generated at ";
  let paragraph = body.getParagraphs()[0].getText();
  if (isNaN(Date.parse(paragraph.substring(lastUpdateMessage.length)))) {
    console.error("Failed to parse last update time");
    return true;
  } else {
    let lastUpdate = new Date(paragraph.substring(lastUpdateMessage.length));
    Logger.log("The doc was last updated on " + lastUpdate);
    if (lastStatusTimestamp > lastUpdate) {
      return true;
    } else {
      //Cover edge cases of errors updating the doc timestamp without updating its content
      if (body.findText("%[A-Za-z0-9\.\/\u0020\(\)-]+%")) {
        return true;
      } else {
        return false;
      }
    }
  }
}

function copyTemplates() {
  let copiedDocs = new Set();
  let sheetNames = ["Initiative", "Manager"];
  for (sheetName of sheetNames) {
    for (let links of documentLinks.get(sheetName).values()) {
      if (!copiedDocs.has(links.statusDocId)) {
        copyTemplate(links.templateDocId, links.statusDocId);
        copiedDocs.add(links.statusDocId);
      }
    }
  }
}

function copyTemplate(templateDocId, statusDocId) {
  let template = DocumentApp.openById(templateDocId).getBody().copy();
  let doc;
  try {
    doc = DocumentApp.openById(statusDocId);
  } catch (err) {
    try {
      console.warn("Failed to open document, will try again", err);
      doc = DocumentApp.openById(statusDocId);
    } catch (err2) {
      console.error("Failed to open document", err2);
      throw err2;
    }
  }
  let body = doc.getBody();
  body.appendPageBreak();
  body.clear();
  let note = body.insertParagraph(0, "");
  let lastUpdateMessage = "This document is no longer auto-generated. It was last generated at ";
  let text = note.appendText(lastUpdateMessage);
  text.setFontSize(16);
  let time = note.appendText(new Date().toUTCString());
  time.setBold(true);
  let refresh = body.insertParagraph(1, "\n");
  let click = refresh.appendText("Click ");
  click.setFontSize(16);
  click.setBold(false);
  let link = refresh.appendText("here");
  refresh.appendText(" if you think more status entries are available and wish to refresh this document.")
  link.setLinkUrl("https://script.google.com/a/macros/redhat.com/s/AKfycbwGck8vr-ZNHVCKRg8-y1p5hfVt4Iognzm4zZoOd0WrsORfzLOKNVHk4te0-qOnHCWU/exec?command=generate&docId=" + statusDocId);

  let totalElements = template.getNumChildren();
  for (let index = 0; index < totalElements; index++) {
    let element = template.getChild(index);
    let type = element.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH)
      body.appendParagraph(element.copy());
    else if (type === DocumentApp.ElementType.TABLE)
      body.appendTable(element.copy());
    else if (type === DocumentApp.ElementType.LIST_ITEM) {
      let inserted = body.appendListItem(element.getText());
      inserted.setAttributes(element.getAttributes());
      inserted.setListId(element).setGlyphType(element.getGlyphType()).setNestingLevel(element.getNestingLevel());
    } else
      throw new Error("Unknown element type: " + type);
  }
  doc.saveAndClose();
  doc = DocumentApp.openById(statusDocId);
  body = doc.getBody();
  body.getListItems().forEach(li => {
    if (li.findText("%[A-Za-z0-9\.\/\u0020\(\)-]+%")) {
      let attrs = li.getAttributes();
      attrs.FONT_SIZE = 12;
      li.setAttributes(attrs);
      li.setBold(false);
      li.setItalic(false);
    }
    if (li.getNestingLevel() === 0) {
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
    } else if (li.getNestingLevel() === 1) {
      li.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
    } else if (li.getNestingLevel() === 2) {
      li.setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
    }
  });
  body.getParagraphs().forEach(paragraph => {
    if (paragraph.getType() === DocumentApp.ElementType.PARAGRAPH) {
      if (paragraph.getIndentStart() === null)
        paragraph.setIndentStart(0);
      if (paragraph.getIndentFirstLine() === null)
        paragraph.setIndentFirstLine(0);
    }
  });
  doc.saveAndClose();
}

function getStatusMap(responseObjects) {
  let statusMap = new Map();
  responseObjects.forEach(responseObject => {
    let mapKey;
    if (responseObject.effort) {
      mapKey = responseObject.initiative + "." + responseObject.effort;
    } else {
      mapKey = responseObject.initiative;
    }
    Logger.log("Status reported for %s", mapKey);
    let statusArray = getMapArray(statusMap, mapKey);
    statusArray.push(responseObject);
  });
  return statusMap;
}

function readResponseObjects(formId) {
  let form = FormApp.openById(formId);
  let responses = form.getResponses();
  // form.getItems().forEach(value => {
  //   Logger.log("item is %s", value.getTitle());
  // })
  let responseObjects = [];
  // Logger.log("Found %s status entries", responses.length);
  responses.forEach(response => {
    let responseObj = {};
    responseObj.timestamp = response.getTimestamp();
    responseObj.kerberos = response.getRespondentEmail().split('@')[0];
    let answers = response.getItemResponses()
    responseObj.initiative = answers[0].getResponse();
    if (answers.length >= 4) {
      responseObj.effort = answers[1].getResponse();
      responseObj.epic = answers[2].getResponse();
      responseObj.status = answers[3].getResponse();
    } else if (answers.length >= 3) {
      responseObj.epic = answers[1].getResponse();
      responseObj.status = answers[2].getResponse();
    }
    responseObjects.push(responseObj);
  });
  return responseObjects;
}

function getMapArray(map, key) {
  let array = map.get(key);
  if (!array) {
    array = [];
    map.set(key, array);
    // Logger.log("There was no array for %s, but one has been created now", key);
  }
  return array;
}

function insertStatus(statusDocId, statusMap, responseCount) {
  let reportedCount = 0;
  let doc;
  try {
    doc = DocumentApp.openById(statusDocId);
  } catch (err) {
    try {
      console.warn("Failed to open document, will try again", err);
      doc = DocumentApp.openById(statusDocId);
    } catch (err2) {
      console.error("Failed to open document", err2);
      throw err2;
    }
  }
  let body = doc.getBody();
  let totalElements = body.getNumChildren();
  let listItemIndices = [];
  let knownStatusKeys = new Set();
  let knownInitiative = new Set();
  for (let index = 0; index < totalElements; index++) {
    if (body.getChild(index).getType() === DocumentApp.ElementType.LIST_ITEM && body.getChild(index).findText("%[A-Za-z0-9\.\/\u0020\(\)-]+%")) {
      listItemIndices.push(index);
      let key = body.getChild(index).getText();
      key = key.substring(1, key.length - 1);
      knownStatusKeys.add(key);
      if (key.includes(".")) {
        knownInitiative.add(key.split(".")[0]);
      }
    }
  }

  let otherStatusMap = new Map();
  statusMap.forEach((value, key) => {
    if (!knownStatusKeys.has(key)) {
      //Respondent must have selected other, and intended a different category for this item
      let newKey;
      let keyParts = key.split(".");
      if (keyParts.length === 1) {
        newKey = "Other";
      } else if (!knownInitiative.has(keyParts[0])) {
        newKey = "Other";
      } else if (keyParts.length === 2) {
        newKey = keyParts[0] + ".Other";
      } else {
        Logger.log("Assuming that key [%s] includes a .", key);
        newKey = keyParts[0] + ".Other";
        if (!knownStatusKeys.has(newKey)) {
          Logger.log("Did not find %s as a placeholder so this must be an other category off the top: %s", newKey, keyParts[0]);
          newKey = "Other";
        }
      }
      let mappedCategories = getMapArray(otherStatusMap, newKey);
      mappedCategories.push(key);
    }
  });

  for (let index = listItemIndices.length - 1; index >= 0; index--) {
    //Iterate backwards so inserting status does not make indices invalid
    let listItem = body.getChild(listItemIndices[index]);
    let key = listItem.getText();
    key = key.substring(1, key.length - 1);
    // Logger.log("Actual key is %s", key);
    let statuses = statusMap.get(key);
    if (statuses) {
      Logger.log("Found %s items for %s", statuses.length, key);
      statuses.forEach(value => {
        let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
        if (hasNoStatus(key)) {
          let associateName = getAssociateName(value.kerberos).split(" ")[0];
          inserted.setText(associateName);
        } else {
          inserted.setText("");
          getStatusText(value).forEach(part => {
            inserted.appendText(part.text);
            if (part.isLink) {
              inserted.appendText("\u200B");
              let endOffsetInclusive = inserted.getText().length - 2;
              let startOffset = endOffsetInclusive - part.text.length + 1;
              inserted.setLinkUrl(startOffset, endOffsetInclusive, part.url);
              inserted.setForegroundColor(endOffsetInclusive + 1, endOffsetInclusive + 1, "#000000");
            }
          });
        }
        reportedCount++;
      });
      statusMap.delete(key);
    } else if (key === "Missing Status") {
      let missing = getMissingStatus(readResponseObjects(globalLinks.statusFormId), [statusDocId]);
      missing.statusRequired.forEach(kerberos => {
        let associateName = getAssociateName(kerberos).split(" ")[0];
        let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
        inserted.setText(associateName);
      });
    } else {
      Logger.log("Found no items for %s", key);
    }
    let mappedStatuses = otherStatusMap.get(key);
    if (mappedStatuses) {
      mappedStatuses.forEach(mappedKey => {
        let otherStatuses = statusMap.get(mappedKey);
        otherStatuses.forEach(value => {
          let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
          inserted.setText(">>> " + mappedKey + " - ");
          reportedCount++;
          getStatusText(value).forEach(part => {
            inserted.appendText(part.text);
            if (part.isLink) {
              inserted.appendText("\u200B");
              let endOffsetInclusive = inserted.getText().length - 2;
              let startOffset = endOffsetInclusive - part.text.length + 1;
              inserted.setLinkUrl(startOffset, endOffsetInclusive, part.url);
              inserted.setForegroundColor(endOffsetInclusive + 1, endOffsetInclusive + 1, "#000000");
            }
          });
        });
        statusMap.delete(key);
      });
      otherStatusMap.delete(key);
    }
  }

  if (reportedCount !== responseCount) {
    let message = Utilities.formatString("Warning: there are %d status entries in the form but only %d were reported in this document", responseCount, reportedCount);
    Logger.log(message);
    let paragraph = body.getParagraphs()[0].editAsText();
    paragraph.setText(message);
    paragraph.setFontSize(22);
    paragraph.setBold(true);
  }

  body.getListItems().forEach(listItem => {
    if (listItem.findText("%[A-Za-z0-9\.\/\u0020\(\)-]+%")) {
      body.removeChild(listItem);
    }
  });
  doc.saveAndClose();
}

function getStatusText(responseObject) {
  let name = getAssociateName(responseObject.kerberos);
  if (!name) {
    name = responseObject.kerberos;
  }
  let statusParts = getStatusParts(responseObject.epic + ":\n");
  if (!responseObject.status) {
    Logger.log(responseObject.kerberos + " does not have a status");
  }
  statusParts = statusParts.concat(getStatusParts(responseObject.status));
  statusParts.push({
    isLink: false,
    text: "\n[By " + name + " on " + responseObject.timestamp + "]\n"
  });
  return statusParts;
}

function getKerberosMap(spreadsheetId) {
  let map = new Map();
  let sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Rover");
  let columns = sheet.getLastColumn();
  let header = sheet.getRange(1, 1, 1, columns).getValues()[0];
  // Logger.log(header);
  let values = sheet.getRange(2, 1, sheet.getLastRow(), columns).getValues();
  values.forEach(value => {
    let record = new Map();
    for (let col = 0; col < columns; col++) {
      if (header[col]) {
        record.set(header[col], value[col]);
      }
    }
    if (isNaN(Date.parse(record.get("Termination"))) && value[1]) {
      //Only include associates who don't have a Separation Date
      map.set(value[1], record);
    }
  });
  return map;
}

function getAllUsers() {
  return kerberosMap.keys();
}

function getAssociateName(kerberos) {
  const personMap = kerberosMap.get(kerberos);
  return personMap.get("Name");
}

function getAssociateEmail(kerberos) {
  const personMap = kerberosMap.get(kerberos);
  return personMap.get("Email");
}

function getAssociateManager(kerberos) {
  const personMap = kerberosMap.get(kerberos);
  return personMap.get("Manager");
}

function getAssociateTitle(kerberos) {
  const personMap = kerberosMap.get(kerberos);
  return personMap.get("Job Title");
}

function getStatusParts(string) {
  // Logger.log("Original string is [%s]", string);
  let statusParts = [];
  const regex = new RegExp("\\[([^)]*)\]\\s*\\(([^\)]*)\\)", "g");
  let startIndex = 0;
  let result;
  while ((result = regex.exec(string)) !== null) {
    // Logger.log(`Found ${result[0]} at ${result.index}, with text ${result[1]} and link ${result[2]}. Next starts at ${regex.lastIndex}.`);
    if (result.index > startIndex) {
      //There is regular text before match, that needs to be added
      statusParts.push({
        isLink: false,
        text: string.substring(startIndex, result.index)
      });
    }
    if (result[1] === "") {
      result[1] = "link";
    }
    statusParts.push({
      isLink: true,
      text: result[1],
      url: result[2]
    });
    startIndex = regex.lastIndex;
  }
  if (startIndex < string.length) {
    statusParts.push({
      isLink: false,
      text: string.substring(startIndex)
    });
  }
  return statusParts;
}

function createDraftEmails() {
  if (isPaused('createDraftEmails')) {
    return;
  }
  let sheet = SpreadsheetApp.openById(globalLinks.statusEmailsId).getSheetByName("emails");
  var drafts = [];
  let values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  let signature = getSignature();
  values.forEach(value => {
    drafts.push({
      to: value[0],
      subject: value[3],
      labels: value[4],
      options: {
        bcc: value[2],
        cc: value[1],
        from: 'Babak Mozaffari <Babak@redhat.com>',
        htmlBody: signature
      }
    });
    drafts.forEach(draft => {
      Logger.log(draft);
    });
  });
  drafts.forEach(params => {
    var draft = GmailApp.createDraft(params.to, params.subject, "", params.options);
    params.labels.split(',').forEach(label => {
      Logger.log("Label is " + label);
      var gmailLabel = GmailApp.getUserLabelByName(label);
      draft.getMessage().getThread().addLabel(gmailLabel);
    });
  });
}

function getSignature() {
  return '<div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><br><br>Regards, Babak</div><div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><br></div></div></div></div></div></div></div></div><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-weight: bold; margin: 0px; padding: 0px; font-size: 14px; text-transform: capitalize;">Babak Mozaffari</p><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: 12px; margin: 0px 0px 4px; text-transform: capitalize;">He / Him / His</p><p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; margin: 0px;"><span style="font-size: 12px; text-transform: capitalize;">Director &amp; Distinguished Engineer</span></p><p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; margin: 0px;"><span style="font-size: 12px; text-transform: capitalize;">Workloads, AppDev, OpenShift AI</span></p><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: 12px; margin: 0px; text-transform: capitalize;">Ecosystem Engineering</p><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: 12px; margin: 0px; text-transform: capitalize;"><a href="https://www.redhat.com/" target="_blank" style="color: rgb(0, 136, 206); margin: 0px;">Red Hat</a></p><div style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: medium; margin-bottom: 4px;"></div><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; margin: 0px; font-size: 12px;"><span style="margin: 0px; padding: 0px;"><a href="mailto:Babak@redhat.com" target="_blank" style="color: rgb(0, 0, 0); margin: 0px;">Babak@redhat.com</a>&nbsp; &nbsp;</span><br>M: <a href="tel:+1-310-857-8604" target="_blank" style="color: rgb(0, 0, 0); margin: 0px;">+1-310-857-8604</a>&nbsp; &nbsp;&nbsp;</p><div style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: medium; margin-top: 12px;"><table border="0"><tbody><tr><td width="100px" style="margin: 0px;"><a href="https://red.ht/sig" target="_blank" style="color: rgb(17, 85, 204);"><img src="https://static.redhat.com/libs/redhat/brand-assets/latest/corp/logo.png" width="90" height="auto"></a></td></tr></tbody></table></div><div style="text-align: start;color: rgb(34, 34, 34);font-size: small;"><div><div><div><div><div><div><div><div style="color: rgb(0, 0, 0);font-size: medium;"><table border="0"><tbody><tr><td width="100px"><br></td></tr></tbody></table></div></div></div></div></div></div></div></div></div>';
}

function sendDraftEmails() {
  let sheet = SpreadsheetApp.openById(globalLinks.statusEmailsId).getSheetByName("emails");
  let emailSetup = sheet.getRange(2, 4, sheet.getLastRow() - 1, 6).getValues();
  const fwdMap = new Map(emailSetup.map((row) => [row[0], row[2]]));
  let count = 0
  let total = GmailApp.getDrafts().length;
  while (count < total) {
    try {
      let drafts = GmailApp.getDrafts();
      if (drafts.length === 0) {
        Logger.log("Unexpected error, draft message not found and must have been processed by another thread");
        return count;
      } else {
        var message = drafts[0].getMessage();
        let labels = message.getThread().getLabels();
        Logger.log("Sending draft with subject [%s] addressed to %s", message.getSubject(), message.getTo());
        var sentMessage = drafts[0].send();
        var forwardEmail = fwdMap.get(sentMessage.getSubject());
        if (forwardEmail) {
          Logger.log("Forwarding to %s", forwardEmail);
          sentMessage.forward(forwardEmail);
        }
        let sentThread = sentMessage.getThread();
        for (label of labels) {
          sentThread.addLabel(label);
        }
        count++;
      }
    } catch (err) {
      Logger.log("Failed to read drafts and send email, will try again", err);
    }
  }
  return count;
}

function compareAssignments() {
  let userAssignmentMap = getUserAssignmentMap();
  nextUser: for (let kerberos of userAssignmentMap.keys()) {
    //Check if user's manager is using the weekly status framework
    let managerUID = getAssociateManager(kerberos);
    if (documentLinks.get("Managers").has(managerUID)) {
      continue;
    }

    //Check if user's assignments leverage the weekly status framework
    for (let assignment of userAssignmentMap.get(kerberos)) {
      let initiative = assignment.split('\.')[0]
      if (documentLinks.get("Initiatives").has(initiative)) {
        continue nextUser;
      }
    }
    //If there was no match based on manager or initiative, the user status is not required for weekly status compilation
    userAssignmentMap.delete(kerberos);
  }

  let statusMap = new Map();
  let responseObjects = readResponseObjects(globalLinks.statusFormId);
  for (let responseObject of responseObjects) {
    let statusArray = getMapArray(statusMap, responseObject.kerberos);
    let assignment;
    if ( hasNoStatus(responseObject.initiative) ) {
      userAssignmentMap.delete(responseObject.kerberos);
      continue;
    } else {
      assignment = responseObject.initiative;
    }
    statusArray.push(assignment);
  };

  //Remove from the list those who are on PTO as per their personal calendars
  for (let kerberos of userAssignmentMap.keys()) {
    if (isOnPTO(kerberos.concat("@redhat.com"))) {
      Logger.log("The personal calendar of %s shows as OOO, so won't mark them as missing status", getAssociateName(kerberos));
      userAssignmentMap.delete(kerberos);
    }
  }

  let noActivity = new Map();
  let noAssignment = new Map();
  let benched = new Map();
  for (let [kerberos, assignments] of userAssignmentMap) {
    let statuses = statusMap.get(kerberos);
    if (assignments.length === 0 && statuses !== undefined) {
      noAssignment.set(kerberos, statuses);
    } else if (assignments.length === 0 && statuses === undefined) {
      benched.set(kerberos, "");
    } else if (assignments.length > 0 && statuses === undefined) {
      noActivity.set(kerberos, assignments);
    } else if (assignments.length > 0 && statuses !== undefined) {
      assignments.forEach(assignment => {
        if (!statuses.includes(assignment)) {
          getMapArray(noActivity, kerberos).push(assignment);
        }
      });
      statuses.forEach(status => {
        if (!assignments.includes(status)) {
          getMapArray(noAssignment, kerberos).push(status);
        }
      });
    }
  }
  return [noAssignment, benched, noActivity];
}

function getUserAssignmentMap() {
  let assignmentMap = new Map();
  let assignmentSheet = SpreadsheetApp.openById(globalLinks.rosterSheetId).getSheetByName("Roster by Person");
  for (let row = 3; row < assignmentSheet.getLastRow(); row++) {
    let values = assignmentSheet.getRange(row, 1, 1, 5).getValues();
    let kerberos = values[0][4];
    if (kerberos.length === 0) {
      break;
    }
    let assignmentArray = getMapArray(assignmentMap, kerberos);
    assignmentArray.push(values[0][0]);
  }
  let availabilitySheet = SpreadsheetApp.openById(globalLinks.rosterSheetId).getSheetByName("Availability");
  for (let row = 2; row < availabilitySheet.getLastRow(); row++) {
    let values = availabilitySheet.getRange(row, 1, 1, 4).getValues();
    let utilization = values[0][3];
    if (utilization === 0) {
      let kerberos = values[0][1];
      Logger.log("%s is not assigned any tasks", values[0][0]);
      getMapArray(assignmentMap, kerberos); // empty array pushed for associate
    }
  }
  return assignmentMap;
}

function printMismatch() {
  const response = getMismatchResponse("swkale", textFormatting);
  if (response.length === 0) {
    response = "No missing status entries this week";
  }
  Logger.log(response);
}

function isOnPTO(email) {
  //Look for events on user calendar right now to increase performance
  let nextMinute = new Date();
  nextMinute.setMinutes(nextMinute.getMinutes() + 2);
  let params = {
    timeMin: Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
    timeMax: Utilities.formatDate(nextMinute, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ'),
    showDeleted: false,
  };
  let response = Calendar.Events.list(email, params);
  let ooo = false;
  response.items.forEach(entry => {
    if (entry.eventType === "outOfOffice" && isFullDay(entry) ) {
      ooo = true;
    }
  });
  return ooo;
}

function isFullDay(event) {
  if (event.start != null && event.start.date != null ) {
    //If a date and not a dateTime, this is a full day event as entered
    return true;
  } else {
    let start = Utilities.parseDate(event.start.dateTime, "UTC", 'yyyy-MM-dd\'T\'HH:mm:ssX');
    let end = Utilities.parseDate(event.end.dateTime, "UTC", 'yyyy-MM-dd\'T\'HH:mm:ssX');
    let durationHours = (end - start) / 3600000;
    return durationHours > 4;
  }
}

function validateRecipients() {
  if (isPaused('validateRecipient')) {
    return;
  }
  let sheet = SpreadsheetApp.openById(globalLinks.statusEmailsId).getSheetByName("emails");
  let values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const recipients = new Set();
  values.forEach(value => {
    for (const index of [0, 1, 2]) {
      for (let email of value[index].split(',')) {
        if (email.includes('<')) {
          email = email.match(/<([^>]*)>/)[1];
        }
        email = email.trim();
        if (!email.includes('-') && email.includes('@')) {
          //Skip email addresses with dashes in them, as they are very likely groups
          recipients.add(email);
        }
      }
    }
  });
  for (const email of recipients) {
    let person = People.People.searchDirectoryPeople({
      readMask: 'names',
      query: email,
      sources: [
        'DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'
      ]
    });

    if (person.totalSize === 1) {
      //Proper and valid person email
      continue;
    }

    try {
      GroupsApp.getGroupByEmail(email);
      //No exception implies this is a group, so no need to flag it
      continue;
    } catch (e) {
    }

    const subject = "Could not find " + email + " in the corporate directory!";
    const body = "Validate whether the user still exists! Start by checking https://rover.redhat.com/people/profile/" + email.split("@")[0];
    Logger.log(subject + "\n" + body);
    GmailApp.sendEmail("babak@redhat.com", subject, body);
  }
}

function hasNoStatus(category) {
  if (category === "PTO / No Status" || category === "Learning" ) {
    return true;
  } else {
    return false;
  }
}

function isPaused(trigger) {
  let sheet = SpreadsheetApp.openById('1vrAk1qITGj72Vm4yVcsYpu3rHWZhALgA0uW7tCiX1sc').getSheetByName('main');
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    let name = sheet.getRange(row, 1, 1, 1).getValue();
    if (name === trigger) {
      let now = new Date();
      let start = sheet.getRange(row, 2, 1, 1).getValue();
      if (start && start > now) {
        return false;
      }
      let end = sheet.getRange(row, 3, 1, 1).getValue();
      if (end && end < now) {
        return false;
      }
      //Script is paused
      Logger.log("The trigger called %s is paused with a start date of %s and end date of %s", trigger, start, end);
      return true;
    } else if (!name) {
      return false;
    }
  }
}
