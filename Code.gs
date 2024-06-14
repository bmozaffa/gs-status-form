let documentLinks = getDocumentLinks();
let globalLinks = getGlobalLinks();
let kerberosMap = getKerberosMap(globalLinks.roverSheetId);

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

function doGet(e) {
  let command = e.parameter.command;
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
    response = compileStatus();
  } else if (command === "archive") {
    let archived = archiveReports(e.parameter.days);
    response = "Successfully archived " + archived + " form submissions";
  } else if (command === "send-drafts") {
    let sentMessages = sendDraftEmails();
    response = "Successfully sent " + sentMessages + " emails from saved drafts";
  } else if (command === "mismatch") {
    response = getMismatchResponse("<p>", "<br/>");
  } else {
    response = "Provide the command (missing, generate, etc) as a request parameter: https://script.google.com/..../exec?command=generate";
  }
  return HtmlService.createHtmlOutput(response);
}

function getMismatchResponse(paragraphBreak, lineBreak) {
  let mismatch = compareAssignments();
  let response = "";
  if (mismatch[0].size > 0) {
    response += paragraphBreak + "Activities outside of Roster assignments";
  }
  for (let kerberos of mismatch[0].keys()) {
    response += lineBreak + kerberos + ": ";
    response += mismatch[0].get(kerberos).join(", ");
  }
  if (mismatch[1].size > 0) {
    response += paragraphBreak + "Associates with neither activity nor assignment";
  }
  for (let kerberos of mismatch[1]) {
    response += lineBreak + kerberos;
  }
  if (mismatch[2].size > 0) {
    response += paragraphBreak + "Assignments without any activity this week";
  }
  for (let kerberos of mismatch[2].keys()) {
    response += lineBreak + kerberos + ": ";
    response += mismatch[2].get(kerberos).join(", ");
  }
  return response;
}

function getMissingStatusReport() {
  let allStatusDocumentIds = Array.from(documentLinks.get("Templates").keys())
  let missing = getMissingStatus(readResponseObjects(globalLinks.statusFormId), allStatusDocumentIds);
  let missingHierarchy = new Map();
  missing.statusRequired.forEach(kerberos => {
    let associateInfo = missing.kerberosMap.get(kerberos);
    let associateName = associateInfo.get("Name").split(" ")[0];
    let managerUID = associateInfo.get("Manager UID");
    let managerName = missing.kerberosMap.get(managerUID).get("Name").split(" ")[0];
    let associates = getMapArray(missingHierarchy, managerName);
    associates.push(associateName);
    console.log(">> %s", associateName);
  });
  return missingHierarchy;
}

function notifyMissingStatus() {
  if (utils.isPaused('notifyMissingStatus')) {
    return;
  }
  let allStatusDocumentIds = Array.from(documentLinks.get("Templates").keys())
  let missing = getMissingStatus(readResponseObjects(globalLinks.statusFormId), allStatusDocumentIds);

  let subjectBase = "Missing weekly status report for ";
  let body = "This is an automated message to notify you that a weekly status report was not detected for the week.\n\nPlease submit your status report promptly. Kindly ignore this message if you are off work and a status entry is not expected.";
  missing.statusRequired.forEach(kerberos => {
    let associateInfo = missing.kerberosMap.get(kerberos);
    let managerInfo = missing.kerberosMap.get(associateInfo.get("Manager UID"));
    let to = associateInfo.get("Email");
    let cc = managerInfo.get("Email");
    let subject = subjectBase + associateInfo.get("Name").split(" ")[0];
    GmailApp.sendEmail(to, subject, body, {cc: cc});
    console.log("%s who reports to %s is missing status, sent them an email!", associateInfo.get("Name"), managerInfo.get("Name"));
    console.log("Sending an email to %s with subject line [%s] and body [%s] and copying %s", to, subject, body, cc);
  });
}

function getMissingStatus(responseObjects, statusDocIds) {
  //Figure out who need to submit status report
  let statusRequired = new Set();
  const excludedStatus = ["Director, Software Engineering_Global", "Senior Manager, Software Engineering_Global", "Manager, Software Engineering", "Associate Manager, Software Engineering"];
  kerberosMap.forEach((value, key) => {
    if (excludedStatus.includes(value.get("Job Title"))) {
      //Exclude managers from status entry
    } else if (value.get("Job Title").includes("Director")) {
      //Exclude directors from missing entry
    } else if (value.get("Job Title").includes("Intern")) {
      //Exclude interns from missing entry
    } else if (key) {
      //Add entry only if non-empty
      statusRequired.add(key);
    }
  });

  //Remove from the list those who have reported status
  for (let responseObj of responseObjects) {
    statusRequired.delete(responseObj.kerberos);
  }

  //Only report missing status for associates relevant to provided status documents
  let userAssignmentMap = getUserAssignmentMap();
  nextUser: for (let kerberos of statusRequired) {
    //Check if user's manager owns this status
    let associateInfo = kerberosMap.get(kerberos);
    let managerUID = associateInfo.get("Manager UID");
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
      console.log("The personal calendar of %s shows as OOO, so won't mark them as missing status", kerberosMap.get(kerberos).get("Name"));
      statusRequired.delete(kerberos);
    }
  }

  console.log("Missing status for %s people", statusRequired.size);
  let missingStatus = {};
  missingStatus.statusRequired = statusRequired;
  missingStatus.kerberosMap = kerberosMap;
  return missingStatus;
}

function archiveReports(days) {
  let form = FormApp.openById(globalLinks.statusFormId);
  let cutoff = new Date();
  if (!days) {
    days = 7;
  }
  cutoff.setDate(cutoff.getDate() - days);
  console.log("Cutoff is " + cutoff);
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
  console.log("Archived " + count + " responses out of " + total);
  return count;
}

function logStatus() {
  let responseObjects = readResponseObjects(globalLinks.statusFormId);
  let map = getStatusMap(responseObjects);
  map.forEach((statusList, category) => {
    if (category === "PTO / Learning / No Status") {
      statusList.forEach(responseObject => {
        console.log(responseObject.kerberos + " has PTO / Learning / No Status");
      });
    } else {
      statusList.forEach(responseObject => {
        let status = getStatusText(responseObject, kerberosMap).map(status => {
          return status.text;
        }).join("");
        console.log(status);
      });
    }
  });
}

function compileStatus() {
  if (utils.isPaused('compileStatus')) {
    return;
  }
  let statusDocs = new Set();
  for (let statusDocId of documentLinks.get("Templates").keys()) {
    if (needsUpdate(globalLinks.statusFormId, statusDocId)) {
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
    let associateInfo = kerberosMap.get(responseObject.kerberos);
    let managerUID = associateInfo.get("Manager UID");
    return documentLinks.get("Managers").get(managerUID) === statusDocId;
  }
}

function needsUpdate(formId, statusDocId) {
  let form = FormApp.openById(formId);
  let lastStatus;
  form.getResponses().forEach(response => {
    if (lastStatus) {
      if (response.getTimestamp() > lastStatus) {
        lastStatus = response.getTimestamp();
      }
    } else {
      lastStatus = response.getTimestamp();
    }
  });
  console.log("Got latest status entry dated " + lastStatus);
  if (!lastStatus) {
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
    console.log("The doc was last updated on " + lastUpdate);
    if (lastStatus > lastUpdate) {
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
  link.setLinkUrl("https://script.google.com/a/macros/redhat.com/s/AKfycbwGck8vr-ZNHVCKRg8-y1p5hfVt4Iognzm4zZoOd0WrsORfzLOKNVHk4te0-qOnHCWU/exec?command=generate");

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
    console.log("Status reported for %s", mapKey);
    let statusArray = getMapArray(statusMap, mapKey);
    statusArray.push(responseObject);
  });
  return statusMap;
}

function readResponseObjects(formId) {
  let form = FormApp.openById(formId);
  let responses = form.getResponses();
  // form.getItems().forEach(value => {
  //   console.log("item is %s", value.getTitle());
  // })
  let responseObjects = [];
  // console.log("Found %s status entries", responses.length);
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
    // console.log("There was no array for %s, but one has been created now", key);
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
        console.log("Assuming that key [%s] includes a .", key);
        newKey = keyParts[0] + ".Other";
        if (!knownStatusKeys.has(newKey)) {
          console.log("Did not find %s as a placeholder so this must be an other category off the top: %s", newKey, keyParts[0]);
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
    // console.log("Actual key is %s", key);
    let statuses = statusMap.get(key);
    if (statuses) {
      console.log("Found %s items for %s", statuses.length, key);
      statuses.forEach(value => {
        let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
        if (key === "PTO / Learning / No Status") {
          let associateInfo = kerberosMap.get(value.kerberos);
          let associateName = associateInfo.get("Name").split(" ")[0];
          inserted.setText(associateName);
        } else {
          inserted.setText("");
          getStatusText(value, kerberosMap).forEach(part => {
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
        let associateInfo = missing.kerberosMap.get(kerberos);
        let associateName = associateInfo.get("Name").split(" ")[0];
        let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
        inserted.setText(associateName);
      });
    } else {
      console.log("Found no items for %s", key);
    }
    let mappedStatuses = otherStatusMap.get(key);
    if (mappedStatuses) {
      mappedStatuses.forEach(mappedKey => {
        let otherStatuses = statusMap.get(mappedKey);
        otherStatuses.forEach(value => {
          let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
          inserted.setText(">>> " + mappedKey + " - ");
          reportedCount++;
          getStatusText(value, kerberosMap).forEach(part => {
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
    console.log(message);
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

function getStatusText(responseObject, kerberosMap) {
  let name = responseObject.kerberos;
  if (kerberosMap.get(responseObject.kerberos)) {
    name = kerberosMap.get(responseObject.kerberos).get("Name");
  }
  let statusParts = getStatusParts(responseObject.epic + ":\n");
  if (!responseObject.status) {
    console.log(responseObject.kerberos + " does not have a status");
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
  // console.log(header);
  let values = sheet.getRange(2, 1, sheet.getLastRow(), columns).getValues();
  values.forEach(value => {
    let record = new Map();
    for (let col = 0; col < columns; col++) {
      if (header[col]) {
        record.set(header[col], value[col]);
      }
    }
    if (isNaN(Date.parse(record.get("Separation Date")))) {
      //Only include associates who don't have a Separation Date
      map.set(value[1], record);
    }
  });
  return map;
}

function getStatusParts(string) {
  // console.log("Original string is [%s]", string);
  let statusParts = [];
  const regex = new RegExp("\\[([^)]*)\]\\s*\\(([^\)]*)\\)", "g");
  let startIndex = 0;
  let result;
  while ((result = regex.exec(string)) !== null) {
    // console.log(`Found ${result[0]} at ${result.index}, with text ${result[1]} and link ${result[2]}. Next starts at ${regex.lastIndex}.`);
    if (result.index > startIndex) {
      //There is regular text before match, that needs to be added
      statusParts.push({
        isLink: false,
        text: string.substring(startIndex, result.index)
      });
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
  if (utils.isPaused('createDraftEmails')) {
    return;
  }
  let sheet = SpreadsheetApp.openById(globalLinks.statusEmailsId).getSheetByName("emails");
  var drafts = [];
  let values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  let signature = '<div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><br><br>Regards, Babak</div><div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><div dir="ltr"><br></div></div></div></div></div></div></div></div><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-weight: bold; margin: 0px; padding: 0px; font-size: 14px; text-transform: capitalize;">Babak Mozaffari</p><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: 12px; margin: 0px 0px 4px; text-transform: capitalize;">He / Him / His</p><p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; margin: 0px;"><span style="font-size: 12px; text-transform: capitalize;">Director &amp; Distinguished Engineer</span></p><p style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; margin: 0px;"><span style="font-size: 12px; text-transform: capitalize;">Workloads, AppDev, OpenShift AI</span></p><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: 12px; margin: 0px; text-transform: capitalize;">Ecosystem Engineering</p><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: 12px; margin: 0px; text-transform: capitalize;"><a href="https://www.redhat.com/" target="_blank" style="color: rgb(0, 136, 206); margin: 0px;">Red Hat</a></p><div style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: medium; margin-bottom: 4px;"></div><p style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; margin: 0px; font-size: 12px;"><span style="margin: 0px; padding: 0px;"><a href="mailto:Babak@redhat.com" target="_blank" style="color: rgb(0, 0, 0); margin: 0px;">Babak@redhat.com</a>&nbsp; &nbsp;</span><br>M: <a href="tel:+1-310-857-8604" target="_blank" style="color: rgb(0, 0, 0); margin: 0px;">+1-310-857-8604</a>&nbsp; &nbsp;&nbsp;</p><div style="font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; color: rgb(0, 0, 0); font-family: RedHatText, sans-serif; font-size: medium; margin-top: 12px;"><table border="0"><tbody><tr><td width="100px" style="margin: 0px;"><a href="https://red.ht/sig" target="_blank" style="color: rgb(17, 85, 204);"><img src="https://static.redhat.com/libs/redhat/brand-assets/latest/corp/logo.png" width="90" height="auto"></a></td></tr></tbody></table></div><div style="text-align: start;color: rgb(34, 34, 34);font-size: small;"><div><div><div><div><div><div><div><div style="color: rgb(0, 0, 0);font-size: medium;"><table border="0"><tbody><tr><td width="100px"><br></td></tr></tbody></table></div></div></div></div></div></div></div></div></div>';

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
      console.warn("Failed to read drafts and send email, will try again", err);
    }
  }
  return count;
}

function compareAssignments() {
  let userAssignmentMap = getUserAssignmentMap();
  nextUser: for (let kerberos of userAssignmentMap.keys()) {
    //Check if user's manager is using the weekly status framework
    let associateInfo = kerberosMap.get(kerberos);
    let managerUID = associateInfo.get("Manager UID");
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
    if (responseObject.initiative === 'PTO / Learning / No Status') {
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
      console.log("The personal calendar of %s shows as OOO, so won't mark them as missing status", kerberosMap.get(kerberos).get("Name"));
      userAssignmentMap.delete(kerberos);
    }
  }

  let noActivity = new Map();
  let noAssignment = new Map();
  let benched = new Set();
  for (let [kerberos, assignments] of userAssignmentMap) {
    let statuses = statusMap.get(kerberos);
    if (assignments === undefined && statuses !== undefined) {
      noAssignment.set(kerberos, statuses);
    } else if (assignments === undefined && statuses === undefined) {
      benched.add(kerberos);
    } else if (assignments !== undefined && statuses === undefined) {
      noActivity.set(kerberos, assignments);
    } else if (assignments !== undefined && statuses !== undefined) {
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
      console.log("%s is not assigned any tasks", values[0][0]);
      getMapArray(assignmentMap, kerberos); // empty array pushed for associate
    }
  }
  return assignmentMap;
}

function printMismatch() {
  let response = getMismatchResponse("\n\n\n", "\n");
  if (response.length === 0) {
    response = "No missing status entries this week";
  }
  console.log(response);
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
