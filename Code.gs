function getLinks() {
  let links = {};
  links.statusFormId = "14g3I22FRYFV9kbE9csTdlbrKK3XbnHXo8rsMDg-Z_HI";
  links.templateDocId = "1BzP4CzYdhu_VQ2-0TrCsa9drSFiKxFFzLYy4m9jpaKU";
  links.statusDocId = "14mWOqhXRyxFST6EZz6laB7ZCFhU7xDRQxKGzam4NDCQ";
  links.roverSheetId = "1i7y_tFpeO68SetmsU2t-C6LsFETuZtkJGY5AVZ2PHW8";
  links.statusEmailsId = "1pke_nZSAwVFL9iIx-HKgaaZU4lMmr5aParHgN9wGdXE";
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
  } else {
    response = "Provide the command (missing, generate, etc) as a request parameter: https://script.google.com/..../exec?command=generate";
  }
  return HtmlService.createHtmlOutput(response);
}

function getMissingStatusReport() {
  let links = getLinks();
  let formMap = readForm(links.statusFormId);
  let missing = getMissingStatus(formMap, links.roverSheetId);
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
  let links = getLinks();
  let formMap = readForm(links.statusFormId);
  let missing = getMissingStatus(formMap, links.roverSheetId);

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

function getMissingStatus(formMap, roverSheetId) {
  let kerberosMap = getKerberosMap(roverSheetId);

  //Figure out who need to submit status report
  let statusRequired = new Set();
  const excludedStatus = ["Director, Software Engineering_Global", "Manager, Software Engineering", "Associate Manager, Software Engineering"];
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
  formMap.forEach(statusList => {
    statusList.forEach(responseObject => {
      statusRequired.delete(responseObject.kerberos);
    });
  });
  console.log("Missing status for %s people", statusRequired.size);
  let missingStatus = {};
  missingStatus.statusRequired = statusRequired;
  missingStatus.kerberosMap = kerberosMap;
  return missingStatus;
}

function archiveReports(days) {
  let links = getLinks();
  let form = FormApp.openById(links.statusFormId);
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
  let links = getLinks();
  let map = readForm(links.statusFormId);
  let kerberosMap = getKerberosMap(links.roverSheetId);
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
  let links = getLinks();
  if (needsUpdate(links.statusFormId, links.statusDocId)) {
    copyTemplate(links.templateDocId, links.statusDocId);
    let responseObjects = readResponseObjects(links.statusFormId);
    let statusMap = getStatusMap(responseObjects);
    insertStatus(links.statusDocId, links.roverSheetId, statusMap, responseObjects.length);
    let form = FormApp.openById(links.statusFormId);
    return "Successfully generated status report based on " + form.getResponses().length + " form submissions";
  } else {
    return "There is no status entry since document was generated";
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

  let doc = DocumentApp.openById(statusDocId);
  let lastUpdateMessage = "This document is no longer auto-generated. It was last generated at ";
  let paragraph = doc.getBody().getParagraphs()[0].getText();
  if (isNaN(Date.parse(paragraph.substring(lastUpdateMessage.length)))) {
    console.error("Failed to parse last update time");
    return true;
  } else {
    let lastUpdate = new Date(paragraph.substring(lastUpdateMessage.length));
    console.log("The doc was last updated on " + lastUpdate);
    return lastStatus > lastUpdate;
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
  link.setLinkUrl("https://script.google.com/a/macros/redhat.com/s/AKfycbwAxjr-cpOLXs_Y9KERBs7SktzkfSKO9om4RHX5nFzIHKV1D-z8AS-9tajF1fQUt90q/exec?command=generate");

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
    if (li.findText("%[A-Za-z\.\/\u0020]+%")) {
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

function readForm(formId) {
  let responseObjects = readResponseObjects(formId);
  return getStatusMap(responseObjects);
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

function insertStatus(statusDocId, roverSheetId, statusMap, responseCount) {
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
    if (body.getChild(index).getType() === DocumentApp.ElementType.LIST_ITEM && body.getChild(index).findText("%[A-Za-z\.\/\u0020]+%")) {
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

  let kerberosMap = getKerberosMap(roverSheetId);
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
            }
          });
        }
        reportedCount++;
      });
      statusMap.delete(key);
    } else if (key === "Missing Status") {
      let missing = getMissingStatus(statusMap, roverSheetId);
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
    if (listItem.findText("%[A-Za-z\.\/\u0020]+%")) {

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
  const spreadsheetId = getLinks().statusEmailsId;
  let sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("emails");
  var drafts = [];
  let values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  values.forEach(value => {
    drafts.push({
      to: value[0],
      subject: value[3],
      labels: value[4],
      options: {
        bcc: value[2],
        cc: value[1],
        from: 'Babak Mozaffari <Babak@redhat.com>'
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
