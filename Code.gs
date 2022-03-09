function getLinks() {
  let links = {};
  links.statusFormId = "14g3I22FRYFV9kbE9csTdlbrKK3XbnHXo8rsMDg-Z_HI";
  links.templateDocId = "1BzP4CzYdhu_VQ2-0TrCsa9drSFiKxFFzLYy4m9jpaKU";
  links.statusDocId = "11UHpXCgpma_wDrXYfsD6VxVUAhHvrugSwMkPyKx2rlk";
  links.roverSheetId = "1i7y_tFpeO68SetmsU2t-C6LsFETuZtkJGY5AVZ2PHW8";
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
    let links = getLinks();
    let form = FormApp.openById(links.statusFormId);
    let statusEntries = form.getResponses().length;
    compileStatus();
    response = "Successfully generated status report based on " + statusEntries + " form submissions";
  } else if (command === "archive") {
    let form = FormApp.openById(getLinks().statusFormId);
    let statusEntries = form.getResponses().length;
    archiveReports();
    response = "Successfully archived " + statusEntries + " form submissions";
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
  let formMap = readForm(statusFormId);
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
  kerberosMap.forEach((value, key) => {
    if (value.get("Job Title").includes("Manager")) {
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
      statusRequired.delete(responseObject.kerberbos);
    });
  });
  console.log("Missing status for %s people", statusRequired.size);
  let missingStatus = {};
  missingStatus.statusRequired = statusRequired;
  missingStatus.kerberosMap = kerberosMap;
  return missingStatus;
}

function archiveReports() {
  let links = getLinks();
  // let archiveSheet = SpreadsheetApp.openById(links.?????).getSheets()[0];
  // let responseObjects = readResponseObjects(links.statusFormId);
  // responseObjects.forEach(obj => {
  //   console.log(obj);
  //   archiveSheet.appendRow([obj.kerberbos, obj.timestamp, obj.initiative, obj.effort, obj.epic, obj.status]);
  // });
  let form = FormApp.openById(links.statusFormId);
  form.deleteAllResponses();
}

function logStatus() {
  let links = getLinks();
  let map = readForm(links.statusFormId);
  let kerberosMap = getKerberosMap(links.roverSheetId);
  map.forEach(statusList => {
    statusList.forEach(responseObject => {
      let status = getStatusText(responseObject, kerberosMap).map(status => {
        return status.text;
      }).join("");
      console.log(status);
    });
  });
}

function compileStatus() {
  let links = getLinks();
  copyTemplate(links.templateDocId, links.statusDocId);
  let statusMap = readForm(links.statusFormId);
  insertStatus(links.statusDocId, links.roverSheetId, statusMap);
}

function copyTemplate(templateDocId, statusDocId) {
  let template = DocumentApp.openById(templateDocId).getBody().copy();
  let doc = DocumentApp.openById(statusDocId);
  let body = doc.getBody();
  body.appendPageBreak();
  body.clear();
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
    if (li.findText("%.*%")) {
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
  let statusMap = new Map();
  let responseObjects = readResponseObjects(formId);
  // console.log("Found %s status entries", responseObjects.length);
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
    responseObj.kerberbos = response.getRespondentEmail().split('@')[0];
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

function insertStatus(statusDocId, roverSheetId, statusMap) {
  let doc = DocumentApp.openById(statusDocId);
  let body = doc.getBody();
  let totalElements = body.getNumChildren();
  let listItemIndices = [];
  let knownStatusKeys = new Set();
  for (let index = 0; index < totalElements; index++) {
    if (body.getChild(index).getType() === DocumentApp.ElementType.LIST_ITEM && body.getChild(index).findText("%.*%")) {
      listItemIndices.push(index);
      let key = body.getChild(index).getText();
      key = key.substring(1, key.length - 1);
      knownStatusKeys.add(key);
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
      } else if (keyParts.length === 2) {
        newKey = keyParts[0] + ".Other";
      } else {
        throw new Error("Unexpected use of dot notation");
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
    let kerberosMap = getKerberosMap(roverSheetId);
    if (statuses) {
      statuses.forEach(value => {
        let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
        if (key === "PTO / Learning / No Status") {
          let associateInfo = kerberosMap.get(value.kerberbos);
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
    }
    let mappedStatuses = otherStatusMap.get(key);
    if (mappedStatuses) {
      mappedStatuses.forEach(mappedKey => {
        let otherStatuses = statusMap.get(mappedKey);
        otherStatuses.forEach(value => {
          let inserted = body.insertListItem(listItemIndices[index] + 1, listItem.copy()).editAsText();
          inserted.setText(">>> " + mappedKey + " - ");
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

  statusMap.forEach((value, key) => {
    console.log("Did not report status for %s", key);
  });

  body.getListItems().forEach(listItem => {
    if (listItem.findText("%.*%")) {
      body.removeChild(listItem);
    }
  });
  doc.saveAndClose();
}

function getStatusText(responseObject, kerberosMap) {
  let statusParts = getStatusParts(responseObject.epic + ":\n");
  statusParts = statusParts.concat(getStatusParts(responseObject.status));
  statusParts.push({
    isLink: false,
    text: "\n[By " + kerberosMap.get(responseObject.kerberbos).get("Name") + " on " + responseObject.timestamp + "]\n"
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
    map.set(value[1], record);
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
