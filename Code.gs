function getLinks() {
  var links = new Object();
  links.statusFormId = "14g3I22FRYFV9kbE9csTdlbrKK3XbnHXo8rsMDg-Z_HI";
  links.templateDocId = "1BzP4CzYdhu_VQ2-0TrCsa9drSFiKxFFzLYy4m9jpaKU";
  links.statusDocId = "11UHpXCgpma_wDrXYfsD6VxVUAhHvrugSwMkPyKx2rlk";
  links.roverSheetId = "1i7y_tFpeO68SetmsU2t-C6LsFETuZtkJGY5AVZ2PHW8";
  return links;
}

function doGet(e) {
  var command = e.parameter.command;
  var response;
  if (command === "missing") {
    var response = "";
    var missingHierarchy = getMissingStatusReport();
    missingHierarchy.forEach((associates, manager) => {
      response += "<p>" + manager + ": ";
      response += associates.join(", ");
      response += " has/have not entered status";
    });
    if (response.length === 0) {
      response = "No missing status entries this week";
    }
  } else if (command === "generate") {
    var links = getLinks();
    var form = FormApp.openById(links.statusFormId);
    var statusEntries = form.getResponses().length;
    compileStatus();
    response = "Successfully generated status report based on " + statusEntries + " form submissions";
  } else {
    response = "Provide the command (missing, generate, etc) as a request parameter: https://script.google.com/..../exec?command=generate";
  }
  return HtmlService.createHtmlOutput(response);
}

function logMissingStatus() {
  var missingHierarchy = getMissingStatusReport();
  missingHierarchy.forEach((associates, manager) => {
    console.log("For %s:", manager);
    associates.forEach(associate => {
      console.log(">>> %s", associate);
    });
  });
}

function getMissingStatusReport() {
  var links = getLinks();
  var missing = getMissingStatus(links.statusFormId, links.roverSheetId);
  var missingHierarchy = new Map();
  missing.statusRequired.forEach(kerberos => {
    var associateInfo = missing.kerberosMap.get(kerberos);
    var associateName = associateInfo.get("Name").split(" ")[0];
    var managerUID = associateInfo.get("Manager UID");
    var managerName = missing.kerberosMap.get(managerUID).get("Name").split(" ")[0];
    var associates = getMapArray(missingHierarchy, managerName);
    associates.push(associateName);
    console.log(">> %s", associateName);
  });
  return missingHierarchy;
}

function notifyMissingStatus() {
  var links = getLinks();
  var missing = getMissingStatus(links.statusFormId, links.roverSheetId);

  var subjectBase = "Missing weekly status report for ";
  var body = "This is an automated message to notify you that a weekly status report was not detected for the week.\n\nPlease submit your status report promptly. Kindly ignore this message if you are off work and a status entry is not expected.";
  missing.statusRequired.forEach(kerberos => {
    var associateInfo = missing.kerberosMap.get(kerberos);
    var managerInfo = missing.kerberosMap.get(associateInfo.get("Manager UID"));
    var to = associateInfo.get("Email");
    var cc = managerInfo.get("Email");
    var subject = subjectBase + associateInfo.get("Name").split(" ")[0];
    GmailApp.sendEmail(to, subject, body, { cc: cc });
    console.log("%s who reports to %s is missing status, sent them an email!", associateInfo.get("Name"), managerInfo.get("Name"));
    console.log("Sending an email to %s with subject line [%s] and body [%s] and copying %s", to, subject, body, cc);
  });
}

function getMissingStatus(statusFormId, roverSheetId) {
  var map = readForm(statusFormId);
  var kerberosMap = getKereberosMap(roverSheetId);

  //Figure out who need to submit status report
  var statusRequired = new Set();
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
  map.forEach(statusList => {
    statusList.forEach(responseObject => {
      statusRequired.delete(responseObject.kerberbos);
    });
  });
  console.log("Missing status for %s people", statusRequired.size);
  var missingStatus = new Object();
  missingStatus.statusRequired = statusRequired;
  missingStatus.kerberosMap = kerberosMap;
  return missingStatus;
}

function archiveReports() {
  var links = getLinks();
  // var archiveSheet = SpreadsheetApp.openById(links.?????).getSheets()[0];
  // var responseObjects = readResponseObjects(links.statusFormId);
  // responseObjects.forEach(obj => {
  //   console.log(obj);
  //   archiveSheet.appendRow([obj.kerberbos, obj.timestamp, obj.initiative, obj.effort, obj.epic, obj.status]);
  // });
  var form = FormApp.openById(links.statusFormId);
  form.deleteAllResponses();
}

function printForm() {
  var links = getLinks();
  var map = readForm(links.statusFormId);
  var kerberosMap = getKereberosMap(links.roverSheetId);
  map.forEach(statusList => {
    statusList.forEach(responseObject => {
      console.log(getStatusText(responseObject, kerberosMap));
    });
  });
}

function compileStatus() {
  var links = getLinks();
  copyTemplate(links.templateDocId, links.statusDocId);
  var statusMap = readForm(links.statusFormId);
  insertStatus(links.statusDocId, links.roverSheetId, statusMap);
}

function copyTemplate(templateDocId, statusDocId) {
  var template = DocumentApp.openById(templateDocId).getBody().copy();
  var doc = DocumentApp.openById(statusDocId);
  var body = doc.getBody();
  body.appendPageBreak();
  body.clear();
  var totalElements = template.getNumChildren();
  for (var index = 0; index < totalElements; index++) {
    var element = template.getChild(index);
    var type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH)
      body.appendParagraph(element.copy());
    else if (type == DocumentApp.ElementType.TABLE)
      body.appendTable(element.copy());
    else if (type == DocumentApp.ElementType.LIST_ITEM) {
      let inserted = body.appendListItem(element.getText());
      inserted.setAttributes(element.getAttributes());
      inserted.setListId(element).setGlyphType(element.getGlyphType()).setNestingLevel(element.getNestingLevel());
    } else
      throw new Error("Unknown element type: " + type);
  }
  body.getListItems().forEach(li => {
    if (li.getNestingLevel() === 0) {
      li.setGlyphType(DocumentApp.GlyphType.BULLET);
    } else if (li.getNestingLevel() === 1) {
      li.setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
    } else if (li.getNestingLevel() === 2) {
      li.setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
    }
  });
  body.getParagraphs().forEach(paragraph=>{
      if (paragraph.getType() === DocumentApp.ElementType.PARAGRAPH) {
        if (paragraph.getIndentStart() === null)
          paragraph.setIndentStart(0);
        if (paragraph.getIndentFirstLine() === null)
          paragraph.setIndentFirstLine(0);
      } else if (paragraph.getType() === DocumentApp.ElementType.PARAGRAPH) {
      }
  });
  doc.saveAndClose();
}

function readForm(formId) {
  var statusMap = new Map();
  var responseObjects = readResponseObjects(formId);
  // console.log("Found %s status entries", responseObjects.length);
  responseObjects.forEach(responseObject => {
    var mapKey;
    if (responseObject.effort) {
      mapKey = responseObject.initiative + "." + responseObject.effort;
    } else {
      mapKey = responseObject.initiative;
    }
    console.log("Status reported for %s", mapKey);
    var statusArray = getMapArray(statusMap, mapKey);
    statusArray.push(responseObject);
  });
  return statusMap;
}

function readResponseObjects(formId) {
  var form = FormApp.openById(formId);
  var responses = form.getResponses();
  // form.getItems().forEach(value => {
  //   console.log("item is %s", value.getTitle());
  // })
  var responseObjects = [];
  // console.log("Found %s status entries", responses.length);
  responses.forEach(response => {
    var responseObj = new Object();
    responseObj.timestamp = response.getTimestamp();
    responseObj.kerberbos = response.getRespondentEmail().split('@')[0];
    var answers = response.getItemResponses()
    responseObj.initiative = answers[0].getResponse();
    if (responseObj.initiative === "PTO / Learning / No Status") {
      responseObj.epic = responseObj.kerberbos;
      responseObj.status = "N/A";
    } else if (answers.length < 4) {
      responseObj.epic = answers[1].getResponse();
      responseObj.status = answers[2].getResponse();
    } else {
      responseObj.effort = answers[1].getResponse();
      responseObj.epic = answers[2].getResponse();
      responseObj.status = answers[3].getResponse();
    }
    responseObjects.push(responseObj);
  });
  return responseObjects;
}

function getMapArray(map, key) {
  var array = map.get(key);
  if (!array) {
    array = [];
    map.set(key, array);
    // console.log("There was no array for %s, but one has been created now", key);
  }
  return array;
}

function insertStatus(statusDocId, roverSheetId, statusMap) {
  var doc = DocumentApp.openById(statusDocId);
  var body = doc.getBody();
  var totalElements = body.getNumChildren();
  var listItemIndices = [];
  var knownStatusKeys = new Set();
  for (var index = 0; index < totalElements; index++) {
    if (body.getChild(index).getType() == DocumentApp.ElementType.LIST_ITEM && body.getChild(index).findText("%.*%")) {
      listItemIndices.push(index);
      var key = body.getChild(index).getText();
      key = key.substring(1, key.length - 1);
      knownStatusKeys.add(key);
    }
  }

  var otherStatusMap = new Map();
  statusMap.forEach((value, key) => {
    if (!knownStatusKeys.has(key)) {
      //Respondent must have selected other, and intended a different category for this item
      var newKey;
      var keyParts = key.split(".");
      if (keyParts.length === 1) {
        newKey = "Other";
      } else if (keyParts.length === 2) {
        newKey = keyParts[0] + ".Other";
      } else {
        throw new Error("Unexpected use of dot notation");
      }
      var mappedCategories = getMapArray(otherStatusMap, newKey);
      mappedCategories.push(key);
    }
  });

  for (var index = listItemIndices.length - 1; index >= 0; index--) {
    //Iterate backwards so inserting status does not make indices invalid
    var listItem = body.getChild(listItemIndices[index]);
    let attrs = listItem.getAttributes();
    attrs.FONT_SIZE = 12;
    var key = listItem.getText();
    key = key.substring(1, key.length - 1);
    // console.log("Actual key is %s", key);
    var statuses = statusMap.get(key);
    var kerberosMap = getKereberosMap(roverSheetId);
    if (statuses) {
      statuses.forEach(value => {
        let newListItem = body.insertListItem(listItemIndices[index], listItem.copy());
        let inserted = newListItem.editAsText();
        inserted.setText("");
        getStatusText(value, kerberosMap).forEach(part => {
          inserted.appendText(part.text);
          if (part.isLink) {
            inserted.appendText("\u200B");
            var endOffsetInclusive = inserted.getText().length - 2;
            var startOffset = endOffsetInclusive - part.text.length + 1;
            inserted.setLinkUrl(startOffset, endOffsetInclusive, part.url);
          }
        });
        newListItem.setAttributes(attrs);
      });
      statusMap.delete(key);
    }
    var mappedStatuses = otherStatusMap.get(key);
    if (mappedStatuses) {
      mappedStatuses.forEach(mappedKey => {
        var otherStatuses = statusMap.get(mappedKey);
        otherStatuses.forEach(value => {
          let newListItem = body.insertListItem(listItemIndices[index], listItem.copy());
          let inserted = newListItem.editAsText();
          inserted.setText(">>> " + mappedKey + " - ");
          getStatusText(value, kerberosMap).forEach(part => {
            inserted.appendText(part.text);
            if (part.isLink) {
              inserted.appendText("\u200B");
              var endOffsetInclusive = inserted.getText().length - 2;
              var startOffset = endOffsetInclusive - part.text.length + 1;
              inserted.setLinkUrl(startOffset, endOffsetInclusive, part.url);
            }
          });
          newListItem.setAttributes(attrs);
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
      listItem.editAsText().setText("");
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

function getKereberosMap(spreadsheetId) {
  var map = new Map();
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Rover");
  var columns = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, columns).getValues()[0];
  // console.log(header);
  var values = sheet.getRange(2, 1, sheet.getLastRow(), columns).getValues();
  values.forEach(value => {
    var record = new Map();
    for (var col = 0; col < columns; col++) {
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
  const regex = new RegExp("\\[([^)]*)\\]\\s*\\(([^\\)]*)\\)", "g");
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
