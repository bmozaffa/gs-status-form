function getLinks() {
  var links = new Object();
  links.statusFormId = "14g3I22FRYFV9kbE9csTdlbrKK3XbnHXo8rsMDg-Z_HI";
  links.templateDocId = "1BzP4CzYdhu_VQ2-0TrCsa9drSFiKxFFzLYy4m9jpaKU";
  links.statusDocId = "11UHpXCgpma_wDrXYfsD6VxVUAhHvrugSwMkPyKx2rlk";
  links.roverSheetId = "1i7y_tFpeO68SetmsU2t-C6LsFETuZtkJGY5AVZ2PHW8";
  return links;
}

function doGet(e) {
  var links = getLinks();
  var form = FormApp.openById(links.statusFormId);
  var statusEntries = form.getResponses().length;
  compileStatus();
  var response = "Successfully generated status report based on " + statusEntries + " form submissions";
  return HtmlService.createHtmlOutput(response);
}

function getMissingStatus() {
  var links = getLinks();
  var map = readForm(links.statusFormId);
  var kerberosMap = getKereberosMap(links.roverSheetId);
  map.forEach(statusList => {
    statusList.forEach(responseObject => {
      kerberosMap.delete(responseObject.kerberbos);
    });
  });
  console.log("Missing status for %s people", kerberosMap.size);
  kerberosMap.forEach((value, key) => {
    console.log(key);
  });
}

function archiveReports() {
  var links = getLinks();
  var archiveSheet = SpreadsheetApp.openById(links.roverSheetId).getSheets()[0];
  var responseObjects = readResponseObjects(links.statusFormId);
  responseObjects.forEach(obj => {
    console.log(obj);
    archiveSheet.appendRow([obj.kerberbos, obj.timestamp, obj.initiative, obj.effort, obj.epic, obj.status]);
  });
  // var form = FormApp.openById(links.statusFormId);
  // form.deleteAllResponses();
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
  body.clear();
  var totalElements = template.getNumChildren();
  for (var index = 0; index < totalElements; index++) {
    var element = template.getChild(index).copy();
    var type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH)
      doc.appendParagraph(element);
    else if (type == DocumentApp.ElementType.TABLE)
      doc.appendTable(element);
    else if (type == DocumentApp.ElementType.LIST_ITEM) {
      var glyphType = element.getGlyphType();
      doc.appendListItem(element);
      element.setGlyphType(glyphType);
    } else
      throw new Error("Unknown element type: " + type);
  }
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
    if (answers.length < 4) {
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
    var key = listItem.getText();
    key = key.substring(1, key.length - 1);
    // console.log("Actual key is %s", key);
    var statuses = statusMap.get(key);
    var kerberosMap = getKereberosMap(roverSheetId);
    if (statuses) {
      statuses.forEach(value => {
        var inserted = body.insertListItem(listItemIndices[index], listItem.copy());
        var statusText = getStatusText(value, kerberosMap);
        inserted.editAsText().setText(statusText);
      });
      statusMap.delete(key);
    }
    var mappedStatuses = otherStatusMap.get(key);
    if (mappedStatuses) {
      mappedStatuses.forEach(mappedKey => {
        var otherStatuses = statusMap.get(mappedKey);
        otherStatuses.forEach(value => {
          var inserted = body.insertListItem(listItemIndices[index], listItem.copy());
          var statusText = ">>> " + mappedKey + " - " + getStatusText(value, kerberosMap);
          inserted.editAsText().setText(statusText);
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
  var respondent = kerberosMap.get(responseObject.kerberbos);
  return responseObject.epic + ":\n" + responseObject.status + "\n[By " + respondent + " on " + responseObject.timestamp + "]\n";
}

function getKereberosMap(spreadsheetId) {
  var map = new Map();
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Rover");
  var values = sheet.getRange(2, 1, sheet.getLastRow(), 2).getValues();
  values.forEach(value => {
    map.set(value[1], value[0]);
  });
  return map;
}
