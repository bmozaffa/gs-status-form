const AI_SERVICE = "MaaS";

function onNewStatusEntry(event) {
  const response = event.response;
  let edited = getEditedResponse(response);
  const responseId = response.getId();
  let sheet = SpreadsheetApp.openById(getGlobalLinks().llmEditsSheetId).getSheetByName("Edits");
  var sheetValues = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  let updatedRow;
  for (var i = 0; i < sheetValues.length; i++) {
    if (sheetValues[i][0] === responseId && sheetValues[i][1] === response.getRespondentEmail()) {
      updatedRow = i + 1;
    }
  }
  if (updatedRow) {
    sheet.getRange(updatedRow, 3, 1, 3).setValues([[response.getTimestamp(), getResponseObject(response).status, edited]]);
  } else {
    sheet.appendRow([responseId, response.getRespondentEmail(), response.getTimestamp(), getResponseObject(response).status, edited]);
  }
}

const preprompt = 'You are a technical writer reporting engineering activities to engineering leaders and stakeholders. What follows is the jira epic or topic, as well as the status entry written by a software engineer. Rewrite this entry as one or more bullet points, as action items and without adding pronouns. Focus on the status part, and working the epic in there only if it makes sense and adds value. The output should only contain the content representing the combined status without any introductions or explanations.\n\nWhere a Markdown link is provided, include it in your revision, but do not add any links that do not exist in the provided text. If there are links, want them to be part of the text and for it to flow together seamlessly, instead of providing them as disjointed and dedicated words and phrases. If someone is reading the status and ignoring the hyperlinks, it should read normal to them. Where a link is provided without Markdown formatting, find appropriate text to represent the link and use Markdown syntax in the format of [Link text](Link URL) to write the link. As far as possible, avoid using the jira ticket number (usually in the form of ABCD-1234) or GitHub repository name, and instead use plain English words that make sense in their place.\n\nSome of these engineers are non-English speakers and you may need to carefully parse their writing and make more changes. Where the status entry is overly verbose, try to shorten it and if necessary, lose some of the extra details.\n\n';

function getEditedResponse(response) {
  const responseObject = getResponseObject(response);
  if (responseObject.status) {
    const prompt = preprompt + responseObject.epic + "\n\n" + responseObject.status;
    if (AI_SERVICE === "MaaS") {
      const maasResponse = callMaaS(prompt);
      Logger.log("This API call used up %s tokens", maasResponse.usage.total_tokens);
      const edited = maasResponse.choices[0].message.content;
      Logger.log("LLM edited the status entry to\n\n%s", edited);
      return edited;
    } else if (AI_SERVICE === "VertexAI") {
      const vertexResponse = callVertexAI(prompt);
      Logger.log("This API call used up %s tokens", vertexResponse.usageMetadata.totalTokenCount);
      const edited = vertexResponse.candidates[0].content.parts[0].text;
      Logger.log("LLM edited the status entry to\n\n%s", edited);
      return edited;
    } else {
      Logger.log("No AI Service configured!");
      return "";
    }
  } else {
    return "";
  }
}

function callVertexAI(prompt) {
  const accessToken = getVertexAiAccessToken();
  const llmConfig = JSON.parse(PropertiesService.getScriptProperties().getProperty('LLM_CONFIG'));
  const url = `https://${llmConfig.location}-aiplatform.googleapis.com/v1/projects/${llmConfig.projectId}/locations/${llmConfig.location}/publishers/google/models/${llmConfig.modelId}:generateContent`;

  const payload = {
    "contents": [
      {
        "role": "USER",
        "parts": [
          {"text": prompt}
        ]
      }
    ],
    "generationConfig": { // Optional config
      "temperature": 0.7,
      "maxOutputTokens": 4096
    }
    // "safetySettings": { ... } // Optional safety settings
  };

  const headers = {
    'Authorization': 'Bearer ' + accessToken,
    'Content-Type': 'application/json'
  };

  const options = {
    'method': 'POST',
    'contentType': 'application/json',
    'headers': headers,
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseJson = JSON.parse(response.getContentText());

  if (response.getResponseCode() >= 400) {
    Logger.log("Vertex AI API error: " + response.getContentText());
    throw new Error("Vertex AI API call failed: " + responseJson.error || "Unknown error");
  }

  Logger.log(JSON.stringify(responseJson));
  return responseJson;
}

/**
 * Gets a valid OAuth2 access token for the service account.
 * Uses CacheService to cache the token and reuse it until expiration.
 *
 * @return {string} The OAuth2 access token.
 */
function getVertexAiAccessToken() {
  const serviceAccountKey = JSON.parse(PropertiesService.getScriptProperties().getProperty('VERTEX_AI_KEY')); // Load key
  const cache = CacheService.getScriptCache();
  const cacheKey = 'vertexAiAccessToken_' + serviceAccountKey.client_email; // Unique key per service account

  let token = cache.get(cacheKey);
  if (token) {
    Logger.log('Using cached access token.');
    return token;
  }

  Logger.log('No valid cached token found, requesting a new one.');

  const jwt = createSignedJwt(serviceAccountKey); // Pass the key object
  if (!jwt) {
    throw new Error("Failed to create signed JWT.");
  }

  const tokenResponse = exchangeJwtForToken(serviceAccountKey, jwt); // Pass the key object
  if (!tokenResponse || !tokenResponse.access_token) {
    Logger.log('Failed to exchange JWT for token. Response: ' + JSON.stringify(tokenResponse));
    throw new Error("Failed to retrieve access token from Google.");
  }

  token = tokenResponse.access_token;
  const expiresIn = tokenResponse.expires_in; // Duration in seconds

  // Cache the token, subtracting a small buffer (e.g., 60 seconds)
  const cacheExpiration = Math.max(1, expiresIn - 60);
  cache.put(cacheKey, token, Math.min(cacheExpiration, 21600)); // Max cache duration is 6 hours
  Logger.log('New access token obtained and cached.');

  return token;
}

/**
 * Creates a signed JWT for service account authentication.
 *
 * @param {object} serviceAccountKey The parsed service account key object.
 * @return {string} The base64-encoded signed JWT.
 */
function createSignedJwt(serviceAccountKey) {
  const header = {
    alg: "RS256",
    typ: "JWT"
  };

  const now = Math.floor(Date.now() / 1000);
  const expires = now + 60; // Token valid for 1 hour

  const claimSet = {
    iss: serviceAccountKey.client_email,
    scope: "https://www.googleapis.com/auth/cloud-platform",
    aud: serviceAccountKey.token_uri,
    exp: expires,
    iat: now
  };

  const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(header), Utilities.Charset.UTF_8);
  const encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(claimSet), Utilities.Charset.UTF_8);
  const unsignedJwt = encodedHeader + "." + encodedClaimSet;

  try {
    // *** CRITICAL: Ensure private_key includes the \n newlines ***
    // If parsing removed them, signing will fail. Storing the full JSON string usually preserves them.
    const signatureBytes = Utilities.computeRsaSha256Signature(unsignedJwt, serviceAccountKey.private_key);
    const encodedSignature = Utilities.base64EncodeWebSafe(signatureBytes);
    return unsignedJwt + "." + encodedSignature;
  } catch (e) {
    Logger.log("Error signing JWT: " + e);
    // Check if the private key format might be the issue (often missing newlines after parsing/storage)
    if (serviceAccountKey.private_key && serviceAccountKey.private_key.indexOf('\n') === -1) {
      Logger.log("Warning: Private key retrieved from properties might be missing newline characters ('\\n'). Ensure the JSON string stored in Script Properties includes them verbatim from the original file.");
    } else if (!serviceAccountKey.private_key) {
      Logger.log("Error: private_key field was not found or is empty in the parsed service account key.");
    }
    return null;
  }
}

/**
 * Exchanges the signed JWT for an OAuth2 access token.
 *
 * @param {object} serviceAccountKey The parsed service account key object.
 * @param {string} signedJwt The signed JWT.
 * @return {object | null} The parsed JSON response from the token endpoint, or null on error.
 */
function exchangeJwtForToken(serviceAccountKey, signedJwt) {
  const tokenUrl = serviceAccountKey.token_uri;

  const payload = {
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion: signedJwt
  };

  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    payload: Object.keys(payload).map(key => key + '=' + encodeURIComponent(payload[key])).join('&'),
    muteHttpExceptions: true // Important to check response code manually
  };

  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      return JSON.parse(responseBody);
    } else {
      Logger.log(`Error exchanging JWT for token. Status: ${responseCode}, URL: ${tokenUrl}, Body: ${responseBody}`);
      return null;
    }
  } catch (e) {
    Logger.log("Exception during token exchange request: " + e);
    return null;
  }
}

function callMaaS(prompt) {
  const accessToken = PropertiesService.getScriptProperties().getProperty('MAAS_KEY');
  const url = 'https://mixtral-8x7b-instruct-v0-1-maas-apicast-production.apps.prod.rhoai.rh-aiservices-bu.com/v1/chat/completions';

  const payload = {
    "model": "mistralai/Mixtral-8x7B-Instruct-v0.1",
    "messages": [
      {
        "role": "user",
        "content": prompt
      }
    ],
    "temperature": 0.7,
    "max_tokens": 4096,
    "top_p": 1.0,
    "stream": false
  };

  const headers = {
    'Authorization': 'Bearer ' + accessToken,
    'Content-Type': 'application/json'
  };

  const options = {
    'method': 'POST',
    'contentType': 'application/json',
    'headers': headers,
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseJson = JSON.parse(response.getContentText());

  if (response.getResponseCode() >= 400) {
    Logger.log("MaaS AI API error: " + response.getContentText());
    throw new Error("MaaS AI API call failed: " + responseJson.error || "Unknown error");
  }

  Logger.log(JSON.stringify(responseJson));
  return responseJson;
}
