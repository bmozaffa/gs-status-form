const AI_SERVICE = "GeminiAPI";

function simpleTestLLM() {
  console.log( getModelResponse("What is the capital of France?") );
}

function manuallyInvokeLLM() {
  return generateEditsByLLM(10469);
}

function onStatusFormSubmission(event) {
  return generateEditsByLLM(event.range.rowStart);
}

function preserveOriginalUrls(responseObject, edited) {
  const originalText = getStatusEntry(responseObject);

  // Extract URLs from both original and edited text
  const urlRegex = /https?:\/\/[^\s\)]+/g;
  const originalUrls = originalText.match(urlRegex) || [];
  const editedUrls = edited.match(urlRegex) || [];

  let result = edited;

  // For each URL in the edited text, find the most similar original URL
  editedUrls.forEach(editedUrl => {
    let bestMatch = null;
    let minDistance = Infinity;

    originalUrls.forEach(originalUrl => {
      const distance = calculateLevenshteinDistance(editedUrl, originalUrl);
      const maxLength = Math.max(editedUrl.length, originalUrl.length);
      const similarity = 1 - (distance / maxLength);

      // Consider URLs similar if they have >85% similarity and differ by at most 3 characters
      if (similarity > 0.85 && distance <= 3 && distance < minDistance) {
        minDistance = distance;
        bestMatch = originalUrl;
      }
    });

    // If we found a similar but different URL, replace it with the original
    if (bestMatch && editedUrl !== bestMatch) {
      Logger.log("Replacing URL [%s] with suspected original of [%s] under the assumption that the LLM edit has corrupted it", editedUrl, bestMatch);
      result = result.replace(editedUrl, bestMatch);
    }
  });

  return result;
}

function calculateLevenshteinDistance(str1, str2) {
  const matrix = [];

  // Create matrix
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }

  // Fill matrix
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  return matrix[str2.length][str1.length];
}

function generateEditsByLLM(row) {
  const archives = SpreadsheetApp.openById(getGlobalLinks().statusArchivesSheetId).getSheets()[0];
  const archivesRow = archives.getRange(row, 1, 1, archives.getLastColumn()).getValues()[0];
  const responseObject = {
    timestamp: archivesRow[0],
    email: archivesRow[1],
    initiative: archivesRow[2],
    epic: archivesRow[6],
    status: archivesRow[7]
  };
  let edited = "";
  if (responseObject.status) {
    edited = getModelResponse( preprompt + getStatusEntry(responseObject) );
    edited = preserveOriginalUrls(responseObject, edited);
  }
  let sheet = SpreadsheetApp.openById(getGlobalLinks().llmEditsSheetId).getSheetByName("Edits");
  var sheetValues = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  let updatedRow;
  for (var i = 0; i < sheetValues.length; i++) {
    if (sheetValues[i][0] === row) {
      updatedRow = i + 1;
    }
  }
  if (updatedRow) {
    sheet.getRange(updatedRow, 3, 1, 3).setValues([[responseObject.timestamp, getStatusEntry( responseObject ), edited]]);
  } else {
    sheet.appendRow([row, responseObject.email, responseObject.timestamp, responseObject.status, edited]);
  }
  return edited;
}

const preprompt = 'You are a technical writer reporting engineering activities to engineering leaders and stakeholders. I will provide you with either a Jira epic or a high level description of the work, followed by the status entry written by a software engineer. Rewrite these as one or more clear, concise bullet points detailing engineering activities. Avoid using pronouns. Focus on the status part, and work in the epic/context only if it makes sense and adds value. The output should only contain the content representing the combined status, without any introductions or explanations.\n\nAny string consisting of a few characters followed by a dash and then digits (e.g., APPENG-1234) should be assumed to be a Jira ticket and formatted as a Markdown link, preceded by https://issues.redhat.com/browse/. Similarly, 8-digit support case numbers should be formatted as Markdown links, preceded by https://access.redhat.com/support/cases/#/case/.\n\nWhen a Markdown link is provided, integrate it naturally into the most relevant existing words in the rewritten status, while keeping in mind that GitHub pull requests typically provide updates about bug fixes and are better fits for terms like fix or update or change, and Jira tickets are typically the bug or feature enhancement description and their links better fit such terms. **The hyperlink text should be as concise as possible, ideally 1-3 words, representing the core concept of the link without being overly descriptive.** Do not add new words or phrases solely to accommodate a link. The link text must be a natural part of the sentence structure, ensuring the status reads like normal English even if the hyperlink is ignored. If a link is provided without Markdown formatting, find the most appropriate and **brief** descriptive text from the original status, or use very minimal, direct wording, and format it as [Link text](Link URL). Avoid using Jira ticket numbers or GitHub repository names in the text; instead, use plain English that makes sense in context. **Only use Markdown for hyperlinks; do not use any other Markdown formatting like bolding or code blocks. Specifically, do not use backticks (``) for model names or other proper nouns; treat them as normal text.**\n\nSome of these engineers are non-English speakers and you may need to carefully parse their writing and make more changes. Prioritize conciseness and eliminate superfluous details. When an issue (e.g., from an epic or Jira ticket) is followed by a status describing its resolution, combine these details into a **single, non-redundant statement**. Ensure each bullet point ends without punctuation (no periods, commas, or semicolons at the very end), and use semicolons or rephrase sentences to avoid periods within the bullet point itself. Now, rewrite this:\n\n';

function getStatusEntry(responseObject) {
  return responseObject.epic + "\n\n" + responseObject.status;
}

function getModelResponse(prompt) {
  if (AI_SERVICE === "MaaS") {
    const maasResponse = callMaaS(prompt);
    Logger.log("This API call used up %s tokens", maasResponse.usage.total_tokens);
    const edited = maasResponse.choices[0].message.content;
    Logger.log("LLM edited the status entry to\n\n%s", edited);
    return edited;
  } else if (AI_SERVICE === "DeployedModel") {
    const localResponse = callDeployedModel(prompt);
    Logger.log("This API call used up %s tokens", localResponse.usage.total_tokens);
    const edited = localResponse.choices[0].message.content;
    Logger.log("LLM edited the status entry to\n\n%s", edited);
    return edited;
  } else if (AI_SERVICE === "VertexAI") {
    const vertexResponse = callVertexAI(prompt);
    Logger.log("This API call used up %s tokens", vertexResponse.usageMetadata.totalTokenCount);
    const edited = vertexResponse.candidates[0].content.parts[0].text;
    Logger.log("LLM edited the status entry to\n\n%s", edited);
    return edited;
  } else if (AI_SERVICE === "GeminiAPI") {
    const geminiResponse = callGeminiAPI(prompt);
    Logger.log("This API call used up %s tokens", geminiResponse.usageMetadata.totalTokenCount);
    const edited = geminiResponse.candidates[0].content.parts[0].text;
    Logger.log("LLM edited the status entry to\n\n%s", edited);
    return edited;
  } else {
    Logger.log("No AI Service configured!");
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
    Logger.log(token);
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
  const expires = now + 3600; // Token valid for 1 hour

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

function callDeployedModel(prompt) {
  const baseUrl = PropertiesService.getScriptProperties().getProperty('DEPLOYED_MODEL_URL');
  const modelName = PropertiesService.getScriptProperties().getProperty('DEPLOYED_MODEL_NAME');
  const url = baseUrl + '/v1/chat/completions';

  console.log(url);
  console.log(modelName);
  const payload = {
    "model": modelName,
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
    Logger.log("Deployed model API error: " + response.getContentText());
    throw new Error("Deployed model API call failed: " + responseJson.error || "Unknown error");
  }

  Logger.log(JSON.stringify(responseJson));
  return responseJson;
}

function callGeminiAPI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const generationConfig = {
    temperature: 1.5,
    candidateCount: 1,
    topP: 0.95,
    topK: 40,
    responseMimeType: 'text/plain',
  };

  const payload = {
    generationConfig,
    contents: [
      {
        parts: [
          { text: prompt },
        ],
      },
    ],
  };

  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() >= 400) {
    Logger.log("Deployed model API error: " + response.getContentText());
    throw new Error("Deployed model API call failed: " + responseJson.error || "Unknown error");
  }

  const responseJson = JSON.parse(response);
  Logger.log(JSON.stringify(responseJson));
  return responseJson;
}
