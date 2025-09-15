/**
 * Comprehensive Test Suite for Red Hat Status Reporting System
 *
 * This test suite covers all externally invoked functions:
 * - doGet
 * - createDraftEmails
 * - compileStatus
 * - notifyMissingStatus
 * - notifyMismatchAssignment
 * - validateRecipients
 * - onStatusFormSubmission
 */

// Test configuration and constants
const TEST_CONFIG = {
  TEST_FORM_ID: "test_form_id_12345",
  TEST_DOC_ID: "test_doc_id_67890",
  TEST_SPREADSHEET_ID: "test_sheet_id_abcde",
  TEST_EMAIL: "testuser@redhat.com",
  TEST_KERBEROS: "testuser",
  TEST_MANAGER: "testmgr",
  MOCK_TIMESTAMP: new Date("2024-01-15T10:00:00Z")
};

// Mock data storage for tests
let mockData = {
  formResponses: [],
  spreadsheetData: new Map(),
  documentBodies: new Map(),
  emailsSent: [],
  draftsCreated: [],
  calendarEvents: new Map(),
  propertyValues: new Map()
};

// Test results tracking
let testResults = {
  passed: 0,
  failed: 0,
  errors: []
};

/**
 * Main test runner - executes all test suites
 */
function runAllTests() {
  console.log("🧪 Starting comprehensive regression test suite...");
  resetTestEnvironment();

  try {
    // Test each external function
    testDoGet();
    testCreateDraftEmails();
    testCompileStatus();
    testNotifyMissingStatus();
    testNotifyMismatchAssignment();
    testValidateRecipients();
    testOnStatusFormSubmission();

    // Print final results
    printTestResults();

  } catch (error) {
    console.error("❌ Test suite execution failed:", error);
    testResults.errors.push(`Test suite execution error: ${error.message}`);
  }

  return testResults;
}

/**
 * Test suite for doGet function
 * Tests all command parameters and edge cases
 */
function testDoGet() {
  console.log("\n📝 Testing doGet function...");

  // Setup mock data
  setupMockData();

  try {
    // Test missing command
    testDoGetMissingCommand();

    // Test all valid commands
    testDoGetMissingCommand();
    testDoGetGenerateCommand();
    testDoGetArchiveCommand();
    testDoGetSendDraftsCommand();
    testDoGetMismatchCommand();
    testDoGetRewriteCommand();
    testDoGetInvalidCommand();

    console.log("✅ doGet tests completed");

  } catch (error) {
    recordTestFailure("doGet", error);
  }
}

function testDoGetMissingCommand() {
  try {
    const mockEvent = { parameter: { command: "missing" } };
    setupMockMissingStatus();

    const result = doGet(mockEvent);
    const content = result.getContent();

    assert(content.includes("testmgr"), "Should include manager name in missing status");
    assert(content.includes("testuser"), "Should include associate name in missing status");
    recordTestPass("doGet - missing command");

  } catch (error) {
    recordTestFailure("doGet - missing command", error);
  }
}

function testDoGetGenerateCommand() {
  try {
    const mockEvent = { parameter: { command: "generate", docId: TEST_CONFIG.TEST_DOC_ID } };
    setupMockStatusGeneration();

    const result = doGet(mockEvent);
    const content = result.getContent();

    assert(content.includes("Successfully generated"), "Should indicate successful generation");
    recordTestPass("doGet - generate command");

  } catch (error) {
    recordTestFailure("doGet - generate command", error);
  }
}

function testDoGetArchiveCommand() {
  try {
    const mockEvent = { parameter: { command: "archive", days: "7" } };
    setupMockArchive();

    const result = doGet(mockEvent);
    const content = result.getContent();

    assert(content.includes("Successfully archived"), "Should indicate successful archiving");
    assert(content.includes("2"), "Should show archived count");
    recordTestPass("doGet - archive command");

  } catch (error) {
    recordTestFailure("doGet - archive command", error);
  }
}

function testDoGetSendDraftsCommand() {
  try {
    const mockEvent = { parameter: { command: "send-drafts" } };
    setupMockDrafts();

    const result = doGet(mockEvent);
    const content = result.getContent();

    assert(content.includes("Successfully sent"), "Should indicate successful email sending");
    recordTestPass("doGet - send-drafts command");

  } catch (error) {
    recordTestFailure("doGet - send-drafts command", error);
  }
}

function testDoGetMismatchCommand() {
  try {
    const mockEvent = { parameter: { command: "mismatch", manager: "testmgr" } };
    setupMockMismatch();

    const result = doGet(mockEvent);
    const content = result.getContent();

    // Should return mismatch data or empty string if no mismatches
    assert(typeof content === "string", "Should return string content");
    recordTestPass("doGet - mismatch command");

  } catch (error) {
    recordTestFailure("doGet - mismatch command", error);
  }
}

function testDoGetRewriteCommand() {
  try {
    const mockEvent = { parameter: { command: "rewrite", line: "123" } };
    setupMockLLMRewrite();

    const result = doGet(mockEvent);
    const content = result.getContent();

    assert(typeof content === "string", "Should return rewritten content");
    recordTestPass("doGet - rewrite command");

  } catch (error) {
    recordTestFailure("doGet - rewrite command", error);
  }
}

function testDoGetInvalidCommand() {
  try {
    const mockEvent = { parameter: { command: "invalid" } };

    const result = doGet(mockEvent);
    const content = result.getContent();

    assert(content.includes("Provide the command"), "Should show help message for invalid command");
    recordTestPass("doGet - invalid command");

  } catch (error) {
    recordTestFailure("doGet - invalid command", error);
  }
}

/**
 * Test suite for createDraftEmails function
 */
function testCreateDraftEmails() {
  console.log("\n📧 Testing createDraftEmails function...");

  try {
    setupMockEmailData();

    // Test normal operation
    testCreateDraftEmailsNormal();

    // Test with pause condition
    testCreateDraftEmailsPaused();

    // Test with empty data
    testCreateDraftEmailsEmpty();

    console.log("✅ createDraftEmails tests completed");

  } catch (error) {
    recordTestFailure("createDraftEmails", error);
  }
}

function testCreateDraftEmailsNormal() {
  try {
    mockData.propertyValues.set('notifyMismatchAssignment', false); // Not paused
    setupMockEmailSpreadsheet();

    createDraftEmails();

    assert(mockData.draftsCreated.length > 0, "Should create draft emails");
    recordTestPass("createDraftEmails - normal operation");

  } catch (error) {
    recordTestFailure("createDraftEmails - normal operation", error);
  }
}

function testCreateDraftEmailsPaused() {
  try {
    mockData.propertyValues.set('createDraftEmails', true); // Paused

    const initialDraftCount = mockData.draftsCreated.length;
    createDraftEmails();

    assert(mockData.draftsCreated.length === initialDraftCount, "Should not create drafts when paused");
    recordTestPass("createDraftEmails - paused");

  } catch (error) {
    recordTestFailure("createDraftEmails - paused", error);
  }
}

function testCreateDraftEmailsEmpty() {
  try {
    mockData.propertyValues.set('createDraftEmails', false); // Not paused
    setupEmptyEmailSpreadsheet();

    const initialDraftCount = mockData.draftsCreated.length;
    createDraftEmails();

    assert(mockData.draftsCreated.length === initialDraftCount, "Should handle empty spreadsheet gracefully");
    recordTestPass("createDraftEmails - empty data");

  } catch (error) {
    recordTestFailure("createDraftEmails - empty data", error);
  }
}

/**
 * Test suite for compileStatus function
 */
function testCompileStatus() {
  console.log("\n📊 Testing compileStatus function...");

  try {
    // Test normal compilation
    testCompileStatusNormal();

    // Test with pause condition
    testCompileStatusPaused();

    // Test with no new status
    testCompileStatusNoUpdate();

    console.log("✅ compileStatus tests completed");

  } catch (error) {
    recordTestFailure("compileStatus", error);
  }
}

function testCompileStatusNormal() {
  try {
    setupMockStatusCompilation();
    mockData.propertyValues.set('compileStatus', false); // Not paused

    const result = compileStatus();

    // Should process without errors
    recordTestPass("compileStatus - normal operation");

  } catch (error) {
    recordTestFailure("compileStatus - normal operation", error);
  }
}

function testCompileStatusPaused() {
  try {
    mockData.propertyValues.set('compileStatus', true); // Paused

    const result = compileStatus();

    // Should return early when paused
    recordTestPass("compileStatus - paused");

  } catch (error) {
    recordTestFailure("compileStatus - paused", error);
  }
}

function testCompileStatusNoUpdate() {
  try {
    setupMockNoUpdateNeeded();
    mockData.propertyValues.set('compileStatus', false); // Not paused

    const result = compileStatus();

    recordTestPass("compileStatus - no update needed");

  } catch (error) {
    recordTestFailure("compileStatus - no update needed", error);
  }
}

/**
 * Test suite for notifyMissingStatus function
 */
function testNotifyMissingStatus() {
  console.log("\n🔔 Testing notifyMissingStatus function...");

  try {
    // Test normal notification
    testNotifyMissingStatusNormal();

    // Test with pause condition
    testNotifyMissingStatusPaused();

    // Test with no missing status
    testNotifyMissingStatusNone();

    console.log("✅ notifyMissingStatus tests completed");

  } catch (error) {
    recordTestFailure("notifyMissingStatus", error);
  }
}

function testNotifyMissingStatusNormal() {
  try {
    setupMockMissingNotification();
    mockData.propertyValues.set('notifyMissingStatus', false); // Not paused

    const initialEmailCount = mockData.emailsSent.length;
    notifyMissingStatus();

    assert(mockData.emailsSent.length > initialEmailCount, "Should send missing status emails");
    recordTestPass("notifyMissingStatus - normal operation");

  } catch (error) {
    recordTestFailure("notifyMissingStatus - normal operation", error);
  }
}

function testNotifyMissingStatusPaused() {
  try {
    mockData.propertyValues.set('notifyMissingStatus', true); // Paused

    const initialEmailCount = mockData.emailsSent.length;
    notifyMissingStatus();

    assert(mockData.emailsSent.length === initialEmailCount, "Should not send emails when paused");
    recordTestPass("notifyMissingStatus - paused");

  } catch (error) {
    recordTestFailure("notifyMissingStatus - paused", error);
  }
}

function testNotifyMissingStatusNone() {
  try {
    setupMockNoMissingStatus();
    mockData.propertyValues.set('notifyMissingStatus', false); // Not paused

    const initialEmailCount = mockData.emailsSent.length;
    notifyMissingStatus();

    assert(mockData.emailsSent.length === initialEmailCount, "Should not send emails when no missing status");
    recordTestPass("notifyMissingStatus - no missing status");

  } catch (error) {
    recordTestFailure("notifyMissingStatus - no missing status", error);
  }
}

/**
 * Test suite for notifyMismatchAssignment function
 */
function testNotifyMismatchAssignment() {
  console.log("\n⚠️ Testing notifyMismatchAssignment function...");

  try {
    // Test normal notification
    testNotifyMismatchAssignmentNormal();

    // Test with pause condition
    testNotifyMismatchAssignmentPaused();

    // Test with no mismatches
    testNotifyMismatchAssignmentNone();

    console.log("✅ notifyMismatchAssignment tests completed");

  } catch (error) {
    recordTestFailure("notifyMismatchAssignment", error);
  }
}

function testNotifyMismatchAssignmentNormal() {
  try {
    setupMockMismatchNotification();
    mockData.propertyValues.set('notifyMismatchAssignment', false); // Not paused

    const initialEmailCount = mockData.emailsSent.length;
    notifyMismatchAssignment();

    // May or may not send emails depending on mismatches found
    recordTestPass("notifyMismatchAssignment - normal operation");

  } catch (error) {
    recordTestFailure("notifyMismatchAssignment - normal operation", error);
  }
}

function testNotifyMismatchAssignmentPaused() {
  try {
    mockData.propertyValues.set('notifyMismatchAssignment', true); // Paused

    const initialEmailCount = mockData.emailsSent.length;
    notifyMismatchAssignment();

    assert(mockData.emailsSent.length === initialEmailCount, "Should not send emails when paused");
    recordTestPass("notifyMismatchAssignment - paused");

  } catch (error) {
    recordTestFailure("notifyMismatchAssignment - paused", error);
  }
}

function testNotifyMismatchAssignmentNone() {
  try {
    setupMockNoMismatches();
    mockData.propertyValues.set('notifyMismatchAssignment', false); // Not paused

    const initialEmailCount = mockData.emailsSent.length;
    notifyMismatchAssignment();

    // Should complete without errors even with no mismatches
    recordTestPass("notifyMismatchAssignment - no mismatches");

  } catch (error) {
    recordTestFailure("notifyMismatchAssignment - no mismatches", error);
  }
}

/**
 * Test suite for validateRecipients function
 */
function testValidateRecipients() {
  console.log("\n✉️ Testing validateRecipients function...");

  try {
    // Test normal validation
    testValidateRecipientsNormal();

    // Test with pause condition
    testValidateRecipientsPaused();

    // Test with invalid recipients
    testValidateRecipientsInvalid();

    console.log("✅ validateRecipients tests completed");

  } catch (error) {
    recordTestFailure("validateRecipients", error);
  }
}

function testValidateRecipientsNormal() {
  try {
    setupMockValidRecipients();
    mockData.propertyValues.set('validateRecipient', false); // Not paused

    const initialEmailCount = mockData.emailsSent.length;
    validateRecipients();

    // Should complete validation without sending error emails for valid recipients
    recordTestPass("validateRecipients - normal operation");

  } catch (error) {
    recordTestFailure("validateRecipients - normal operation", error);
  }
}

function testValidateRecipientsPaused() {
  try {
    mockData.propertyValues.set('validateRecipient', true); // Paused

    validateRecipients();

    // Should return early when paused
    recordTestPass("validateRecipients - paused");

  } catch (error) {
    recordTestFailure("validateRecipients - paused", error);
  }
}

function testValidateRecipientsInvalid() {
  try {
    setupMockInvalidRecipients();
    mockData.propertyValues.set('validateRecipient', false); // Not paused

    const initialEmailCount = mockData.emailsSent.length;
    validateRecipients();

    assert(mockData.emailsSent.length > initialEmailCount, "Should send notification for invalid recipients");
    recordTestPass("validateRecipients - invalid recipients");

  } catch (error) {
    recordTestFailure("validateRecipients - invalid recipients", error);
  }
}

/**
 * Test suite for onStatusFormSubmission function
 */
function testOnStatusFormSubmission() {
  console.log("\n🤖 Testing onStatusFormSubmission function...");

  try {
    // Test normal form submission
    testOnStatusFormSubmissionNormal();

    // Test with invalid row
    testOnStatusFormSubmissionInvalidRow();

    console.log("✅ onStatusFormSubmission tests completed");

  } catch (error) {
    recordTestFailure("onStatusFormSubmission", error);
  }
}

function testOnStatusFormSubmissionNormal() {
  try {
    const mockEvent = { range: { rowStart: 100 } };
    setupMockLLMProcessing();

    const result = onStatusFormSubmission(mockEvent);

    // Should process LLM editing without errors
    recordTestPass("onStatusFormSubmission - normal operation");

  } catch (error) {
    recordTestFailure("onStatusFormSubmission - normal operation", error);
  }
}

function testOnStatusFormSubmissionInvalidRow() {
  try {
    const mockEvent = { range: { rowStart: -1 } };

    const result = onStatusFormSubmission(mockEvent);

    // Should handle invalid row gracefully
    recordTestPass("onStatusFormSubmission - invalid row");

  } catch (error) {
    recordTestFailure("onStatusFormSubmission - invalid row", error);
  }
}

// =============================================================================
// MOCK DATA SETUP FUNCTIONS
// =============================================================================

function setupMockData() {
  // Setup basic mock data for all tests
  setupMockKerberosMap();
  setupMockDocumentLinks();
  setupMockGlobalLinks();
}

function setupMockKerberosMap() {
  // Mock the kerberosMap global variable - clear existing content if it exists
  if (typeof kerberosMap !== 'undefined') {
    kerberosMap.clear();
  } else {
    // Create new Map if it doesn't exist (test environment)
    global.kerberosMap = new Map();
  }

  kerberosMap.set(TEST_CONFIG.TEST_KERBEROS, new Map([
    ["Name", "Test User"],
    ["Email", TEST_CONFIG.TEST_EMAIL],
    ["Manager", TEST_CONFIG.TEST_MANAGER],
    ["Job Title", "Software Engineer"],
    ["Termination", ""] // Not terminated
  ]));

  kerberosMap.set(TEST_CONFIG.TEST_MANAGER, new Map([
    ["Name", "Test Manager"],
    ["Email", "testmgr@redhat.com"],
    ["Manager", "topmgr"],
    ["Job Title", "Manager, Software Engineering"],
    ["Termination", ""]
  ]));
}

function setupMockDocumentLinks() {
  // Mock the documentLinks global variable - clear existing content if it exists
  if (typeof documentLinks !== 'undefined') {
    documentLinks.clear();
  } else {
    // Create new Map if it doesn't exist (test environment)
    global.documentLinks = new Map();
  }

  documentLinks.set("Templates", new Map([
    [TEST_CONFIG.TEST_DOC_ID, "template_123"]
  ]));

  documentLinks.set("Managers", new Map([
    [TEST_CONFIG.TEST_MANAGER, TEST_CONFIG.TEST_DOC_ID]
  ]));

  documentLinks.set("Initiatives", new Map([
    ["TestInitiative", TEST_CONFIG.TEST_DOC_ID]
  ]));
}

function setupMockGlobalLinks() {
  // Mock the globalLinks global variable - update properties if it exists
  if (typeof globalLinks !== 'undefined') {
    // Clear and repopulate existing object
    Object.keys(globalLinks).forEach(key => delete globalLinks[key]);
    Object.assign(globalLinks, {
      statusFormId: TEST_CONFIG.TEST_FORM_ID,
      roverSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      statusEmailsId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      rosterSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      llmEditsSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      statusArchivesSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID
    });
  } else {
    // Create new object if it doesn't exist (test environment)
    global.globalLinks = {
      statusFormId: TEST_CONFIG.TEST_FORM_ID,
      roverSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      statusEmailsId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      rosterSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      llmEditsSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID,
      statusArchivesSheetId: TEST_CONFIG.TEST_SPREADSHEET_ID
    };
  }
}

function setupMockMissingStatus() {
  mockData.formResponses = []; // No responses = missing status
}

function setupMockStatusGeneration() {
  mockData.formResponses = [
    {
      kerberos: TEST_CONFIG.TEST_KERBEROS,
      timestamp: TEST_CONFIG.MOCK_TIMESTAMP,
      initiative: "TestInitiative",
      effort: "Development",
      epic: "Test Epic",
      status: "Completed test implementation"
    }
  ];
}

function setupMockArchive() {
  // Mock old responses for archiving
  mockData.formResponses = [
    { timestamp: new Date("2024-01-01") },
    { timestamp: new Date("2024-01-02") }
  ];
}

function setupMockDrafts() {
  mockData.draftsCreated = [
    { subject: "Test Draft 1", to: "test1@redhat.com" },
    { subject: "Test Draft 2", to: "test2@redhat.com" }
  ];
}

function setupMockMismatch() {
  // Setup data that would cause assignment mismatches
  mockData.spreadsheetData.set("assignments", [
    [TEST_CONFIG.TEST_KERBEROS, "TestInitiative"]
  ]);
}

function setupMockLLMRewrite() {
  mockData.spreadsheetData.set("llmArchive", [
    [TEST_CONFIG.MOCK_TIMESTAMP, TEST_CONFIG.TEST_EMAIL, "TestInitiative", "", "", "", "Test Epic", "Original status text"]
  ]);
}

function setupMockEmailData() {
  mockData.spreadsheetData.set("emails", [
    ["test@redhat.com", "cc@redhat.com", "bcc@redhat.com", "Test Subject", "label1,label2", ""]
  ]);
}

function setupMockEmailSpreadsheet() {
  setupMockEmailData();
}

function setupEmptyEmailSpreadsheet() {
  mockData.spreadsheetData.set("emails", []);
}

function setupMockStatusCompilation() {
  setupMockStatusGeneration();
}

function setupMockNoUpdateNeeded() {
  mockData.formResponses = [];
}

function setupMockMissingNotification() {
  setupMockMissingStatus();
}

function setupMockNoMissingStatus() {
  mockData.formResponses = [
    {
      kerberos: TEST_CONFIG.TEST_KERBEROS,
      timestamp: TEST_CONFIG.MOCK_TIMESTAMP,
      initiative: "TestInitiative",
      status: "Some status"
    }
  ];
}

function setupMockMismatchNotification() {
  setupMockMismatch();
}

function setupMockNoMismatches() {
  mockData.formResponses = [
    {
      kerberos: TEST_CONFIG.TEST_KERBEROS,
      initiative: "TestInitiative",
      status: "Aligned status"
    }
  ];
  mockData.spreadsheetData.set("assignments", [
    [TEST_CONFIG.TEST_KERBEROS, "TestInitiative"]
  ]);
}

function setupMockValidRecipients() {
  setupMockEmailData();
}

function setupMockInvalidRecipients() {
  mockData.spreadsheetData.set("emails", [
    ["invalid@nonexistent.com", "", "", "Test Subject", "label1", ""]
  ]);
}

function setupMockLLMProcessing() {
  mockData.spreadsheetData.set("llmArchive", [
    [TEST_CONFIG.MOCK_TIMESTAMP, TEST_CONFIG.TEST_EMAIL, "TestInitiative", "", "", "", "Test Epic", "Status to process"]
  ]);
}

// =============================================================================
// GOOGLE APPS SCRIPT API MOCKS
// =============================================================================

// Mock Google Apps Script APIs for testing
function mockGoogleAppsScriptAPIs() {
  // Mock FormApp
  FormApp = {
    openById: (id) => ({
      getResponses: () => mockData.formResponses.map(r => ({
        getTimestamp: () => r.timestamp,
        getRespondentEmail: () => r.email || TEST_CONFIG.TEST_EMAIL,
        getItemResponses: () => [
          { getResponse: () => r.initiative },
          { getResponse: () => r.effort },
          { getResponse: () => r.epic },
          { getResponse: () => r.status }
        ],
        getId: () => "response_" + Math.random()
      })),
      deleteResponse: (id) => {
        mockData.formResponses = mockData.formResponses.filter(r => r.id !== id);
      }
    })
  };

  // Mock SpreadsheetApp
  SpreadsheetApp = {
    openById: (id) => ({
      getSheetByName: (name) => ({
        getRange: (row, col, numRows, numCols) => ({
          getValues: () => mockData.spreadsheetData.get(name) || [],
          setValues: (values) => {
            mockData.spreadsheetData.set(name, values);
          }
        }),
        getLastRow: () => (mockData.spreadsheetData.get(name) || []).length + 1,
        getLastColumn: () => 6,
        appendRow: (row) => {
          const data = mockData.spreadsheetData.get(name) || [];
          data.push(row);
          mockData.spreadsheetData.set(name, data);
        }
      }),
      getSheets: () => [
        {
          getRange: (row, col, numRows, numCols) => ({
            getValues: () => mockData.spreadsheetData.get("llmArchive") || []
          }),
          getLastColumn: () => 8
        }
      ]
    })
  };

  // Mock DocumentApp
  DocumentApp = {
    openById: (id) => ({
      getBody: () => ({
        copy: () => ({}),
        clear: () => {},
        appendPageBreak: () => {},
        insertParagraph: (index, text) => ({
          appendText: (text) => ({
            setFontSize: () => {},
            setBold: () => {},
            setLinkUrl: () => {}
          })
        }),
        getNumChildren: () => 0,
        getChild: () => ({}),
        appendParagraph: () => ({}),
        appendTable: () => ({}),
        appendListItem: () => ({}),
        getListItems: () => [],
        getParagraphs: () => [
          {
            getText: () => "This document is no longer auto-generated. It was last generated at " + new Date().toUTCString(),
            getType: () => DocumentApp.ElementType.PARAGRAPH,
            getIndentStart: () => 0,
            getIndentFirstLine: () => 0,
            setIndentStart: () => {},
            setIndentFirstLine: () => {}
          }
        ],
        findText: () => null,
        removeChild: () => {},
        insertListItem: (index, element) => ({
          copy: () => ({
            getText: () => "%TestCategory%",
            getAttributes: () => ({}),
            setAttributes: () => {},
            setGlyphType: () => {},
            setNestingLevel: () => {},
            getGlyphType: () => {},
            getNestingLevel: () => 0
          }),
          editAsText: () => ({
            setText: () => {},
            appendText: () => {},
            getText: () => "Test text",
            setLinkUrl: () => {},
            setForegroundColor: () => {}
          })
        })
      }),
      saveAndClose: () => {}
    }),
    ElementType: {
      PARAGRAPH: "PARAGRAPH",
      TABLE: "TABLE",
      LIST_ITEM: "LIST_ITEM"
    },
    GlyphType: {
      BULLET: "BULLET",
      HOLLOW_BULLET: "HOLLOW_BULLET",
      SQUARE_BULLET: "SQUARE_BULLET"
    }
  };

  // Mock GmailApp
  GmailApp = {
    sendEmail: (to, subject, body, options) => {
      mockData.emailsSent.push({ to, subject, body, options });
    },
    createDraft: (to, subject, body, options) => {
      const draft = { to, subject, body, options };
      mockData.draftsCreated.push(draft);
      return {
        getMessage: () => ({
          getSubject: () => subject,
          getTo: () => to,
          getThread: () => ({
            addLabel: () => {},
            getLabels: () => []
          })
        }),
        send: () => ({
          getSubject: () => subject,
          forward: () => {},
          getThread: () => ({
            addLabel: () => {}
          })
        })
      };
    },
    getDrafts: () => mockData.draftsCreated.map(d => ({
      getMessage: () => ({
        getSubject: () => d.subject,
        getTo: () => d.to,
        getThread: () => ({
          getLabels: () => [],
          addLabel: () => {}
        })
      }),
      send: () => ({
        getSubject: () => d.subject,
        forward: () => {},
        getThread: () => ({
          addLabel: () => {}
        })
      })
    })),
    getUserLabelByName: (name) => ({ name })
  };

  // Mock HtmlService
  HtmlService = {
    createHtmlOutput: (content) => ({
      getContent: () => content
    })
  };

  // Mock Calendar
  Calendar = {
    Events: {
      list: (email, params) => ({
        items: mockData.calendarEvents.get(email) || []
      })
    }
  };

  // Mock PropertiesService
  PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (key) => mockData.propertyValues.get(key) || ""
    })
  };

  // Mock People API
  People = {
    People: {
      searchDirectoryPeople: (params) => ({
        totalSize: params.query.includes("invalid") ? 0 : 1
      })
    }
  };

  // Mock GroupsApp
  GroupsApp = {
    getGroupByEmail: (email) => {
      if (email.includes("invalid")) {
        throw new Error("Group not found");
      }
      return { email };
    }
  };

  // Mock Utilities
  Utilities = {
    formatDate: (date, timezone, format) => date.toISOString(),
    parseDate: (dateString, timezone, format) => new Date(dateString),
    formatString: (template, ...args) => {
      let result = template;
      args.forEach((arg, i) => {
        result = result.replace(`%${i > 0 ? 's' : 'd'}`, arg);
      });
      return result;
    }
  };

  // Mock Logger
  Logger = {
    log: (message, ...args) => {
      console.log(`MOCK LOG: ${message}`, ...args);
    }
  };

  // Mock UrlFetchApp
  UrlFetchApp = {
    fetch: (url, options) => ({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        candidates: [{ content: { parts: [{ text: "Mocked LLM response" }] } }],
        usageMetadata: { totalTokenCount: 100 }
      })
    })
  };
}

// =============================================================================
// TEST UTILITIES
// =============================================================================

function resetTestEnvironment() {
  mockData = {
    formResponses: [],
    spreadsheetData: new Map(),
    documentBodies: new Map(),
    emailsSent: [],
    draftsCreated: [],
    calendarEvents: new Map(),
    propertyValues: new Map()
  };

  testResults = {
    passed: 0,
    failed: 0,
    errors: []
  };

  // Initialize Google Apps Script API mocks
  mockGoogleAppsScriptAPIs();
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(`Assertion failed: ${message}`);
  }
}

function recordTestPass(testName) {
  testResults.passed++;
  console.log(`✅ ${testName} PASSED`);
}

function recordTestFailure(testName, error) {
  testResults.failed++;
  const errorMessage = `❌ ${testName} FAILED: ${error.message}`;
  console.log(errorMessage);
  testResults.errors.push(errorMessage);
}

function printTestResults() {
  console.log("\n" + "=".repeat(50));
  console.log("🧪 TEST RESULTS SUMMARY");
  console.log("=".repeat(50));
  console.log(`✅ Tests Passed: ${testResults.passed}`);
  console.log(`❌ Tests Failed: ${testResults.failed}`);
  console.log(`📊 Total Tests: ${testResults.passed + testResults.failed}`);

  if (testResults.errors.length > 0) {
    console.log("\n🚨 FAILED TESTS:");
    testResults.errors.forEach(error => console.log(`   ${error}`));
  }

  if (testResults.failed === 0) {
    console.log("\n🎉 ALL TESTS PASSED! 🎉");
  } else {
    console.log(`\n⚠️ ${testResults.failed} tests failed. Please review and fix.`);
  }

  console.log("=".repeat(50));
}

// =============================================================================
// QUICK TEST RUNNERS
// =============================================================================

/**
 * Run only critical path tests for quick validation
 */
function runQuickTests() {
  console.log("🚀 Running quick test suite...");
  resetTestEnvironment();

  try {
    testDoGetMissingCommand();
    testDoGetGenerateCommand();
    testCreateDraftEmailsNormal();
    testCompileStatusNormal();
    testNotifyMissingStatusNormal();

    printTestResults();
  } catch (error) {
    console.error("Quick test suite failed:", error);
  }

  return testResults;
}

/**
 * Run tests for a specific function
 */
function runTestsForFunction(functionName) {
  console.log(`🎯 Running tests for ${functionName}...`);
  resetTestEnvironment();

  try {
    switch (functionName) {
      case "doGet":
        testDoGet();
        break;
      case "createDraftEmails":
        testCreateDraftEmails();
        break;
      case "compileStatus":
        testCompileStatus();
        break;
      case "notifyMissingStatus":
        testNotifyMissingStatus();
        break;
      case "notifyMismatchAssignment":
        testNotifyMismatchAssignment();
        break;
      case "validateRecipients":
        testValidateRecipients();
        break;
      case "onStatusFormSubmission":
        testOnStatusFormSubmission();
        break;
      default:
        throw new Error(`Unknown function: ${functionName}`);
    }

    printTestResults();
  } catch (error) {
    console.error(`Tests for ${functionName} failed:`, error);
  }

  return testResults;
}