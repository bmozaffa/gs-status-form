/**
 * Test Runner and Coverage Validation
 *
 * This file provides test execution utilities and validates that all
 * external functions have comprehensive test coverage.
 */

/**
 * Test Coverage Validator
 * Ensures all external functions are properly tested
 */
class TestCoverageValidator {
  constructor() {
    this.externalFunctions = [
      "doGet",
      "createDraftEmails",
      "compileStatus",
      "notifyMissingStatus",
      "notifyMismatchAssignment",
      "validateRecipients",
      "onStatusFormSubmission"
    ];

    this.testCoverage = new Map();
    this.initializeCoverage();
  }

  initializeCoverage() {
    this.externalFunctions.forEach(func => {
      this.testCoverage.set(func, {
        totalTests: 0,
        passedTests: 0,
        failedTests: 0,
        scenarios: []
      });
    });
  }

  recordTest(functionName, testName, passed, scenario = "default") {
    if (!this.testCoverage.has(functionName)) {
      console.warn(`⚠️ Unknown function in test coverage: ${functionName}`);
      return;
    }

    const coverage = this.testCoverage.get(functionName);
    coverage.totalTests++;

    if (passed) {
      coverage.passedTests++;
    } else {
      coverage.failedTests++;
    }

    if (!coverage.scenarios.includes(scenario)) {
      coverage.scenarios.push(scenario);
    }

    this.testCoverage.set(functionName, coverage);
  }

  validateCoverage() {
    console.log("\n🔍 VALIDATING TEST COVERAGE");
    console.log("=" .repeat(50));

    let overallValid = true;
    const minTestsPerFunction = 3; // Minimum tests per function
    const requiredScenarios = ["normal", "error", "edge_case"];

    this.externalFunctions.forEach(func => {
      const coverage = this.testCoverage.get(func);
      let functionValid = true;

      console.log(`\n📊 ${func}:`);
      console.log(`   Total Tests: ${coverage.totalTests}`);
      console.log(`   Passed: ${coverage.passedTests}`);
      console.log(`   Failed: ${coverage.failedTests}`);
      console.log(`   Scenarios: ${coverage.scenarios.join(", ")}`);

      // Check minimum test count
      if (coverage.totalTests < minTestsPerFunction) {
        console.log(`   ❌ Insufficient tests (minimum: ${minTestsPerFunction})`);
        functionValid = false;
      }

      // Check pass rate
      const passRate = coverage.totalTests > 0 ? coverage.passedTests / coverage.totalTests : 0;
      if (passRate < 0.8) {
        console.log(`   ❌ Low pass rate: ${(passRate * 100).toFixed(1)}%`);
        functionValid = false;
      }

      // Check scenario coverage
      const missingScenarios = requiredScenarios.filter(scenario =>
        !coverage.scenarios.includes(scenario)
      );
      if (missingScenarios.length > 0) {
        console.log(`   ⚠️ Missing scenarios: ${missingScenarios.join(", ")}`);
        // Don't fail for missing scenarios, just warn
      }

      if (functionValid) {
        console.log(`   ✅ Adequate coverage`);
      } else {
        overallValid = false;
      }
    });

    console.log("\n" + "=".repeat(50));
    if (overallValid) {
      console.log("🎉 ALL FUNCTIONS HAVE ADEQUATE TEST COVERAGE");
    } else {
      console.log("⚠️ SOME FUNCTIONS NEED MORE TEST COVERAGE");
    }

    return overallValid;
  }

  generateCoverageReport() {
    const report = {
      timestamp: new Date().toISOString(),
      summary: {
        totalFunctions: this.externalFunctions.length,
        adequateCoverage: 0,
        totalTests: 0,
        totalPassed: 0,
        totalFailed: 0
      },
      functions: {}
    };

    this.externalFunctions.forEach(func => {
      const coverage = this.testCoverage.get(func);

      report.summary.totalTests += coverage.totalTests;
      report.summary.totalPassed += coverage.passedTests;
      report.summary.totalFailed += coverage.failedTests;

      if (coverage.totalTests >= 3 && coverage.passedTests / coverage.totalTests >= 0.8) {
        report.summary.adequateCoverage++;
      }

      report.functions[func] = {
        tests: coverage.totalTests,
        passed: coverage.passedTests,
        failed: coverage.failedTests,
        passRate: coverage.totalTests > 0 ? coverage.passedTests / coverage.totalTests : 0,
        scenarios: coverage.scenarios
      };
    });

    return report;
  }
}

/**
 * Enhanced Test Runner with Coverage Tracking
 */
class EnhancedTestRunner {
  constructor() {
    this.coverageValidator = new TestCoverageValidator();
    this.testResults = {
      startTime: null,
      endTime: null,
      totalTests: 0,
      passed: 0,
      failed: 0,
      errors: [],
      performance: {}
    };
  }

  runAllTestsWithCoverage() {
    console.log("🧪 STARTING COMPREHENSIVE TEST SUITE WITH COVERAGE VALIDATION");
    console.log("=".repeat(70));

    this.testResults.startTime = new Date();

    try {
      // Reset test environment
      resetTestEnvironment();

      // Run all test suites with coverage tracking
      this.runDoGetTests();
      this.runCreateDraftEmailsTests();
      this.runCompileStatusTests();
      this.runNotifyMissingStatusTests();
      this.runNotifyMismatchAssignmentTests();
      this.runValidateRecipientsTests();
      this.runOnStatusFormSubmissionTests();

      this.testResults.endTime = new Date();

      // Generate comprehensive results
      this.printFinalResults();
      this.coverageValidator.validateCoverage();

      return {
        testResults: this.testResults,
        coverageReport: this.coverageValidator.generateCoverageReport()
      };

    } catch (error) {
      console.error("❌ Test suite execution failed:", error);
      this.testResults.errors.push(`Test suite execution error: ${error.message}`);
      return null;
    }
  }

  // Enhanced test methods with coverage tracking
  runDoGetTests() {
    console.log("\n📝 Testing doGet function with coverage tracking...");

    const testCases = [
      { name: "missing command", scenario: "normal", test: this.testDoGetMissingCommand },
      { name: "generate command", scenario: "normal", test: this.testDoGetGenerateCommand },
      { name: "archive command", scenario: "normal", test: this.testDoGetArchiveCommand },
      { name: "send-drafts command", scenario: "normal", test: this.testDoGetSendDraftsCommand },
      { name: "mismatch command", scenario: "normal", test: this.testDoGetMismatchCommand },
      { name: "rewrite command", scenario: "normal", test: this.testDoGetRewriteCommand },
      { name: "invalid command", scenario: "error", test: this.testDoGetInvalidCommand },
      { name: "no parameters", scenario: "error", test: this.testDoGetNoParameters },
      { name: "malformed parameters", scenario: "edge_case", test: this.testDoGetMalformedParams }
    ];

    this.runTestCases("doGet", testCases);
  }

  runCreateDraftEmailsTests() {
    console.log("\n📧 Testing createDraftEmails function with coverage tracking...");

    const testCases = [
      { name: "normal operation", scenario: "normal", test: this.testCreateDraftEmailsNormal },
      { name: "paused operation", scenario: "normal", test: this.testCreateDraftEmailsPaused },
      { name: "empty data", scenario: "edge_case", test: this.testCreateDraftEmailsEmpty },
      { name: "invalid email format", scenario: "error", test: this.testCreateDraftEmailsInvalidFormat },
      { name: "missing spreadsheet", scenario: "error", test: this.testCreateDraftEmailsMissingSheet }
    ];

    this.runTestCases("createDraftEmails", testCases);
  }

  runCompileStatusTests() {
    console.log("\n📊 Testing compileStatus function with coverage tracking...");

    const testCases = [
      { name: "normal compilation", scenario: "normal", test: this.testCompileStatusNormal },
      { name: "paused operation", scenario: "normal", test: this.testCompileStatusPaused },
      { name: "no updates needed", scenario: "edge_case", test: this.testCompileStatusNoUpdate },
      { name: "invalid document ID", scenario: "error", test: this.testCompileStatusInvalidDocId },
      { name: "large dataset", scenario: "edge_case", test: this.testCompileStatusLargeDataset }
    ];

    this.runTestCases("compileStatus", testCases);
  }

  runNotifyMissingStatusTests() {
    console.log("\n🔔 Testing notifyMissingStatus function with coverage tracking...");

    const testCases = [
      { name: "normal notification", scenario: "normal", test: this.testNotifyMissingStatusNormal },
      { name: "paused operation", scenario: "normal", test: this.testNotifyMissingStatusPaused },
      { name: "no missing status", scenario: "edge_case", test: this.testNotifyMissingStatusNone },
      { name: "all users on PTO", scenario: "edge_case", test: this.testNotifyMissingStatusAllPTO },
      { name: "email send failure", scenario: "error", test: this.testNotifyMissingStatusEmailFailure }
    ];

    this.runTestCases("notifyMissingStatus", testCases);
  }

  runNotifyMismatchAssignmentTests() {
    console.log("\n⚠️ Testing notifyMismatchAssignment function with coverage tracking...");

    const testCases = [
      { name: "normal notification", scenario: "normal", test: this.testNotifyMismatchAssignmentNormal },
      { name: "paused operation", scenario: "normal", test: this.testNotifyMismatchAssignmentPaused },
      { name: "no mismatches", scenario: "edge_case", test: this.testNotifyMismatchAssignmentNone },
      { name: "multiple managers", scenario: "normal", test: this.testNotifyMismatchAssignmentMultipleMgrs },
      { name: "assignment error", scenario: "error", test: this.testNotifyMismatchAssignmentError }
    ];

    this.runTestCases("notifyMismatchAssignment", testCases);
  }

  runValidateRecipientsTests() {
    console.log("\n✉️ Testing validateRecipients function with coverage tracking...");

    const testCases = [
      { name: "valid recipients", scenario: "normal", test: this.testValidateRecipientsNormal },
      { name: "paused operation", scenario: "normal", test: this.testValidateRecipientsPaused },
      { name: "invalid recipients", scenario: "normal", test: this.testValidateRecipientsInvalid },
      { name: "mixed valid/invalid", scenario: "edge_case", test: this.testValidateRecipientsMixed },
      { name: "API unavailable", scenario: "error", test: this.testValidateRecipientsAPIFailure }
    ];

    this.runTestCases("validateRecipients", testCases);
  }

  runOnStatusFormSubmissionTests() {
    console.log("\n🤖 Testing onStatusFormSubmission function with coverage tracking...");

    const testCases = [
      { name: "normal submission", scenario: "normal", test: this.testOnStatusFormSubmissionNormal },
      { name: "invalid row", scenario: "error", test: this.testOnStatusFormSubmissionInvalidRow },
      { name: "LLM API failure", scenario: "error", test: this.testOnStatusFormSubmissionLLMFailure },
      { name: "missing data", scenario: "edge_case", test: this.testOnStatusFormSubmissionMissingData },
      { name: "concurrent submissions", scenario: "edge_case", test: this.testOnStatusFormSubmissionConcurrent }
    ];

    this.runTestCases("onStatusFormSubmission", testCases);
  }

  runTestCases(functionName, testCases) {
    testCases.forEach(testCase => {
      try {
        setupMockData(); // Reset for each test
        const startTime = Date.now();

        testCase.test.call(this);

        const duration = Date.now() - startTime;
        this.recordTestResult(functionName, testCase.name, true, testCase.scenario, duration);

      } catch (error) {
        this.recordTestResult(functionName, testCase.name, false, testCase.scenario, 0, error);
      }
    });
  }

  recordTestResult(functionName, testName, passed, scenario, duration, error = null) {
    this.testResults.totalTests++;

    if (passed) {
      this.testResults.passed++;
      console.log(`✅ ${functionName} - ${testName} PASSED (${duration}ms)`);
    } else {
      this.testResults.failed++;
      const errorMessage = `❌ ${functionName} - ${testName} FAILED: ${error?.message || 'Unknown error'}`;
      console.log(errorMessage);
      this.testResults.errors.push(errorMessage);
    }

    // Record in coverage validator
    this.coverageValidator.recordTest(functionName, testName, passed, scenario);

    // Track performance
    if (!this.testResults.performance[functionName]) {
      this.testResults.performance[functionName] = [];
    }
    this.testResults.performance[functionName].push({
      testName,
      duration,
      passed
    });
  }

  // Additional test methods for comprehensive coverage
  testDoGetNoParameters() {
    const mockEvent = { parameter: {} };
    const result = doGet(mockEvent);
    const content = result.getContent();
    assert(content.includes("Provide the command"), "Should show help for missing command");
  }

  testDoGetMalformedParams() {
    const mockEvent = { parameter: { command: "generate", docId: null } };
    // Should handle gracefully
    doGet(mockEvent);
  }

  testCreateDraftEmailsInvalidFormat() {
    setupMockEmailData();
    mockData.spreadsheetData.set("emails", [
      ["invalid-email", "", "", "", "", ""] // Invalid email format
    ]);

    // Should handle invalid email gracefully
    createDraftEmails();
  }

  testCreateDraftEmailsMissingSheet() {
    // Don't setup email data - simulate missing sheet
    try {
      createDraftEmails();
    } catch (error) {
      // Expected to throw error for missing sheet
      throw error;
    }
  }

  testCompileStatusInvalidDocId() {
    const result = compileStatusDocs("invalid_doc_id");
    // Should handle invalid doc ID gracefully
  }

  testCompileStatusLargeDataset() {
    // Setup large dataset
    const generator = new PerformanceTestDataGenerator();
    mockData.formResponses = generator.generateManyFormResponses(50);

    compileStatusDocs();
  }

  testNotifyMissingStatusAllPTO() {
    setupMockMissingStatus();
    // Set all users as on PTO
    mockData.calendarEvents.set(TEST_CONFIG.TEST_EMAIL, [
      { eventType: "outOfOffice", start: { date: "2024-01-15" } }
    ]);

    notifyMissingStatus();
  }

  testNotifyMissingStatusEmailFailure() {
    setupMockMissingNotification();
    // Mock email sending failure
    const originalGmailApp = GmailApp;
    GmailApp = {
      sendEmail: () => { throw new Error("Email send failed"); }
    };

    try {
      notifyMissingStatus();
    } finally {
      GmailApp = originalGmailApp;
    }
  }

  testNotifyMismatchAssignmentMultipleMgrs() {
    setupMockMismatch();
    // Add multiple managers
    kerberosMap.set("mgr2", new Map([
      ["Name", "Manager Two"],
      ["Email", "mgr2@redhat.com"],
      ["Job Title", "Manager, Software Engineering"]
    ]));

    notifyMismatchAssignment();
  }

  testNotifyMismatchAssignmentError() {
    // Simulate error in assignment comparison
    const originalCompareAssignments = compareAssignments;
    compareAssignments = () => { throw new Error("Assignment comparison failed"); };

    try {
      notifyMismatchAssignment();
    } finally {
      compareAssignments = originalCompareAssignments;
    }
  }

  testValidateRecipientsMixed() {
    mockData.spreadsheetData.set("emails", [
      ["valid@redhat.com", "", "", "Valid", "label1", ""],
      ["invalid@nowhere.com", "", "", "Invalid", "label2", ""],
      ["another-valid@redhat.com", "", "", "Valid2", "label3", ""]
    ]);

    validateRecipients();
  }

  testValidateRecipientsAPIFailure() {
    setupMockValidRecipients();
    // Mock API failure
    const originalPeople = People;
    People = {
      People: {
        searchDirectoryPeople: () => { throw new Error("API unavailable"); }
      }
    };

    try {
      validateRecipients();
    } finally {
      People = originalPeople;
    }
  }

  testOnStatusFormSubmissionLLMFailure() {
    const mockEvent = { range: { rowStart: 100 } };
    setupMockLLMProcessing();

    // Mock LLM API failure
    const originalUrlFetchApp = UrlFetchApp;
    UrlFetchApp = {
      fetch: () => { throw new Error("LLM API unavailable"); }
    };

    try {
      onStatusFormSubmission(mockEvent);
    } finally {
      UrlFetchApp = originalUrlFetchApp;
    }
  }

  testOnStatusFormSubmissionMissingData() {
    const mockEvent = { range: { rowStart: 999 } }; // Row with no data
    onStatusFormSubmission(mockEvent);
  }

  testOnStatusFormSubmissionConcurrent() {
    const mockEvent1 = { range: { rowStart: 100 } };
    const mockEvent2 = { range: { rowStart: 101 } };

    setupMockLLMProcessing();

    // Simulate concurrent submissions
    onStatusFormSubmission(mockEvent1);
    onStatusFormSubmission(mockEvent2);
  }

  printFinalResults() {
    const duration = this.testResults.endTime - this.testResults.startTime;

    console.log("\n" + "=".repeat(70));
    console.log("🧪 COMPREHENSIVE TEST RESULTS SUMMARY");
    console.log("=".repeat(70));
    console.log(`⏱️ Execution Time: ${duration}ms`);
    console.log(`✅ Tests Passed: ${this.testResults.passed}`);
    console.log(`❌ Tests Failed: ${this.testResults.failed}`);
    console.log(`📊 Total Tests: ${this.testResults.totalTests}`);
    console.log(`📈 Pass Rate: ${(this.testResults.passed / this.testResults.totalTests * 100).toFixed(1)}%`);

    if (this.testResults.errors.length > 0) {
      console.log("\n🚨 FAILED TESTS:");
      this.testResults.errors.forEach(error => console.log(`   ${error}`));
    }

    // Performance summary
    console.log("\n⚡ PERFORMANCE SUMMARY:");
    Object.entries(this.testResults.performance).forEach(([funcName, tests]) => {
      const avgDuration = tests.reduce((sum, t) => sum + t.duration, 0) / tests.length;
      const passRate = tests.filter(t => t.passed).length / tests.length * 100;
      console.log(`   ${funcName}: ${avgDuration.toFixed(1)}ms avg, ${passRate.toFixed(1)}% pass rate`);
    });

    console.log("=".repeat(70));
  }
}

/**
 * Convenience functions for different test execution modes
 */

/**
 * Run the complete test suite with coverage validation
 */
function runComprehensiveTests() {
  const runner = new EnhancedTestRunner();
  return runner.runAllTestsWithCoverage();
}

/**
 * Run performance tests
 */
function runPerformanceTests() {
  console.log("🚀 RUNNING PERFORMANCE TESTS");

  const generator = new PerformanceTestDataGenerator();
  const startTime = Date.now();

  // Test with large datasets
  resetTestEnvironment();

  // Generate large user set
  const users = generator.generateLargeUserSet(500);
  users.forEach((userData, kerberos) => {
    kerberosMap.set(kerberos, userData);
  });

  // Generate many responses
  mockData.formResponses = generator.generateManyFormResponses(1000);

  // Test compileStatus performance
  const compileStart = Date.now();
  try {
    compileStatus();
    console.log(`✅ compileStatus with 1000 responses: ${Date.now() - compileStart}ms`);
  } catch (error) {
    console.log(`❌ compileStatus performance test failed: ${error.message}`);
  }

  // Test missing status performance
  const missingStart = Date.now();
  try {
    const missing = getMissingStatus(mockData.formResponses, ["test_doc"]);
    console.log(`✅ getMissingStatus with 500 users: ${Date.now() - missingStart}ms`);
  } catch (error) {
    console.log(`❌ getMissingStatus performance test failed: ${error.message}`);
  }

  const totalTime = Date.now() - startTime;
  console.log(`🏁 Total performance test time: ${totalTime}ms`);

  return {
    totalTime,
    userCount: 500,
    responseCount: 1000
  };
}

/**
 * Validate that all external functions have tests
 */
function validateTestCompleteness() {
  const validator = new TestCoverageValidator();
  return validator.validateCoverage();
}