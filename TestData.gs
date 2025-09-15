/**
 * Test Data Generator and Mock Utilities
 *
 * This file provides comprehensive test data generators and utilities
 * for creating realistic test scenarios for the Red Hat Status Reporting System.
 */

/**
 * Generator for realistic test data
 */
class TestDataGenerator {
  constructor() {
    this.currentTimestamp = new Date("2024-01-15T10:00:00Z");
    this.userCounter = 1;
    this.responseCounter = 1;
  }

  /**
   * Generate a realistic form response object
   */
  generateFormResponse(overrides = {}) {
    const baseResponse = {
      id: `response_${this.responseCounter++}`,
      timestamp: new Date(this.currentTimestamp.getTime() + (this.responseCounter * 60000)),
      kerberos: `testuser${this.userCounter}`,
      email: `testuser${this.userCounter}@redhat.com`,
      initiative: "TestInitiative",
      effort: "Development",
      epic: "APPENG-1234: Implement new feature",
      status: "Completed initial implementation and added unit tests. Ready for code review."
    };

    this.userCounter++;
    return { ...baseResponse, ...overrides };
  }

  /**
   * Generate multiple form responses
   */
  generateFormResponses(count, overrides = []) {
    const responses = [];
    for (let i = 0; i < count; i++) {
      const override = overrides[i] || {};
      responses.push(this.generateFormResponse(override));
    }
    return responses;
  }

  /**
   * Generate user data for Kerberos map
   */
  generateUserData(kerberos, overrides = {}) {
    const baseData = new Map([
      ["Name", `Test User ${kerberos}`],
      ["Email", `${kerberos}@redhat.com`],
      ["Manager", "testmgr"],
      ["Job Title", "Software Engineer"],
      ["Termination", ""]
    ]);

    Object.entries(overrides).forEach(([key, value]) => {
      baseData.set(key, value);
    });

    return baseData;
  }

  /**
   * Generate manager data
   */
  generateManagerData(kerberos, overrides = {}) {
    return this.generateUserData(kerberos, {
      "Job Title": "Manager, Software Engineering",
      "Manager": "topmgr",
      ...overrides
    });
  }

  /**
   * Generate director data
   */
  generateDirectorData(kerberos, overrides = {}) {
    return this.generateUserData(kerberos, {
      "Job Title": "Director, Software Engineering_Global",
      "Manager": "vp",
      ...overrides
    });
  }

  /**
   * Generate assignment data for roster spreadsheet
   */
  generateAssignmentData(kerberos, assignments) {
    return assignments.map(assignment => [assignment, "", "", "", kerberos]);
  }

  /**
   * Generate email data for draft emails
   */
  generateEmailData(count = 3) {
    const emails = [];
    for (let i = 1; i <= count; i++) {
      emails.push([
        `recipient${i}@redhat.com`,
        `cc${i}@redhat.com`,
        `bcc${i}@redhat.com`,
        `Test Subject ${i}`,
        `label${i},automated`,
        ""
      ]);
    }
    return emails;
  }

  /**
   * Generate calendar events for PTO testing
   */
  generateCalendarEvents(email, isPTO = false) {
    if (!isPTO) return [];

    return [
      {
        eventType: "outOfOffice",
        start: { date: "2024-01-15" },
        end: { date: "2024-01-16" },
        summary: "Personal Time Off"
      }
    ];
  }

  /**
   * Generate LLM edit data
   */
  generateLLMEditData(responseObject) {
    return [
      100, // row number
      responseObject.email,
      responseObject.timestamp,
      responseObject.status,
      "• Implemented new feature with comprehensive unit tests\n• Prepared code for review process"
    ];
  }

  /**
   * Generate document template structure
   */
  generateDocumentTemplate() {
    return {
      paragraphs: [
        "This document is no longer auto-generated. It was last generated at " + new Date().toUTCString(),
        "Click here if you think more status entries are available and wish to refresh this document."
      ],
      listItems: [
        "%TestInitiative.Development%",
        "%TestInitiative.Testing%",
        "%TestInitiative.Other%",
        "%Missing Status%",
        "%PTO / No Status%"
      ]
    };
  }
}

/**
 * Test Scenario Builder
 *
 * Builds complete test scenarios with all necessary mock data
 */
class TestScenarioBuilder {
  constructor() {
    this.generator = new TestDataGenerator();
    this.scenario = {
      users: new Map(),
      responses: [],
      assignments: new Map(),
      emails: [],
      calendarEvents: new Map(),
      properties: new Map(),
      documents: new Map()
    };
  }

  /**
   * Add a user to the scenario
   */
  addUser(kerberos, userData = {}) {
    this.scenario.users.set(kerberos, this.generator.generateUserData(kerberos, userData));
    return this;
  }

  /**
   * Add a manager to the scenario
   */
  addManager(kerberos, userData = {}) {
    this.scenario.users.set(kerberos, this.generator.generateManagerData(kerberos, userData));
    return this;
  }

  /**
   * Add a director to the scenario
   */
  addDirector(kerberos, userData = {}) {
    this.scenario.users.set(kerberos, this.generator.generateDirectorData(kerberos, userData));
    return this;
  }

  /**
   * Add form responses to the scenario
   */
  addFormResponses(responses) {
    if (Array.isArray(responses)) {
      this.scenario.responses.push(...responses);
    } else {
      this.scenario.responses.push(responses);
    }
    return this;
  }

  /**
   * Add assignments to the scenario
   */
  addAssignments(kerberos, assignments) {
    this.scenario.assignments.set(kerberos, assignments);
    return this;
  }

  /**
   * Add email data to the scenario
   */
  addEmailData(emailData) {
    this.scenario.emails.push(...emailData);
    return this;
  }

  /**
   * Add calendar events for a user
   */
  addCalendarEvents(email, events) {
    this.scenario.calendarEvents.set(email, events);
    return this;
  }

  /**
   * Set property values
   */
  setProperty(key, value) {
    this.scenario.properties.set(key, value);
    return this;
  }

  /**
   * Add document structure
   */
  addDocument(docId, structure) {
    this.scenario.documents.set(docId, structure);
    return this;
  }

  /**
   * Build and return the complete scenario
   */
  build() {
    return this.scenario;
  }

  /**
   * Apply the scenario to mock data
   */
  applyToMocks(mockData) {
    // Apply users to kerberosMap
    this.scenario.users.forEach((userData, kerberos) => {
      if (typeof kerberosMap !== 'undefined') {
        kerberosMap.set(kerberos, userData);
      }
    });

    // Apply responses to form responses
    mockData.formResponses = [...this.scenario.responses];

    // Apply assignments to spreadsheet data
    const rosterData = [];
    this.scenario.assignments.forEach((assignments, kerberos) => {
      assignments.forEach(assignment => {
        rosterData.push([assignment, "", "", "", kerberos]);
      });
    });
    mockData.spreadsheetData.set("Roster by Person", rosterData);

    // Apply email data
    mockData.spreadsheetData.set("emails", this.scenario.emails);

    // Apply calendar events
    this.scenario.calendarEvents.forEach((events, email) => {
      mockData.calendarEvents.set(email, events);
    });

    // Apply properties
    this.scenario.properties.forEach((value, key) => {
      mockData.propertyValues.set(key, value);
    });

    return mockData;
  }
}

/**
 * Pre-built test scenarios for common testing situations
 */
class CommonTestScenarios {
  static typicalWorkWeek() {
    return new TestScenarioBuilder()
      .addManager("mgr1", { "Name": "Test Manager" })
      .addUser("dev1", { "Manager": "mgr1", "Name": "Developer One" })
      .addUser("dev2", { "Manager": "mgr1", "Name": "Developer Two" })
      .addUser("dev3", { "Manager": "mgr1", "Name": "Developer Three" })
      .addFormResponses([
        {
          kerberos: "dev1",
          email: "dev1@redhat.com",
          initiative: "Project Alpha",
          effort: "Backend Development",
          epic: "ALPHA-123: User Authentication",
          status: "Implemented OAuth2 integration and added security tests"
        },
        {
          kerberos: "dev2",
          email: "dev2@redhat.com",
          initiative: "Project Alpha",
          effort: "Frontend Development",
          epic: "ALPHA-456: User Interface",
          status: "Created responsive login form with validation"
        }
      ])
      .addAssignments("dev1", ["Project Alpha.Backend Development"])
      .addAssignments("dev2", ["Project Alpha.Frontend Development"])
      .addAssignments("dev3", ["Project Alpha.Testing"]) // No status submitted
      .setProperty("compileStatus", false)
      .setProperty("notifyMissingStatus", false);
  }

  static missingStatusScenario() {
    return new TestScenarioBuilder()
      .addManager("mgr1")
      .addUser("dev1", { "Manager": "mgr1" })
      .addUser("dev2", { "Manager": "mgr1" })
      .addUser("dev3", { "Manager": "mgr1" })
      .addFormResponses([
        {
          kerberos: "dev1",
          initiative: "Project Alpha",
          status: "Working on feature implementation"
        }
      ])
      .addAssignments("dev1", ["Project Alpha"])
      .addAssignments("dev2", ["Project Alpha"]) // Missing status
      .addAssignments("dev3", ["Project Alpha"]) // Missing status
      .setProperty("notifyMissingStatus", false);
  }

  static assignmentMismatchScenario() {
    return new TestScenarioBuilder()
      .addManager("mgr1")
      .addUser("dev1", { "Manager": "mgr1" })
      .addUser("dev2", { "Manager": "mgr1" })
      .addFormResponses([
        {
          kerberos: "dev1",
          initiative: "Project Beta", // Not assigned to this
          status: "Working on unassigned project"
        },
        {
          kerberos: "dev2",
          initiative: "PTO / No Status",
          status: ""
        }
      ])
      .addAssignments("dev1", ["Project Alpha"]) // Assigned to different project
      .addAssignments("dev2", ["Project Alpha"]) // On PTO
      .setProperty("notifyMismatchAssignment", false);
  }

  static ptoScenario() {
    return new TestScenarioBuilder()
      .addManager("mgr1")
      .addUser("dev1", { "Manager": "mgr1" })
      .addUser("dev2", { "Manager": "mgr1" })
      .addCalendarEvents("dev1@redhat.com", [
        {
          eventType: "outOfOffice",
          start: { date: "2024-01-15" },
          end: { date: "2024-01-16" }
        }
      ])
      .addFormResponses([
        {
          kerberos: "dev2",
          initiative: "Project Alpha",
          status: "Continuing development work"
        }
      ])
      .addAssignments("dev1", ["Project Alpha"]) // On PTO, should be excluded
      .addAssignments("dev2", ["Project Alpha"]);
  }

  static emailValidationScenario() {
    return new TestScenarioBuilder()
      .addEmailData([
        ["valid@redhat.com", "cc@redhat.com", "", "Valid Email", "label1", ""],
        ["invalid@nonexistent.com", "", "", "Invalid Email", "label2", ""],
        ["group-list@redhat.com", "", "", "Group Email", "label3", ""]
      ])
      .setProperty("validateRecipient", false);
  }

  static llmProcessingScenario() {
    const generator = new TestDataGenerator();
    const response = generator.generateFormResponse({
      epic: "PROJ-789: Database Optimization",
      status: "I optimized the database queries and they are now 50% faster. Also fixed a bug in the connection pooling."
    });

    return new TestScenarioBuilder()
      .addFormResponses([response])
      .setProperty("GEMINI_API_KEY", "mock_api_key")
      .setProperty("AI_SERVICE", "GeminiAPI");
  }

  static pausedOperationsScenario() {
    return new TestScenarioBuilder()
      .addManager("mgr1")
      .addUser("dev1", { "Manager": "mgr1" })
      .setProperty("compileStatus", true) // Paused
      .setProperty("notifyMissingStatus", true) // Paused
      .setProperty("notifyMismatchAssignment", true) // Paused
      .setProperty("createDraftEmails", true) // Paused
      .setProperty("validateRecipient", true); // Paused
  }
}

/**
 * Test Data Validation Utilities
 */
class TestDataValidator {
  static validateFormResponse(response) {
    const required = ["kerberos", "timestamp", "initiative"];
    const missing = required.filter(field => !response[field]);

    if (missing.length > 0) {
      throw new Error(`Missing required fields: ${missing.join(", ")}`);
    }

    if (response.email && !response.email.includes("@")) {
      throw new Error("Invalid email format");
    }

    if (response.timestamp && !(response.timestamp instanceof Date)) {
      throw new Error("Timestamp must be a Date object");
    }
  }

  static validateUserData(userData) {
    const required = ["Name", "Email", "Manager", "Job Title"];
    const missing = required.filter(field => !userData.has(field));

    if (missing.length > 0) {
      throw new Error(`Missing required user fields: ${missing.join(", ")}`);
    }

    if (!userData.get("Email").includes("@")) {
      throw new Error("Invalid email format for user");
    }
  }

  static validateEmailData(emailRow) {
    if (!Array.isArray(emailRow) || emailRow.length < 4) {
      throw new Error("Email data must be array with at least 4 elements");
    }

    const [to, cc, bcc, subject] = emailRow;
    if (!subject || subject.trim().length === 0) {
      throw new Error("Email subject cannot be empty");
    }
  }

  static validateTestScenario(scenario) {
    // Validate users
    scenario.users.forEach((userData, kerberos) => {
      this.validateUserData(userData);
    });

    // Validate responses
    scenario.responses.forEach(response => {
      this.validateFormResponse(response);
    });

    // Validate emails
    scenario.emails.forEach(emailRow => {
      this.validateEmailData(emailRow);
    });

    console.log("✅ Test scenario validation passed");
  }
}

/**
 * Performance Test Data Generator
 *
 * Generates large datasets for performance testing
 */
class PerformanceTestDataGenerator extends TestDataGenerator {
  generateLargeUserSet(count = 100) {
    const users = new Map();
    const managers = ["mgr1", "mgr2", "mgr3", "mgr4", "mgr5"];

    for (let i = 1; i <= count; i++) {
      const kerberos = `user${i.toString().padStart(3, "0")}`;
      const manager = managers[i % managers.length];

      users.set(kerberos, this.generateUserData(kerberos, {
        "Manager": manager,
        "Name": `User ${i}`,
        "Job Title": i % 10 === 0 ? "Senior Software Engineer" : "Software Engineer"
      }));
    }

    return users;
  }

  generateManyFormResponses(count = 200) {
    const responses = [];
    const initiatives = ["ProjectAlpha", "ProjectBeta", "ProjectGamma", "ProjectDelta"];
    const efforts = ["Development", "Testing", "Documentation", "Research"];

    for (let i = 1; i <= count; i++) {
      const kerberos = `user${(i % 100 + 1).toString().padStart(3, "0")}`;
      const initiative = initiatives[i % initiatives.length];
      const effort = efforts[i % efforts.length];

      responses.push(this.generateFormResponse({
        kerberos,
        initiative,
        effort,
        epic: `${initiative.toUpperCase()}-${100 + i}: Feature ${i}`,
        status: `Completed milestone ${i % 5 + 1} for this feature. Progress is on track.`
      }));
    }

    return responses;
  }

  generateComplexAssignmentMatrix(userCount = 100) {
    const assignments = new Map();
    const projects = ["ProjectAlpha", "ProjectBeta", "ProjectGamma", "ProjectDelta"];
    const subProjects = ["Frontend", "Backend", "Database", "API", "Testing"];

    for (let i = 1; i <= userCount; i++) {
      const kerberos = `user${i.toString().padStart(3, "0")}`;
      const userAssignments = [];

      // Each user gets 1-3 assignments
      const assignmentCount = Math.floor(Math.random() * 3) + 1;

      for (let j = 0; j < assignmentCount; j++) {
        const project = projects[Math.floor(Math.random() * projects.length)];
        const subProject = subProjects[Math.floor(Math.random() * subProjects.length)];
        userAssignments.push(`${project}.${subProject}`);
      }

      assignments.set(kerberos, [...new Set(userAssignments)]); // Remove duplicates
    }

    return assignments;
  }
}

/**
 * Export test utilities for use in main test suite
 */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    TestDataGenerator,
    TestScenarioBuilder,
    CommonTestScenarios,
    TestDataValidator,
    PerformanceTestDataGenerator
  };
}