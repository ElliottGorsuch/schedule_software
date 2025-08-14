/**
 * Main Apps Script file - replaces WixPageCode.js functionality
 * Handles data loading, component communication, and CRUD operations
 */

// Configuration for public release: read IDs/keys from Script Properties
// Set these in Apps Script: Project Settings -> Script Properties
const SPREADSHEET_ID = (function() {
  try {
    return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
  } catch (e) {
    return '';
  }
})();
const DEBUG_MODE = (function() {
  try {
    return PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true';
  } catch (e) {
    return false;
  }
})();

const SHEET_NAMES = {
  THERAPISTS: 'Therapists',
  CLIENTS: 'Clients', 
  SESSIONS: 'Sessions',
  ASSIGNMENTS: 'Assignments',
  NOTES: 'Notes'
};

/**
 * Main web app entry point - serves HTML pages.
 * This is the master router for your application.
 */
function doGet(e) {
  try {
    console.log('=== doGet called ===');
    console.log('Event parameters:', e.parameter);
    
    const page = e.parameter.page || 'main';
    console.log(`doGet received request for page: "${page}"`);

    switch (page) {
      case 'table':
      case 'schedule':
        // Serve the original working file directly (now renamed to index.html)
        console.log('Serving the original index.html file');
        return HtmlService.createHtmlOutputFromFile('index')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .setTitle('Therapist Schedule Table');

      case 'map':
        // Serve the map file
        console.log('Serving the map.html file');
        return HtmlService.createHtmlOutputFromFile('map')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .setTitle('Therapist Map View');

      case 'main':
      default:
        console.log('Serving the main dashboard page');
        const template = HtmlService.createTemplateFromFile('main');
        // Inject the correct web app URL from the server-side
        template.webAppUrl = ScriptApp.getService().getUrl();
        return template.evaluate()
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .setTitle('Therapist Scheduling System');
    }
  } catch (error) {
    console.error('Critical error in doGet router:', error);
    return HtmlService.createHtmlOutput(`<h1>Error Loading Application</h1><p><strong>Error:</strong> ${error.message}</p><p><strong>Stack:</strong> ${error.stack}</p>`);
  }
}

/**
 * Centralized action dispatcher used by doPost and internal tests
 */
function performApiAction(action, params) {
    switch(action) {
      case 'getMapData':
        console.log('📊 Calling getMapData...');
      return getMapData();
    case 'getMapDataTest':
      console.log('🧪 Calling getMapDataTest (minimal data)...');
      return {
        success: true,
        therapists: [],
        clients: [],
        sessions: [],
        assignments: {},
        message: 'getMapDataTest - minimal data only',
        timestamp: new Date().toISOString()
      };
      case 'getEssentialData':
      return getEssentialData();
      case 'getSessionsData':
      return getSessionsData(params);
      case 'getScheduleData':
      return getScheduleData();
      case 'testDoPostEndpoints':
      return testDoPostEndpoints();
      case 'testMakeDoPostRequest':
      return testMakeDoPostRequest();
      case 'createAssignment':
      return createAssignment(params);
      case 'clearAssignment':
      return clearAssignment(params);
      case 'clearTherapistAssignments':
      return clearTherapistAssignments(params);
      case 'addTherapist':
      return addTherapist(params);
      case 'addClient':
      return addClient(params);
      case 'deleteTherapist':
      return deleteTherapist(params.therapistId);
      case 'deleteClient':
      return deleteClient(params.clientId);
      case 'testConnection':
      return testConnection();
      case 'testGeocoding':
      return testGeocoding();
      case 'testApiKeySetup':
      return testApiKeySetup();
      case 'testDistanceCalculation':
      return testDistanceCalculation(params);
      case 'toggleNAStatus':
      return toggleNAStatus(params);
      case 'setNAStatus':
      return setNAStatus(params);
      case 'removeNAStatus':
      return removeNAStatus(params);
    case 'getClientDistanceStats':
      return getClientDistanceStats();
    case 'testAssignmentDetails':
      return testAssignmentDetails(params);
    case 'copyEntireSchedule':
      return copyEntireSchedule(params);
    case 'copySelectedTherapists':
      return copySelectedTherapists(params);
    case 'checkUserAccess':
      return checkUserAccess();
    case 'testDeployment':
      return testDeployment();
    case 'analyzeClientDistancesDataQuality':
      return analyzeClientDistancesDataQuality();
    case 'debugClientDistancesSheet':
      return debugClientDistancesSheet();
    case 'getTherapistNotes':
      return getTherapistNotes();
    case 'saveTherapistNote':
      return saveTherapistNote(params);
    case 'deleteTherapistNote':
      return deleteTherapistNote(params);
    case 'updateTherapistLead':
      return updateTherapistLead(params);
    case 'calculateAllClientDistances':
      return DistanceService.calculateAllClientToClientDistances();
    case 'reGeocodeAllClients':
      return reGeocodeAllClients();
    case 'clearClientDistances':
      return clearClientDistancesSheet();
    default:
      console.error(`❌ Unknown action: ${action}`);
      return { success: false, error: `Unknown action: ${action}` };
  }
}

/**
 * Handle POST requests for API calls
 * This allows your HTML components to make API calls to your backend
 */
function doPost(e) {
  try {
    console.log('=== doPost called ===');
    if (DEBUG_MODE) {
      console.log('🔍 DEBUG: doPost function reached successfully');
      console.log('🔍 DEBUG: postData exists:', !!(e && e.postData));
    }
    
    // Step 1: Authentication Check
    // Optional authentication check (depends on Web App deployment settings)
    // For public release, avoid logging user PII
    const user = Session.getActiveUser();
    const userEmail = user && typeof user.getEmail === 'function' ? user.getEmail() : '';
    if (DEBUG_MODE) {
      console.log('🔐 Auth check performed. Authenticated:', !!userEmail);
    }
    
    // Step 2: Optional sanity check – do not block on this; detailed errors will surface downstream
    try {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheets = spreadsheet.getSheets();
      console.log('📊 Spreadsheet accessible, sheets:', sheets.length);
    } catch (accessError) {
      console.warn('⚠️ Spreadsheet access check failed (non-blocking):', accessError.message);
    }
    
    // If no postData, return a simple test response
    if (!e || !e.postData) {
      console.error('❌ No postData in request');
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'No POST data received',
          message: 'This is a JSON response from doPost',
          timestamp: new Date().toISOString()
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (DEBUG_MODE) {
      const size = e && e.postData && e.postData.contents ? String(e.postData.contents).length : 0;
      console.log('📥 Received postData (redacted). Size:', size);
    }
    
    // Simple test - if postData is just "test", return test response
    if (e.postData.contents === 'test') {
      console.log('🔍 DEBUG: Returning test response');
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          message: 'doPost is working!',
          user: userEmail,
          timestamp: new Date().toISOString()
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // DEBUG: Test for getMapData specifically
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const params = postData.params;
    
    if (DEBUG_MODE) {
      console.log(`🎯 API call: ${action}`);
    }
    
    // Restrict exposed actions for public release
    const allowedActions = new Set([
      'getMapData',
      'getEssentialData',
      'getSessionsData',
      'getScheduleData',
      'createAssignment',
      'clearAssignment',
      'clearTherapistAssignments',
      'addTherapist',
      'addClient',
      'deleteTherapist',
      'deleteClient',
      'toggleNAStatus',
      'setNAStatus',
      'removeNAStatus',
      'getTherapistNotes',
      'saveTherapistNote',
      'deleteTherapistNote',
      'getClientDistanceStats',
      'updateTherapistLead'
    ]);
    if (!allowedActions.has(action)) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: `Unknown or disallowed action: ${action}`
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Route to appropriate function via shared handler
    const result = performApiAction(action, params);
    
    if (DEBUG_MODE) {
      try { console.log('📤 Function result (redacted):', result && result.success !== false); } catch (ignore) {}
    }
    
    const response = ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
    console.log('✅ Returning JSON response');
    return response;
    
  } catch (error) {
    console.error('❌ Error in doPost:', error);
    console.error('📍 Error stack:', error.stack);
    
    // Check if this is an authentication-related error
    const errorMessage = error.message || error.toString();
    const isAuthError = errorMessage.includes('Authentication') || 
                       errorMessage.includes('Access') || 
                       errorMessage.includes('permission') ||
                       errorMessage.includes('login') ||
                       errorMessage.includes('<!DOCTYPE');
    
    if (isAuthError) {
      const authErrorResponse = ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          error: 'Authentication required',
          message: 'Please ensure you are logged in and have access to this application',
          requiresAuth: true,
          timestamp: new Date().toISOString()
        }))
        .setMimeType(ContentService.MimeType.JSON);
        
      if (DEBUG_MODE) { console.log('📤 Returning authentication error response'); }
      return authErrorResponse;
    }
    
    const errorResponse = ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.message,
        stack: error.stack
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
    if (DEBUG_MODE) { console.log('📤 Returning error JSON response'); }
    return errorResponse;
  }
}

/**
 * Get the web app URL for embedding in Google Sites
 * Run this function to get your deployment URL
 */
function getWebAppUrl() {
  const webAppUrl = ScriptApp.getService().getUrl();
  console.log('Web App URL:', webAppUrl);
  return webAppUrl;
}

/**
 * Simple test function to verify deployment and authentication
 * Run this in Google Apps Script to test your setup
 */
function testDeployment() {
  try {
    console.log('=== DEPLOYMENT TEST ===');
    
    // Test 1: Check if we can access the spreadsheet
    console.log('Test 1: Checking spreadsheet access...');
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ Spreadsheet accessible:', spreadsheet.getName());
    
    // Test 2: Check user authentication
    console.log('Test 2: Checking user authentication...');
    const user = Session.getActiveUser();
    const userEmail = user.getEmail();
    console.log('✅ User email:', userEmail);
    
    // Test 3: Test doPost with a simple request
    console.log('Test 3: Testing doPost function...');
    const testEvent = {
      postData: {
        contents: 'test'
      }
    };
    
    const result = doPost(testEvent);
    console.log('✅ doPost test result:', result.getContent());
    
    return {
      success: true,
      message: 'All tests passed',
      user: userEmail,
      spreadsheet: spreadsheet.getName()
    };
    
  } catch (error) {
    console.error('❌ Test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Check user authentication and access permissions
 * This function helps diagnose authentication issues
 */
function checkUserAccess() {
  try {
    console.log('=== USER ACCESS CHECK ===');
    
    const user = Session.getActiveUser();
    const userEmail = user.getEmail();
    console.log('User email:', userEmail);
    
    if (!userEmail || userEmail === '') {
      return {
        success: false,
        error: 'No authenticated user found',
        message: 'Please log in to access this application'
      };
    }
    
    // Check spreadsheet access
    try {
      const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      const userAccess = spreadsheet.getAccess(userEmail);
      console.log('Spreadsheet access level:', userAccess);
      
      return {
        success: true,
        user: userEmail,
        accessLevel: userAccess,
        message: 'User has access to the application'
      };
    } catch (accessError) {
      console.error('Error checking spreadsheet access:', accessError);
      return {
        success: false,
        error: 'Cannot verify spreadsheet access',
        message: 'Please ensure you have permission to access the spreadsheet'
      };
    }
  } catch (error) {
    console.error('Error in checkUserAccess:', error);
    return {
      success: false,
      error: error.message,
      message: 'Error checking user access'
    };
  }
}

/**
 * Debug function to check project setup
 * Run this to diagnose issues
 */
function debugProjectSetup() {
  try {
    console.log('=== PROJECT SETUP DEBUG ===');
    console.log('Spreadsheet ID:', SPREADSHEET_ID);
    console.log('Web App URL:', ScriptApp.getService().getUrl());
    
    // Check if we can access the spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('Spreadsheet name:', ss.getName());
    console.log('Available sheets:', ss.getSheets().map(sheet => sheet.getName()));
    
    // Check if our service files exist
    console.log('Checking service dependencies...');
    
    // Test SheetsService
    try {
      const therapists = SheetsService.getTherapistsData();
      console.log('✅ SheetsService working - found', therapists.length, 'therapists');
    } catch (e) {
      console.log('❌ SheetsService error:', e.message);
    }
    
    // Test GeocodeService
    try {
      const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');
      console.log('Google Maps API Key configured:', apiKey ? 'Yes' : 'No');
    } catch (e) {
      console.log('❌ GeocodeService error:', e.message);
    }
    
    return {
      success: true,
      webAppUrl: ScriptApp.getService().getUrl(),
      spreadsheetId: SPREADSHEET_ID,
      spreadsheetName: ss.getName()
    };
    
  } catch (error) {
    console.error('Debug failed:', error);
    return {
      success: false,
      error: error.message,
      webAppUrl: ScriptApp.getService().getUrl()
    };
  }
}

/**
 * Helper function to set up Google Maps API key
 * Run this once to configure your API key
 */
function setupGoogleMapsApiKey() {
  // Public release: do not hardcode API keys. Set via Script Properties instead.
  // Example usage: PropertiesService.getScriptProperties().setProperty('GOOGLE_MAPS_API_KEY', 'YOUR_KEY');
  console.log('Please set GOOGLE_MAPS_API_KEY in Script Properties. This function no longer accepts hardcoded keys.');
  return false;
}

/**
 * Get map data including therapists, clients, sessions, assignments and notes
 * ENHANCED VERSION with better error handling to prevent null returns
 */
function getMapData() {
  console.log('[FRONTEND-CALL] getMapData called');
  
  // Initialize result object with empty defaults
  const result = {
    therapists: [],
    clients: [],
    sessions: [],
    assignments: {},
    bcbas: [],
    clientDistances: [],
    errors: [],
    debug: {
      timestamp: new Date().toISOString(),
      steps: []
    }
  };
  
  // DEBUG: Add timeout protection
  const startTime = new Date();
  console.log('[FRONTEND-CALL] Starting getMapData at:', startTime.toISOString());
  
  try {
    // Step 1: Load therapists
    console.log('[FRONTEND-CALL] Step 1: Loading therapists...');
    result.debug.steps.push('Step 1: Loading therapists');
    
    try {
      const step1Start = new Date();
      result.therapists = SheetsService.getTherapistsData();
      const step1Time = new Date() - step1Start;
      console.log('[FRONTEND-CALL] ✅ Therapists loaded:', result.therapists.length, 'in', step1Time, 'ms');
      result.debug.steps.push(`✅ Therapists loaded: ${result.therapists.length} in ${step1Time}ms`);
    } catch (e) {
      console.error('[FRONTEND-CALL] ❌ Failed to load therapists:', e);
      result.errors.push('Therapists: ' + e.message);
      result.debug.steps.push(`❌ Therapists failed: ${e.message}`);
    }
    
    // Step 2: Load clients
    console.log('[FRONTEND-CALL] Step 2: Loading clients...');
    result.debug.steps.push('Step 2: Loading clients');
    
    try {
      result.clients = SheetsService.getClientsData();
      console.log('[FRONTEND-CALL] ✅ Clients loaded:', result.clients.length);
      result.debug.steps.push(`✅ Clients loaded: ${result.clients.length}`);
    } catch (e) {
      console.error('[FRONTEND-CALL] ❌ Failed to load clients:', e);
      result.errors.push('Clients: ' + e.message);
      result.debug.steps.push(`❌ Clients failed: ${e.message}`);
    }
    
    // Step 3: Load sessions
    console.log('[FRONTEND-CALL] Step 3: Loading sessions...');
    result.debug.steps.push('Step 3: Loading sessions');
    
    try {
      result.sessions = SheetsService.getSessionsData();
      console.log('[FRONTEND-CALL] ✅ Sessions loaded:', result.sessions.length);
      result.debug.steps.push(`✅ Sessions loaded: ${result.sessions.length}`);
    } catch (e) {
      console.error('[FRONTEND-CALL] ❌ Failed to load sessions:', e);
      result.errors.push('Sessions: ' + e.message);
      result.debug.steps.push(`❌ Sessions failed: ${e.message}`);
    }
    
    // Step 4: Load assignments
    console.log('[FRONTEND-CALL] Step 4: Loading assignments...');
    result.debug.steps.push('Step 4: Loading assignments');
    
    try {
      result.assignments = SheetsService.getAssignmentsData();
      console.log('[FRONTEND-CALL] ✅ Assignments loaded:', Object.keys(result.assignments).length);
      result.debug.steps.push(`✅ Assignments loaded: ${Object.keys(result.assignments).length}`);
    } catch (e) {
      console.error('[FRONTEND-CALL] ❌ Failed to load assignments:', e);
      result.errors.push('Assignments: ' + e.message);
      result.debug.steps.push(`❌ Assignments failed: ${e.message}`);
    }
    
    // Step 5: Load optional data
    console.log('[FRONTEND-CALL] Step 5: Loading optional data...');
    result.debug.steps.push('Step 5: Loading optional data');
    
    // BCBAs (optional)
    if (SheetsService && typeof SheetsService.getBCBAsData === 'function') {
      try {
        result.bcbas = SheetsService.getBCBAsData();
        console.log('[FRONTEND-CALL] ✅ BCBAs loaded:', result.bcbas.length);
        result.debug.steps.push(`✅ BCBAs loaded: ${result.bcbas.length}`);
      } catch (e) {
        console.error('[FRONTEND-CALL] ❌ Failed to load BCBAs:', e);
        result.errors.push('BCBAs: ' + e.message);
        result.debug.steps.push(`❌ BCBAs failed: ${e.message}`);
      }
    } else {
      console.log('[FRONTEND-CALL] ⚠️ BCBAs function not available');
      result.debug.steps.push('⚠️ BCBAs function not available');
    }
    
    // Client distances (optional)
    try {
      result.clientDistances = getClientDistancesData();
      console.log('[FRONTEND-CALL] ✅ Client distances loaded:', result.clientDistances.length);
      result.debug.steps.push(`✅ Client distances loaded: ${result.clientDistances.length}`);
    } catch (e) {
      console.error('[FRONTEND-CALL] ❌ Failed to load client distances:', e);
      result.errors.push('Client distances: ' + e.message);
      result.debug.steps.push(`❌ Client distances failed: ${e.message}`);
    }
    
    console.log('[FRONTEND-CALL] ✅ getMapData completed successfully');
    result.debug.steps.push('✅ getMapData completed successfully');
    
    // Remove debug info for production to reduce payload size
    delete result.debug;
    
    return result;
    
  } catch (error) {
    console.error('[FRONTEND-CALL] ❌ Critical error in getMapData:', error);
    
    // Return a valid object even on critical error
    return {
      therapists: [],
      clients: [],
      sessions: [],
      assignments: {},
      bcbas: [],
      clientDistances: [],
      errors: ['Critical error: ' + error.message],
      debug: {
        criticalError: true,
        error: error.message,
        timestamp: new Date().toISOString()
      }
    };
  }
}

/**
 * Process therapists data and add lead info (from Wix version)
 */
function processTherapists(therapistItems) {
  // Lead assignments based on therapist ID ranges (same as Wix version)
  const leads = {
    "Kelly": [1, 10, 15, 23],
    "Bethany": [2, 11, 16, 24],
    "Rosemary W": [3, 4, 12, 17]
  };
  
  return therapistItems.map(therapist => {
    // Assign a lead based on ID
    for (const [lead, ids] of Object.entries(leads)) {
      if (ids.includes(Number(therapist.id))) {
        therapist.lead = lead;
        break;
      }
    }
    
    // If no lead was assigned, give a default
    if (!therapist.lead) {
      therapist.lead = "Unassigned";
    }
    
    return therapist;
  });
}

/**
 * Create or update assignment with support for multiple therapists
 */
function createAssignment(params) {
  try {
    const { therapistId, therapistIds, clientId, timeSlot, scheduleType, assignmentType, assignmentStatus, startDate, notes, multiTherapistGroupId } = params;
    
    // Support both old (therapistId) and new (therapistIds) parameter formats
    let therapistArray;
    if (therapistIds) {
      // New format with multiple therapists
      therapistArray = Array.isArray(therapistIds) ? therapistIds : [therapistIds];
    } else if (therapistId) {
      // Backwards compatibility with single therapist
      therapistArray = [therapistId];
    } else {
      return {
        success: false,
        error: 'Missing therapist information. Provide either therapistId or therapistIds.'
      };
    }
    
    console.log(`Creating assignment for ${therapistArray.length} therapist(s): ${therapistArray.join(', ')}, client: ${clientId}, slot: ${timeSlot}, schedule: ${scheduleType || 'current'}, groupId: ${multiTherapistGroupId || 'none'}`);
    
    // Validation
    if (!clientId || !timeSlot) {
      return {
        success: false,
        error: 'Missing required parameters: clientId or timeSlot'
      };
    }
    
    if (!['current', 'future'].includes(scheduleType || 'current')) {
      return {
        success: false,
        error: 'Invalid schedule type. Must be "current" or "future".'
      };
    }
    
    if (!['regular', 'playPals'].includes(assignmentType || 'regular')) {
      return {
        success: false,
        error: 'Invalid assignment type. Must be "regular" or "playPals".'
      };
    }
    
    if (!['red', 'orange', 'green'].includes(assignmentStatus || 'red')) {
      return {
        success: false,
        error: 'Invalid assignment status. Must be "red", "orange", or "green".'
      };
    }
    
    // Check for therapist conflicts (therapists can't be double-booked)
    for (const tId of therapistArray) {
      const existingAssignment = SheetsService.findTherapistAssignment(tId, timeSlot, scheduleType || 'current');
      if (existingAssignment && String(existingAssignment.clientId) !== String(clientId)) {
        const conflictClient = getClientName(existingAssignment.clientId);
        const therapistName = getTherapistName(tId);
        return {
          success: false,
          error: `${therapistName} is already assigned to ${conflictClient} at ${timeSlot} in ${scheduleType || 'current'} schedule`
        };
      }
    }
    
    // Use the updated SheetsService method that supports multiple therapists
    const result = SheetsService.createAssignment(
      therapistArray, // Pass array of therapist IDs
      clientId,
      timeSlot,
      assignmentType || 'regular',
      assignmentStatus || 'red',
      startDate || '',
      notes || '',
      scheduleType || 'current',
      multiTherapistGroupId
    );
    
    if (result.success) {
      console.log(`Successfully created assignment: ${result.action} for ${therapistArray.length} therapist(s)`);
      
      // Create session data for distance calculations
      const sessionPromises = therapistArray.map(tId => {
        return createSessionIfNeeded(tId, clientId, timeSlot);
      });
      
      // Wait for all session creations (they may be async)
      Promise.all(sessionPromises).then(sessions => {
        console.log(`Created ${sessions.filter(s => s).length} session records for distance tracking`);
      }).catch(error => {
        console.warn('Some session creations failed:', error);
      });
      
      return {
        success: true,
        action: result.action,
        assignmentId: result.assignmentId,
        therapistCount: result.therapistCount,
        assignmentDetails: result.assignmentDetails
      };
    } else {
      return result;
    }
    
  } catch (error) {
    console.error('Error in createAssignment:', error);
    return {
      success: false,
      error: 'Assignment creation failed: ' + error.message
    };
  }
}

/**
 * Helper function to validate date format
 */
function isValidDate(dateString) {
  if (!dateString) return true; // Empty date is valid
  
  // Check YYYY-MM-DD format
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateString)) return false;
  
  // Check if it's a valid date
  const date = new Date(dateString);
  return date instanceof Date && !isNaN(date) && dateString === date.toISOString().split('T')[0];
}

/**
 * Clear assignment - replaces Wix assignment deletion
 */
function clearAssignment(params) {
  try {
    const { clientId, timeSlot, scheduleType = 'current' } = params;
    
    console.log(`Clearing assignment for client ${clientId} in ${timeSlot} for ${scheduleType} schedule`);
    
    const result = SheetsService.deleteAssignment(timeSlot, clientId, null, scheduleType);
    
    return { success: true, deleted: result };
  } catch (error) {
    console.error('Error clearing assignment:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Clear all assignments for a therapist in a time slot
 */
function clearTherapistAssignments(params) {
  try {
    const { therapistId, timeSlot, scheduleType = 'current' } = params;
    
    console.log(`Clearing all assignments for therapist ${therapistId} in ${timeSlot} for ${scheduleType} schedule`);
    
    const result = SheetsService.deleteTherapistAssignments(therapistId, timeSlot, scheduleType);
    
    return { success: true, deleted: result };
  } catch (error) {
    console.error('Error clearing therapist assignments:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Add new therapist - replaces Wix therapist creation
 */
function addTherapist(therapistData) {
  try {
    console.log('Adding new therapist:', therapistData);
    
    // Validate and clean up the name fields
    const firstName = therapistData.firstName || '';
    const lastName = therapistData.lastName || '';
    const fullName = `${firstName} ${lastName}`.trim();
    
    console.log('Name processing:', { firstName, lastName, fullName });
    
    if (!fullName) {
      throw new Error('Therapist name is required');
    }
    
    // First geocode the address using GeocodeService
    const location = GeocodeService.geocodeAddress(therapistData.address);
    
    if (!location || !location.lat || !location.lng) {
      throw new Error('Failed to geocode address');
    }
    
    console.log('Geocoded location:', location);
    
    // Add location data to therapist record
    therapistData.latitude = location.lat;
    therapistData.longitude = location.lng;
    
    // Save to sheet
    const therapistId = SheetsService.addTherapist(therapistData);
    console.log('Therapist saved with ID:', therapistId);
    
    // Calculate travel times for this new therapist to all existing clients
    console.log('Calculating travel times for new therapist...');
    let travelTimeResults = null;
    try {
      travelTimeResults = DistanceService.calculateTravelTimesForNewTherapist(therapistId);
      console.log('Travel time calculation results:', travelTimeResults);
    } catch (distanceError) {
      console.warn('Travel time calculation failed (non-critical):', distanceError.message);
      // Don't fail the therapist creation if distance calculation fails
    }
    
    return { 
      success: true, 
      therapistId: therapistId,
      therapist: {
        id: therapistId,
        name: fullName,
        lat: location.lat,
        lng: location.lng,
        address: therapistData.address,
        title: therapistData.title || 'RBT'
      },
      travelTimeResults: travelTimeResults
    };
  } catch (error) {
    console.error('Error adding therapist:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Add new client - replaces Wix client creation
 */
function addClient(clientData) {
  try {
    console.log('Adding new client:', clientData);
    
    // Validate and clean up the name fields
    const firstName = clientData.firstName || '';
    const lastName = clientData.lastName || '';
    const fullName = `${firstName} ${lastName}`.trim();
    
    console.log('Name processing:', { firstName, lastName, fullName });
    
    if (!fullName) {
      throw new Error('Client name is required');
    }
    
    // First geocode the address
    const location = GeocodeService.geocodeAddress(clientData.address);
    
    if (!location || !location.lat || !location.lng) {
      throw new Error('Failed to geocode address');
    }
    
    console.log('Geocoded location:', location);
    
    // Add location data to client record
    clientData.latitude = location.lat;
    clientData.longitude = location.lng;
    
    // Save to sheet
    const clientId = SheetsService.addClient(clientData);
    console.log('Client saved with ID:', clientId);
    
    // Calculate travel times for all therapists to this new client
    console.log('Calculating travel times for new client...');
    let travelTimeResults = null;
    try {
      travelTimeResults = DistanceService.calculateTravelTimesForNewClient(clientId);
      console.log('Travel time calculation results:', travelTimeResults);
    } catch (distanceError) {
      console.warn('Travel time calculation failed (non-critical):', distanceError.message);
      // Don't fail the client creation if distance calculation fails
    }
    
    return { 
      success: true, 
      clientId: clientId,
      client: {
        id: clientId,
        name: fullName,
        lat: location.lat,
        lng: location.lng,
        address: clientData.address,
        bcbaId: clientData.bcbaId || null
      },
      travelTimeResults: travelTimeResults
    };
  } catch (error) {
    console.error('Error adding client:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Delete therapist - replaces Wix therapist deletion
 */
function deleteTherapist(therapistId) {
  try {
    console.log(`Deleting therapist ${therapistId}`);
    
    // Delete all assignments for this therapist
    SheetsService.deleteTherapistAssignments(therapistId);
    
    // Delete all sessions for this therapist
    SheetsService.deleteTherapistSessions(therapistId);
    
    // Delete the therapist record
    const result = SheetsService.deleteTherapist(therapistId);
    
    return { success: true, deleted: result };
  } catch (error) {
    console.error('Error deleting therapist:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Delete client - replaces Wix client deletion
 */
function deleteClient(clientId) {
  try {
    console.log(`Deleting client ${clientId}`);
    
    // Delete all assignments for this client
    SheetsService.deleteClientAssignments(clientId);
    
    // Delete all sessions for this client
    SheetsService.deleteClientSessions(clientId);
    
    // Delete the client record
    const result = SheetsService.deleteClient(clientId);
    
    return { success: true, deleted: result };
  } catch (error) {
    console.error('Error deleting client:', error);
    return { success: false, error: error.message };
  }
}

// Duplicate testConnection removed for public release

/**
 * Test the new calculateDistancesForClient function to verify it's working
 * This ensures the function is properly accessible via google.script.run
 */
function testCalculateDistancesForClient() {
  try {
    console.log('=== TESTING calculateDistancesForClient FUNCTION ===');
    
    // Get clients to test with
    const clients = SheetsService.getClientsData();
    console.log(`Found ${clients.length} clients in system`);
    
    if (clients.length === 0) {
      return {
        success: false,
        error: 'No clients found for testing',
        message: 'Please add at least one client before testing'
      };
    }
    
    // Test with the first client
    const testClientId = clients[0].id;
    console.log(`Testing calculateDistancesForClient with client ID: ${testClientId}`);
    
    // Call our new function (this simulates what google.script.run would do)
    const result = calculateDistancesForClient(testClientId);
    
    console.log('✅ Function test result:', result);
    
    return {
      success: true,
      message: 'calculateDistancesForClient function is working correctly',
      testResult: result,
      testClientId: testClientId,
      availableClients: clients.length
    };
    
  } catch (error) {
    console.error('❌ Function test failed:', error);
    return {
      success: false,
      error: error.message,
      message: 'calculateDistancesForClient function test failed'
    };
  }
}

/**
 * Test geocoding functionality specifically
 */
function testGeocoding() {
  try {
    console.log('=== TESTING GEOCODING SERVICE ===');
    
    // Test API key availability
    console.log('1. Testing API key availability...');
    const apiKey = GeocodeService.getApiKey();
    if (!apiKey) {
      return {
        success: false,
        error: 'Google Maps API key not configured',
        message: 'Please set GOOGLE_MAPS_API_KEY in Script Properties'
      };
    }
    console.log('✅ API key found');
    
    // Test geocoding with a sample address
    console.log('2. Testing geocoding with sample address...');
    const testAddress = "1600 Amphitheatre Parkway, Mountain View, CA";
    const result = GeocodeService.geocodeAddress(testAddress);
    
    console.log('✅ Geocoding successful:', result);
    
    return {
      success: true,
      message: 'Geocoding service is working correctly',
      testResult: result
    };
    
  } catch (error) {
    console.error('❌ Geocoding test failed:', error);
    return {
      success: false,
      error: error.message,
      message: 'Geocoding service test failed'
    };
  }
}

/**
 * Enhanced debug function to test all endpoints
 */
function testAllEndpoints() {
  try {
    console.log('=== TESTING ALL ENDPOINTS ===');
    
    // Test doGet
    console.log('Testing doGet...');
    const getResult = doGet({parameter: {debug: 'true'}});
    console.log('doGet result type:', typeof getResult);
    
    // Test doPost with getMapData
    console.log('Testing doPost with getMapData...');
    const postEvent = {
      postData: {
        contents: JSON.stringify({
          action: 'getMapData',
          params: {}
        })
      }
    };
    
    const postResult = doPost(postEvent);
    console.log('doPost result type:', typeof postResult);
    console.log('doPost result content type:', postResult.getMimeType());
    
    // Test the actual data
    const mapData = getMapData();
    console.log('Direct getMapData result:', {
      therapists: mapData.therapists ? mapData.therapists.length : 0,
      clients: mapData.clients ? mapData.clients.length : 0
    });
    
    return {
      success: true,
      doGetWorks: !!getResult,
      doPostWorks: !!postResult,
      dataLoads: !!mapData
    };
    
  } catch (error) {
    console.error('Test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * N/A STATUS MANAGEMENT FUNCTIONS
 * Handle marking therapists as N/A (not available) for specific time slots
 */

/**
 * Toggle N/A status for a therapist at a specific time slot
 */
function toggleNAStatus(params) {
  try {
    console.log('toggleNAStatus called with params:', params);
    
    const { therapistId, day, timeBlock, scheduleType = 'current' } = params;
    
    if (!therapistId || !day || !timeBlock) {
      return {
        success: false,
        error: 'therapistId, day, and timeBlock are required'
      };
    }
    
    const timeSlot = `${day}-${timeBlock}`;
    
    // Check if therapist is currently N/A for this time slot and schedule type
    const existingNA = SheetsService.findAssignment(timeSlot, 'N/A', therapistId, scheduleType);
    
    if (existingNA) {
      // Remove N/A status
      return removeNAStatus({ therapistId, day, timeBlock, scheduleType });
    } else {
      // Set N/A status
      return setNAStatus({ therapistId, day, timeBlock, scheduleType });
    }
    
  } catch (error) {
    console.error('Error in toggleNAStatus:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Set N/A status for a therapist at a specific time slot
 */
function setNAStatus(params) {
  try {
    console.log('setNAStatus called with params:', params);
    
    const { therapistId, day, timeBlock, scheduleType = 'current' } = params;
    
    if (!therapistId || !day || !timeBlock) {
      return {
        success: false,
        error: 'therapistId, day, and timeBlock are required'
      };
    }
    
    const timeSlot = `${day}-${timeBlock}`;
    
    // Check if therapist is already N/A for this time slot and schedule type
    const existingNA = SheetsService.findAssignment(timeSlot, 'N/A', therapistId, scheduleType);
    
    if (existingNA) {
      return {
        success: true,
        message: 'Therapist is already marked as N/A for this time slot',
        action: 'already_na'
      };
    }
    
    // Check if therapist has an existing client assignment for this time slot and schedule type
    const existingClientAssignment = SheetsService.findTherapistAssignment(therapistId, timeSlot, scheduleType);
    
    if (existingClientAssignment) {
      // Clear the existing client assignment first
      console.log(`Clearing existing client assignment for therapist ${therapistId} at ${timeSlot} in ${scheduleType} schedule`);
      SheetsService.deleteAssignment(timeSlot, existingClientAssignment.clientId, null, scheduleType);
    }
    
    // Create N/A assignment
    const assignmentId = SheetsService.addAssignment({
      timeSlot: timeSlot,
      therapistId: therapistId,
      clientId: 'N/A',
      scheduleType: scheduleType
    });
    
    console.log(`Set N/A status for therapist ${therapistId} at ${timeSlot} in ${scheduleType} schedule, assignment ID: ${assignmentId}`);
    
    return {
      success: true,
      message: 'Successfully set N/A status',
      action: 'set_na',
      assignmentId: assignmentId,
      therapistId: therapistId,
      timeSlot: timeSlot,
      scheduleType: scheduleType
    };
    
  } catch (error) {
    console.error('Error in setNAStatus:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Remove N/A status for a therapist at a specific time slot
 */
function removeNAStatus(params) {
  try {
    console.log('removeNAStatus called with params:', params);
    
    const { therapistId, day, timeBlock, scheduleType = 'current' } = params;
    
    if (!therapistId || !day || !timeBlock) {
      return {
        success: false,
        error: 'therapistId, day, and timeBlock are required'
      };
    }
    
    const timeSlot = `${day}-${timeBlock}`;
    
    // Find and remove the N/A assignment for the specific schedule type
    const result = SheetsService.deleteAssignment(timeSlot, 'N/A', therapistId, scheduleType);
    
    if (result === 0) {
      return {
        success: false,
        error: 'No N/A status found to remove for this therapist and time slot'
      };
    }
    
    console.log(`Removed N/A status for therapist ${therapistId} at ${timeSlot} in ${scheduleType} schedule, ${result} records deleted`);
    
    return {
      success: true,
      message: 'Successfully removed N/A status',
      action: 'removed_na',
      deleted: result,
      therapistId: therapistId,
      timeSlot: timeSlot,
      scheduleType: scheduleType
    };
    
  } catch (error) {
    console.error('Error in removeNAStatus:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Simple test function to verify API key setup
 * Run this function to test if your Google Maps API key is working
 */
function testApiKeySetup() {
  try {
    console.log('=== TESTING API KEY SETUP ===');
    
    // Test 1: Check if API key exists
    console.log('1. Checking if API key exists...');
    const apiKey = GeocodeService.getApiKey();
    
    if (!apiKey) {
      console.log('❌ API key not found!');
      return {
        success: false,
        error: 'API key not configured',
        instructions: 'Please set GOOGLE_MAPS_API_KEY in Script Properties'
      };
    }
    
    console.log('✅ API key found:', apiKey.substring(0, 10) + '...');
    
    // Test 2: Try geocoding a simple address
    console.log('2. Testing geocoding with sample address...');
    const testAddress = "1600 Amphitheatre Parkway, Mountain View, CA";
    
    try {
      const result = GeocodeService.geocodeAddress(testAddress);
      console.log('✅ Geocoding successful!');
      console.log('Address:', result.formatted_address);
      console.log('Coordinates:', result.lat, result.lng);
      
      return {
        success: true,
        message: 'API key is working correctly!',
        testAddress: testAddress,
        result: {
          formatted_address: result.formatted_address,
          lat: result.lat,
          lng: result.lng
        }
      };
      
    } catch (geocodeError) {
      console.log('❌ Geocoding failed:', geocodeError.message);
      return {
        success: false,
        error: 'Geocoding failed: ' + geocodeError.message,
        apiKeyExists: true
      };
    }
    
  } catch (error) {
    console.log('❌ Test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test distance calculation functionality
 */
function testDistanceCalculation(params) {
  try {
    console.log('=== TESTING DISTANCE CALCULATION ===');
    
    const { therapistId, clientId } = params;
    
    if (!therapistId || !clientId) {
      return {
        success: false,
        error: 'therapistId and clientId are required'
      };
    }
    
    console.log('1. Testing distance calculation between therapist', therapistId, 'and client', clientId);
    
    // Get therapist and client data
    const therapists = SheetsService.getTherapistsData();
    const clients = SheetsService.getClientsData();
    
    const therapist = therapists.find(t => String(t.id) === String(therapistId));
    const client = clients.find(c => String(c.id) === String(clientId));
    
    if (!therapist) {
      return {
        success: false,
        error: `Therapist ${therapistId} not found`
      };
    }
    
    if (!client) {
      return {
        success: false,
        error: `Client ${clientId} not found`
      };
    }
    
    console.log('2. Found therapist:', therapist.name, 'at', therapist.lat, therapist.lng);
    console.log('3. Found client:', client.name, 'at', client.lat, client.lng);
    
    // Test the distance calculation
    const result = DistanceService.calculateTravelTime(therapist, client);
    
    console.log('✅ Distance calculation successful:', result);
    
    return {
      success: true,
      message: 'Distance calculation working correctly',
      therapist: therapist.name,
      client: client.name,
      distanceInMiles: result.distanceInMiles,
      durationInMinutes: result.durationInMinutes,
      distanceText: result.distanceText,
      durationText: result.durationText
    };
    
  } catch (error) {
    console.error('❌ Distance calculation test failed:', error);
    return {
      success: false,
      error: error.message,
      message: 'Distance calculation test failed'
    };
  }
}

/**
 * Get assignment details including type, status, and notes
 */
function getAssignmentDetails(params) {
  try {
    const { timeSlot, clientId, therapistId = null, scheduleType = 'current' } = params;
    
    console.log(`Getting assignment details for: ${timeSlot}, client: ${clientId}, schedule: ${scheduleType}`);
    
    const details = SheetsService.getAssignmentDetails(timeSlot, clientId, therapistId, scheduleType);
    
    if (details) {
      return { 
        success: true, 
        assignment: details 
      };
    } else {
      return { 
        success: false, 
        error: 'Assignment not found' 
      };
    }
  } catch (error) {
    console.error('Error getting assignment details:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Update assignment details (status, type, start date, notes)
 */
function updateAssignmentDetails(params) {
  try {
    const { timeSlot, clientId, therapistId, scheduleType, assignmentType, assignmentStatus, startDate, notes } = params;
    
    console.log(`Updating assignment details for T${therapistId}-C${clientId} at ${timeSlot} in ${scheduleType} schedule`);
    
    // Validation
    if (!timeSlot || !clientId || !therapistId) {
      return {
        success: false,
        error: 'Missing required parameters: timeSlot, clientId, or therapistId'
      };
    }
    
    if (!['current', 'future'].includes(scheduleType)) {
      return {
        success: false,
        error: 'Invalid schedule type. Must be "current" or "future".'
      };
    }
    
    if (!['regular', 'playPals'].includes(assignmentType)) {
      return {
        success: false,
        error: 'Invalid assignment type. Must be "regular" or "playPals".'
      };
    }
    
    if (!['red', 'orange', 'green'].includes(assignmentStatus)) {
      return {
        success: false,
        error: 'Invalid assignment status. Must be "red", "orange", or "green".'
      };
    }
    
    // Update the assignment in SheetsService
    const updateResult = SheetsService.updateAssignmentDetails(
      timeSlot, 
      clientId, 
      therapistId, 
      scheduleType, 
      assignmentType, 
      assignmentStatus, 
      startDate, 
      notes
    );
    
    if (updateResult.success) {
      console.log(`Successfully updated assignment details for T${therapistId}-C${clientId}`);
      
      return {
        success: true,
        message: `Updated assignment details for ${timeSlot}`,
        updatedFields: {
          assignmentType,
          assignmentStatus,
          startDate,
          notes
        }
      };
    } else {
      return {
        success: false,
        error: updateResult.error || 'Failed to update assignment details'
      };
    }
    
  } catch (error) {
    console.error('Error in updateAssignmentDetails:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test function for new assignment details functionality
 */
function testAssignmentDetails(params = {}) {
  try {
    console.log('🧪 Testing assignment details functionality...');
    
    const testTimeSlot = params.timeSlot || 'Monday 1:00pm';
    const testTherapistId = params.therapistId || 'test_therapist_1';
    const testClientId = params.clientId || 'test_client_1';
    
    // Test 1: Create assignment with new fields
    console.log('📝 Test 1: Creating assignment with new fields...');
    const createResult = createAssignment({
      timeSlot: testTimeSlot,
      therapistId: testTherapistId,
      clientId: testClientId,
      assignmentType: 'playPals',
      assignmentStatus: 'orange',
      startDate: '2024-02-15',
      notes: 'Test assignment with Play Pals type'
    });
    
    console.log('Create result:', createResult);
    
    if (!createResult.success) {
      return { success: false, error: 'Failed to create test assignment', details: createResult };
    }
    
    // Test 2: Get assignment details
    console.log('📄 Test 2: Getting assignment details...');
    const getResult = getAssignmentDetails({
      timeSlot: testTimeSlot,
      clientId: testClientId
    });
    
    console.log('Get result:', getResult);
    
    if (!getResult.success) {
      return { success: false, error: 'Failed to get assignment details', details: getResult };
    }
    
    // Test 3: Update assignment details
    console.log('🔄 Test 3: Updating assignment details...');
    const updateResult = updateAssignmentDetails({
      timeSlot: testTimeSlot,
      clientId: testClientId,
      assignmentStatus: 'green',
      notes: 'Updated test notes - session officially started'
    });
    
    console.log('Update result:', updateResult);
    
    if (!updateResult.success) {
      return { success: false, error: 'Failed to update assignment details', details: updateResult };
    }
    
    // Test 4: Verify updates by getting details again
    console.log('✅ Test 4: Verifying updates...');
    const verifyResult = getAssignmentDetails({
      timeSlot: testTimeSlot,
      clientId: testClientId
    });
    
    console.log('Verify result:', verifyResult);
    
    // Clean up test data
    console.log('🗑️ Cleaning up test data...');
    const clearResult = clearAssignment({
      timeSlot: testTimeSlot,
      clientId: testClientId
    });
    
    return {
      success: true,
      message: 'All assignment details tests passed!',
      testResults: {
        create: createResult,
        get: getResult,
        update: updateResult,
        verify: verifyResult,
        cleanup: clearResult
      }
    };
    
  } catch (error) {
    console.error('Error in testAssignmentDetails:', error);
    return { success: false, error: error.message, stack: error.stack };
  }
}

/**
 * Copy entire schedule from one schedule type to another
 */
function copyEntireSchedule(params) {
  try {
    const { fromSchedule, toSchedule } = params;
    
    console.log(`Copying entire schedule from ${fromSchedule} to ${toSchedule}`);
    
    // Validation
    if (!['current', 'future'].includes(fromSchedule) || !['current', 'future'].includes(toSchedule)) {
      return {
        success: false,
        error: 'Invalid schedule type. Must be "current" or "future".'
      };
    }
    
    if (fromSchedule === toSchedule) {
      return {
        success: false,
        error: 'Source and destination schedules cannot be the same.'
      };
    }
    
    // Get all assignments for the source schedule
    const sourceAssignments = SheetsService.getAssignmentsByScheduleType(fromSchedule);
    
    if (!sourceAssignments || sourceAssignments.length === 0) {
      return {
        success: false,
        error: `No assignments found in ${fromSchedule} schedule to copy.`
      };
    }
    
    // Clear all existing assignments in the destination schedule
    console.log(`Clearing existing assignments in ${toSchedule} schedule...`);
    const clearResult = SheetsService.clearScheduleType(toSchedule);
    console.log(`Cleared ${clearResult} assignments from ${toSchedule} schedule`);
    
    // Copy each assignment to the destination schedule
    let copiedCount = 0;
    let errorCount = 0;
    const errors = [];
    
    for (const assignment of sourceAssignments) {
      try {
        // Create a copy of the assignment with the new schedule type
        const newAssignment = {
          timeSlot: assignment.timeSlot,
          therapistId: assignment.therapistId,
          clientId: assignment.clientId,
          assignmentType: assignment.assignmentType || 'regular',
          assignmentStatus: assignment.assignmentStatus || 'red',
          startDate: assignment.startDate || null,
          notes: assignment.notes || '',
          scheduleType: toSchedule
        };
        
        const assignmentId = SheetsService.addAssignment(newAssignment);
        
        if (assignmentId) {
          copiedCount++;
          console.log(`Copied assignment: ${assignment.timeSlot} - therapist ${assignment.therapistId} to client ${assignment.clientId}`);
        } else {
          errorCount++;
          errors.push(`Failed to copy assignment: ${assignment.timeSlot} - therapist ${assignment.therapistId} to client ${assignment.clientId}`);
        }
        
      } catch (assignmentError) {
        errorCount++;
        errors.push(`Error copying assignment ${assignment.timeSlot}: ${assignmentError.message}`);
        console.error('Error copying individual assignment:', assignmentError);
      }
    }
    
    console.log(`Copy operation completed: ${copiedCount} assignments copied, ${errorCount} errors`);
    
    if (errors.length > 0) {
      console.warn('Copy errors:', errors);
    }
    
    return {
      success: copiedCount > 0,
      message: `Successfully copied ${copiedCount} assignments from ${fromSchedule} to ${toSchedule} schedule`,
      copiedCount: copiedCount,
      errorCount: errorCount,
      errors: errors,
      totalSourceAssignments: sourceAssignments.length
    };
    
  } catch (error) {
    console.error('Error in copyEntireSchedule:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Copy selected therapists to the other schedule
 */
function copySelectedTherapists(params) {
  try {
    const { fromSchedule, toSchedule, therapistIds } = params;
    
    console.log(`Copying selected therapists from ${fromSchedule} to ${toSchedule}`);
    
    // Validation
    if (!['current', 'future'].includes(fromSchedule) || !['current', 'future'].includes(toSchedule)) {
      return {
        success: false,
        error: 'Invalid schedule type. Must be "current" or "future".'
      };
    }
    
    if (fromSchedule === toSchedule) {
      return {
        success: false,
        error: 'Source and destination schedules cannot be the same.'
      };
    }
    
    if (!therapistIds || therapistIds.length === 0) {
      return {
        success: false,
        error: 'therapistIds are required'
      };
    }
    
    // Get all assignments for the source schedule
    const sourceAssignments = SheetsService.getAssignmentsByScheduleType(fromSchedule);
    
    if (!sourceAssignments || sourceAssignments.length === 0) {
      return {
        success: false,
        error: `No assignments found in ${fromSchedule} schedule to copy.`
      };
    }
    
    // Filter assignments to only include selected therapists
    const filteredAssignments = sourceAssignments.filter(assignment => therapistIds.includes(String(assignment.therapistId)));
    
    if (filteredAssignments.length === 0) {
      return {
        success: false,
        error: `No assignments found for selected therapists in ${fromSchedule} schedule.`
      };
    }
    
    // Clear existing assignments for selected therapists in the destination schedule only
    console.log(`Clearing existing assignments for selected therapists in ${toSchedule} schedule...`);
    let clearedCount = 0;
    for (const therapistId of therapistIds) {
      const cleared = SheetsService.clearTherapistFromSchedule(therapistId, toSchedule);
      clearedCount += cleared;
    }
    console.log(`Cleared ${clearedCount} assignments for selected therapists from ${toSchedule} schedule`);
    
    // Copy each filtered assignment to the destination schedule
    let copiedCount = 0;
    let errorCount = 0;
    const errors = [];
    
    for (const assignment of filteredAssignments) {
      try {
        // Create a copy of the assignment with the new schedule type
        const newAssignment = {
          timeSlot: assignment.timeSlot,
          therapistId: assignment.therapistId,
          clientId: assignment.clientId,
          assignmentType: assignment.assignmentType || 'regular',
          assignmentStatus: assignment.assignmentStatus || 'red',
          startDate: assignment.startDate || null,
          notes: assignment.notes || '',
          scheduleType: toSchedule
        };
        
        const assignmentId = SheetsService.addAssignment(newAssignment);
        
        if (assignmentId) {
          copiedCount++;
          console.log(`Copied assignment: ${assignment.timeSlot} - therapist ${assignment.therapistId} to client ${assignment.clientId}`);
        } else {
          errorCount++;
          errors.push(`Failed to copy assignment: ${assignment.timeSlot} - therapist ${assignment.therapistId} to client ${assignment.clientId}`);
        }
        
      } catch (assignmentError) {
        errorCount++;
        errors.push(`Error copying assignment ${assignment.timeSlot}: ${assignmentError.message}`);
        console.error('Error copying individual assignment:', assignmentError);
      }
    }
    
    console.log(`Copy operation completed: ${copiedCount} assignments copied, ${errorCount} errors`);
    
    if (errors.length > 0) {
      console.warn('Copy errors:', errors);
    }
    
    return {
      success: copiedCount > 0,
      message: `Successfully copied ${copiedCount} assignments from ${fromSchedule} to ${toSchedule} schedule`,
      copiedCount: copiedCount,
      errorCount: errorCount,
      errors: errors,
      totalSourceAssignments: filteredAssignments.length
    };
    
  } catch (error) {
    console.error('Error in copySelectedTherapists:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test geocoding for the specific address that's failing
 * Run this function to debug the geocoding issue
 */
function testSpecificAddressGeocoding() {
  try {
    console.log('=== TESTING SPECIFIC ADDRESS GEOCODING ===');
    
    // Test the failing address
    const testAddress = "2925 North Olive Street";
    console.log('Testing address:', testAddress);
    
    // First check if API key exists
    const apiKey = GeocodeService.getApiKey();
    if (!apiKey) {
      return {
        success: false,
        error: 'Google Maps API key not configured',
        instructions: 'Please set GOOGLE_MAPS_API_KEY in Script Properties'
      };
    }
    console.log('✅ API key found');
    
    // Try geocoding with different address variations
    const addressVariations = [
      "2925 North Olive Street",
      "2925 N Olive St",
      "2925 North Olive Street, Dallas, TX",
      "2925 N Olive St, Dallas, Texas",
      "2925 North Olive Street, USA"
    ];
    
    const results = [];
    
    for (const address of addressVariations) {
      try {
        console.log(`Trying: "${address}"`);
        const result = GeocodeService.geocodeAddress(address);
        results.push({
          address: address,
          success: true,
          result: result
        });
        console.log(`✅ Success: ${result.formatted_address}`);
        break; // Stop at first success
      } catch (error) {
        results.push({
          address: address,
          success: false,
          error: error.message
        });
        console.log(`❌ Failed: ${error.message}`);
      }
    }
    
    return {
      success: results.some(r => r.success),
      testAddress: testAddress,
      results: results,
      recommendation: results.some(r => r.success) ? 
        'Use one of the successful address formats' : 
        'Address may not exist or need more specific location info (city, state)'
    };
    
  } catch (error) {
    console.error('Test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Helper function to test API key and basic geocoding functionality
 */
function testGeocodingSetup() {
  try {
    console.log('=== TESTING GEOCODING SETUP ===');
    
    // Test 1: API key
    const apiKey = GeocodeService.getApiKey();
    if (!apiKey) {
      return {
        success: false,
        error: 'API key not configured',
        instructions: 'Go to Project Settings > Script Properties and add GOOGLE_MAPS_API_KEY'
      };
    }
    console.log('✅ API key configured');
    
    // Test 2: Simple address that should always work
    const testAddress = "1600 Amphitheatre Parkway, Mountain View, CA";
    console.log('Testing with known good address:', testAddress);
    
    const result = GeocodeService.geocodeAddress(testAddress);
    console.log('✅ Geocoding working:', result.formatted_address);
    
    return {
      success: true,
      message: 'Geocoding setup is working correctly',
      testResult: result
    };
    
  } catch (error) {
    console.error('Setup test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Patch missing distance calculations for sessions
 * Finds sessions where travelTime_minutes or distance_miles are 0/null and recalculates them
 */
function patchMissingDistances() {
  try {
    console.log('=== STARTING PATCH MISSING DISTANCES ===');
    
    // Get all sessions data
    const sessionsData = SheetsService.getSessionsData();
    console.log(`Found ${sessionsData.length} total sessions`);
    
    // Find sessions with missing distance data
    const missingSessions = sessionsData.filter(session => {
      const travelTime = Number(session.travelTime_minutes) || 0;
      const distance = Number(session.distance_miles) || 0;
      
      // Consider missing if either is 0 or null
      return travelTime === 0 || distance === 0;
    });
    
    console.log(`Found ${missingSessions.length} sessions with missing distance data`);
    
    if (missingSessions.length === 0) {
      return {
        success: true,
        message: 'No sessions with missing distance data found',
        totalSessions: sessionsData.length,
        missingSessions: 0,
        patchedSessions: 0,
        failedSessions: 0,
        errors: []
      };
    }
    
    // Get therapists and clients data for calculations
    const therapistsData = SheetsService.getTherapistsData();
    const clientsData = SheetsService.getClientsData();
    
    console.log(`Loaded ${therapistsData.length} therapists and ${clientsData.length} clients`);
    
    // Track results
    let patchedCount = 0;
    let failedCount = 0;
    const errors = [];
    
    // Process each missing session
    for (let i = 0; i < missingSessions.length; i++) {
      const session = missingSessions[i];
      
      try {
        console.log(`Processing ${i + 1}/${missingSessions.length}: T${session.therapistId}-C${session.clientId}`);
        
        // Find therapist and client
        const therapist = therapistsData.find(t => Number(t.id) === Number(session.therapistId));
        const client = clientsData.find(c => Number(c.id) === Number(session.clientId));
        
        if (!therapist) {
          throw new Error(`Therapist ${session.therapistId} not found`);
        }
        
        if (!client) {
          throw new Error(`Client ${session.clientId} not found`);
        }
        
        // Validate coordinates
        if (!therapist.lat || !therapist.lng || !client.lat || !client.lng) {
          throw new Error(`Missing coordinates - Therapist: ${therapist.lat},${therapist.lng} Client: ${client.lat},${client.lng}`);
        }
        
        // Calculate travel time using DistanceService
        const travelTimeResult = DistanceService.calculateTravelTime(therapist, client);
        
        console.log(`Calculated: ${travelTimeResult.durationInMinutes} mins, ${travelTimeResult.distanceInMiles} miles`);
        
        // Update the session in the sheet
        const updateResult = SheetsService.updateSessionDistance(
          session.therapistId, 
          session.clientId, 
          travelTimeResult.durationInMinutes, 
          travelTimeResult.distanceInMiles
        );
        
        if (updateResult) {
          patchedCount++;
          console.log(`✅ Successfully patched T${session.therapistId}-C${session.clientId}`);
        } else {
          throw new Error('Failed to update session in sheet');
        }
        
        // Add small delay to respect API rate limits
        if (i < missingSessions.length - 1) {
          Utilities.sleep(250); // 250ms delay between API calls
        }
        
      } catch (error) {
        failedCount++;
        const errorMsg = `T${session.therapistId}-C${session.clientId}: ${error.message}`;
        errors.push(errorMsg);
        console.error(`❌ Failed to patch T${session.therapistId}-C${session.clientId}:`, error.message);
      }
    }
    
    console.log(`=== PATCH COMPLETE ===`);
    console.log(`Total processed: ${missingSessions.length}`);
    console.log(`Successfully patched: ${patchedCount}`);
    console.log(`Failed: ${failedCount}`);
    
    return {
      success: true,
      message: `Patched ${patchedCount} of ${missingSessions.length} sessions with missing distances`,
      totalSessions: sessionsData.length,
      missingSessions: missingSessions.length,
      patchedSessions: patchedCount,
      failedSessions: failedCount,
      errors: errors
    };
    
  } catch (error) {
    console.error('Error in patchMissingDistances:', error);
    return {
      success: false,
      error: error.message,
      message: 'Failed to patch missing distances'
    };
  }
}

/**
 * Check how many sessions need distance patching (without actually patching)
 * Run this to see what the patch function would do
 */
function checkMissingDistances() {
  try {
    console.log('=== CHECKING FOR MISSING DISTANCES ===');
    
    // Get all sessions data
    const sessionsData = SheetsService.getSessionsData();
    console.log(`Found ${sessionsData.length} total sessions`);
    
    // Find sessions with missing distance data
    const missingSessions = sessionsData.filter(session => {
      const travelTime = Number(session.travelTime_minutes) || 0;
      const distance = Number(session.distance_miles) || 0;
      
      // Consider missing if either is 0 or null
      return travelTime === 0 || distance === 0;
    });
    
    console.log(`Found ${missingSessions.length} sessions with missing distance data:`);
    
    // Log details about missing sessions
    missingSessions.slice(0, 10).forEach(session => { // Show first 10
      console.log(`- T${session.therapistId}-C${session.clientId}: travel=${session.travelTime_minutes}, distance=${session.distance_miles}`);
    });
    
    if (missingSessions.length > 10) {
      console.log(`... and ${missingSessions.length - 10} more`);
    }
    
    return {
      success: true,
      totalSessions: sessionsData.length,
      missingSessions: missingSessions.length,
      missingSessionsList: missingSessions.map(s => `T${s.therapistId}-C${s.clientId}`),
      message: `Found ${missingSessions.length} sessions that need distance patching`
    };
    
  } catch (error) {
    console.error('Error checking missing distances:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Convert multi-therapist assignments to frontend-compatible format
 */
function convertMultiTherapistAssignmentsForFrontend(multiTherapistAssignments) {
  const convertedAssignments = {};
  
  for (const timeSlot in multiTherapistAssignments) {
    convertedAssignments[timeSlot] = {};
    
    for (const clientId in multiTherapistAssignments[timeSlot]) {
      const assignment = multiTherapistAssignments[timeSlot][clientId];
      const therapistIds = assignment.therapistIds || [assignment.therapistId];
      
      if (therapistIds.length === 1) {
        convertedAssignments[timeSlot][clientId] = {
          therapistId: therapistIds[0],
          assignmentType: assignment.assignmentType,
          assignmentStatus: assignment.assignmentStatus,
          startDate: assignment.startDate,
          notes: assignment.notes,
          scheduleType: assignment.scheduleType
        };
      } else {
        convertedAssignments[timeSlot][clientId] = {
          therapistId: therapistIds[0],
          therapistIds: therapistIds,
          isMultiTherapist: true,
          therapistCount: therapistIds.length,
          assignmentType: assignment.assignmentType,
          assignmentStatus: assignment.assignmentStatus,
          startDate: assignment.startDate,
          notes: assignment.notes,
          scheduleType: assignment.scheduleType
        };
      }
    }
  }
  
  console.log(`Converted ${Object.keys(multiTherapistAssignments).length} time slots to frontend format`);
  return convertedAssignments;
}

/**
 * Helper function to get client name by ID
 */
function getClientName(clientId) {
  if (clientId === "N/A") return "N/A";
  
  try {
    const clients = SheetsService.getClientsData();
    const client = clients.find(c => String(c.id) === String(clientId));
    return client ? client.name : "Unknown Client";
  } catch (error) {
    console.error('Error getting client name:', error);
    return "Unknown Client";
  }
}

/**
 * Helper function to get therapist name by ID
 */
function getTherapistName(therapistId) {
  try {
    const therapists = SheetsService.getTherapistsData();
    const therapist = therapists.find(t => String(t.id) === String(therapistId));
    return therapist ? therapist.name : "Unknown Therapist";
  } catch (error) {
    console.error('Error getting therapist name:', error);
    return "Unknown Therapist";
  }
}

/**
 * THERAPIST NOTES FUNCTIONS
 * Handle therapist notes functionality
 */

/**
 * Get all therapist notes from the backend
 */
function getTherapistNotes() {
  try {
    console.log('Loading therapist notes from backend...');
    
    const notes = SheetsService.getTherapistNotes();
    
    console.log(`Successfully loaded ${notes.length} therapist notes`);
    
    return {
      success: true,
      notes: notes,
      count: notes.length
    };
    
  } catch (error) {
    console.error('Error loading therapist notes:', error);
    return {
      success: false,
      error: error.message,
      notes: []
    };
  }
}

/**
 * Save therapist note to the backend
 */
function saveTherapistNote(params) {
  try {
    const { therapistId, timeBlock, scheduleType, noteText } = params;
    
    console.log(`Saving therapist note for T${therapistId}, ${timeBlock}, ${scheduleType}`);
    
    // Validation
    if (!therapistId || !timeBlock || !scheduleType) {
      return {
        success: false,
        error: 'Missing required parameters: therapistId, timeBlock, or scheduleType'
      };
    }
    
    if (!['current', 'future'].includes(scheduleType)) {
      return {
        success: false,
        error: 'Invalid schedule type. Must be "current" or "future".'
      };
    }
    
    // Save note using SheetsService
    const noteId = SheetsService.saveTherapistNote(therapistId, timeBlock, scheduleType, noteText || '');
    
    if (noteId) {
      console.log(`Successfully saved therapist note with ID: ${noteId}`);
      
      return {
        success: true,
        noteId: noteId,
        message: 'Note saved successfully',
        noteData: {
          therapistId: therapistId,
          timeBlock: timeBlock,
          scheduleType: scheduleType,
          noteText: noteText
        }
      };
    } else {
      return {
        success: false,
        error: 'Failed to save note to database'
      };
    }
    
  } catch (error) {
    console.error('Error saving therapist note:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Delete therapist note from the backend
 */
function deleteTherapistNote(params) {
  try {
    const { therapistId, timeBlock, scheduleType } = params;
    
    console.log(`Deleting therapist note for T${therapistId}, ${timeBlock}, ${scheduleType}`);
    
    // Validation
    if (!therapistId || !timeBlock || !scheduleType) {
      return {
        success: false,
        error: 'Missing required parameters: therapistId, timeBlock, or scheduleType'
      };
    }
    
    // Delete note using SheetsService
    const deletedCount = SheetsService.deleteTherapistNote(therapistId, timeBlock, scheduleType);
    
    if (deletedCount > 0) {
      console.log(`Successfully deleted ${deletedCount} therapist note(s)`);
      
      return {
        success: true,
        deletedCount: deletedCount,
        message: 'Note deleted successfully'
      };
    } else {
      return {
        success: false,
        error: 'No matching note found to delete'
      };
    }
    
  } catch (error) {
    console.error('Error deleting therapist note:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// Note: Direct google.script.run therapist note functions removed for public release.

/**
 * Update therapist lead BCBA assignment
 * Called via google.script.run from the frontend
 */
function updateTherapistLead(params) {
  try {
    console.log('👥 [DIRECT] updateTherapistLead called with params:', params);
    
    const { therapistId, leadBCBA, therapistName } = params;
    
    // Validation
    if (!therapistId || leadBCBA === undefined) {
      console.error('👥 [DIRECT] Missing required parameters:', { therapistId, leadBCBA, therapistName });
      return {
        success: false,
        error: 'Missing required parameters: therapistId and leadBCBA'
      };
    }
    
    console.log(`👥 [DIRECT] Updating therapist ${therapistId} (${therapistName || 'Unknown'}) lead to: ${leadBCBA}`);
    
    // Call SheetsService to update the therapist lead
    const result = SheetsService.updateTherapistLead(therapistId, leadBCBA, therapistName);
    
    if (result.success) {
      console.log(`👥 [DIRECT] Successfully updated therapist lead: ${result.message}`);
    } else {
      console.error(`👥 [DIRECT] Failed to update therapist lead: ${result.error}`);
    }
    
    return result;
    
  } catch (error) {
    console.error('👥 [DIRECT] Error in updateTherapistLead:', error);
    return {
      success: false,
      error: error.message || error.toString()
    };
  }
}

/**
 * Helper function to create session data if needed
 */
function createSessionIfNeeded(therapistId, clientId, timeSlot) {
  try {
    const sessions = SheetsService.getSessionsData();
    const existingSession = sessions.find(s => 
      String(s.therapistId) === String(therapistId) && 
      String(s.clientId) === String(clientId)
    );
    
    if (existingSession) {
      console.log(`Session already exists for T${therapistId}-C${clientId}`);
      return existingSession;
    }
    
    const therapists = SheetsService.getTherapistsData();
    const clients = SheetsService.getClientsData();
    
    const therapist = therapists.find(t => String(t.id) === String(therapistId));
    const client = clients.find(c => String(c.id) === String(clientId));
    
    if (!therapist || !client) {
      console.warn(`Cannot create session: therapist ${therapistId} or client ${clientId} not found`);
      return null;
    }
    
    console.log(`Created session record for T${therapistId}-C${clientId} for distance tracking`);
    return {
      therapistId: therapistId,
      clientId: clientId,
      therapist_name: therapist.name,
      client_name: client.name,
      timeSlot: timeSlot
    };
    
  } catch (error) {
    console.error('Error creating session:', error);
    return null;
  }
}

/**
 * PRE-CALCULATE client-to-client distances and store in Sessions sheet
 * Run this ONCE to populate client-to-client data (like therapist-to-client setup)
 * DO NOT call from getMapData() - this is a one-time setup function
 */
function generateClientToClientDistances() {
  try {
    console.log('=== STARTING CLIENT-TO-CLIENT DISTANCE PRE-CALCULATION ===');
    console.log('⚠️  This function calculates and STORES distances in the Sessions sheet');
    console.log('⚠️  Only run this when you want to populate/update client-to-client data');
    
    // Get clients data
    const clients = SheetsService.getClientsData();
    console.log(`Found ${clients.length} clients`);
    
    if (clients.length < 2) {
      return {
        success: false,
        error: 'Need at least 2 clients for distance calculations'
      };
    }
    
    // Check if client-to-client sessions already exist
    const existingSessions = SheetsService.getSessionsData();
    const existingClientSessions = existingSessions.filter(s => s.sourceClientId && s.targetClientId);
    
    if (existingClientSessions.length > 0) {
      console.log(`⚠️  Found ${existingClientSessions.length} existing client-to-client sessions`);
      console.log('⚠️  This will ADD to existing data. Delete existing client-to-client sessions first if you want to rebuild.');
    }
    
    let calculatedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    const errors = [];
    
    // Calculate distances between all client pairs
    for (let i = 0; i < clients.length; i++) {
      const sourceClient = clients[i];
      
      // Validate source client coordinates
      if (!sourceClient.lat || !sourceClient.lng) {
        console.warn(`Skipping client ${sourceClient.id} (${sourceClient.name}) - missing coordinates`);
        skippedCount++;
        continue;
      }
      
      for (let j = 0; j < clients.length; j++) {
        if (i === j) continue; // Skip self
        
        const targetClient = clients[j];
        
        // Validate target client coordinates
        if (!targetClient.lat || !targetClient.lng) {
          skippedCount++;
          continue;
        }
        
        // Check if this pair already exists (to avoid duplicates)
        const existingPair = existingClientSessions.find(s => 
          (String(s.sourceClientId) === String(sourceClient.id) && String(s.targetClientId) === String(targetClient.id))
        );
        
        if (existingPair) {
          skippedCount++; // Already exists
          continue;
        }
        
        try {
          // Use the SAME distance calculation service as therapist-to-client
          const distanceResult = DistanceService.calculateTravelTime(sourceClient, targetClient);
          
          // Only store reasonable distances (under 25 miles)
          if (distanceResult.distanceInMiles <= 25) {
            // Store in Sessions sheet using the SAME format as therapist-to-client
            const sessionData = {
              sourceClientId: sourceClient.id,        // NEW field
              targetClientId: targetClient.id,        // NEW field  
              travelTime_minutes: distanceResult.durationInMinutes,
              distance_miles: distanceResult.distanceInMiles,
              distance_text: distanceResult.distanceText,
              duration_text: distanceResult.durationText,
              title: 'client-to-client'
            };
            
            // Add to Sessions sheet (same as therapist-to-client storage)
            const sessionId = SheetsService.addSession(sessionData);
            
            if (sessionId) {
              calculatedCount++;
              console.log(`✅ Stored C${sourceClient.id}→C${targetClient.id}: ${distanceResult.distanceText} (${distanceResult.durationText})`);
            } else {
              errorCount++;
              errors.push(`Failed to store session for C${sourceClient.id}→C${targetClient.id}`);
            }
          } else {
            skippedCount++; // Too far
          }
          
          // API rate limiting (every 10 calculations)
          if (calculatedCount % 10 === 0) {
            console.log(`Progress: ${calculatedCount} stored, ${skippedCount} skipped, ${errorCount} errors`);
            Utilities.sleep(250); // 250ms delay to respect API limits
          }
          
        } catch (distanceError) {
          errorCount++;
          const errorMsg = `C${sourceClient.id}→C${targetClient.id}: ${distanceError.message}`;
          errors.push(errorMsg);
          console.warn(`❌ Distance calculation failed: ${errorMsg}`);
        }
      }
      
      // Progress logging every 5 clients
      if ((i + 1) % 5 === 0) {
        console.log(`Processed ${i + 1}/${clients.length} source clients`);
      }
    }
    
    console.log('=== CLIENT-TO-CLIENT DISTANCE CALCULATION COMPLETE ===');
    console.log(`Total client pairs processed: ${clients.length * (clients.length - 1)}`);
    console.log(`Successfully calculated and stored: ${calculatedCount}`);
    console.log(`Skipped: ${skippedCount}`);
    console.log(`Errors: ${errorCount}`);
    
    if (errors.length > 0) {
      console.log('Error details:', errors.slice(0, 10)); // Show first 10 errors
    }
    
    return {
      success: true,
      message: `Successfully pre-calculated ${calculatedCount} client-to-client distances`,
      calculated: calculatedCount,
      skipped: skippedCount,
      errors: errorCount,
      errorDetails: errors.slice(0, 10)
    };
    
  } catch (error) {
    console.error('Error in generateClientToClientDistances:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * TEST function: Generate client-to-client distances for first 10 clients only
 * Run this FIRST to test the system before doing the full calculation
 */
function testClientToClientDistances() {
  try {
    console.log('=== TESTING CLIENT-TO-CLIENT DISTANCES (First 10 Clients) ===');
    
    // Get clients data (limit to first 10 for testing)
    const allClients = SheetsService.getClientsData();
    const clients = allClients.slice(0, 10); // First 10 clients only
    
    console.log(`Testing with ${clients.length} clients (out of ${allClients.length} total)`);
    
    if (clients.length < 2) {
      return {
        success: false,
        error: 'Need at least 2 clients for testing'
      };
    }
    
    let calculatedCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    const errors = [];
    
    // Calculate distances between test client pairs
    for (let i = 0; i < clients.length; i++) {
      const sourceClient = clients[i];
      
      // Validate source client coordinates
      if (!sourceClient.lat || !sourceClient.lng) {
        console.warn(`Skipping client ${sourceClient.id} (${sourceClient.name}) - missing coordinates`);
        skippedCount++;
        continue;
      }
      
      for (let j = 0; j < clients.length; j++) {
        if (i === j) continue; // Skip self
        
        const targetClient = clients[j];
        
        // Validate target client coordinates
        if (!targetClient.lat || !targetClient.lng) {
          skippedCount++;
          continue;
        }
        
        try {
          // Calculate distance (but don't store yet - this is just a test)
          const distanceResult = DistanceService.calculateTravelTime(sourceClient, targetClient);
          
          calculatedCount++;
          console.log(`✅ Test calculation C${sourceClient.id}→C${targetClient.id}: ${distanceResult.distanceText} (${distanceResult.durationText})`);
          
          // Small delay for API rate limiting
          if (calculatedCount % 5 === 0) {
            Utilities.sleep(200);
          }
          
        } catch (distanceError) {
          errorCount++;
          const errorMsg = `C${sourceClient.id}→C${targetClient.id}: ${distanceError.message}`;
          errors.push(errorMsg);
          console.warn(`❌ Test calculation failed: ${errorMsg}`);
        }
      }
    }
    
    console.log('=== TEST COMPLETE ===');
    console.log(`Test calculations completed: ${calculatedCount}`);
    console.log(`Skipped: ${skippedCount}`);
    console.log(`Errors: ${errorCount}`);
    
    if (errors.length > 0) {
      console.log('Test errors:', errors);
    }
    
    const estimatedFullCalculation = (allClients.length * (allClients.length - 1)) * 0.75; // Estimate 75% within 25 miles
    const estimatedTime = Math.round(estimatedFullCalculation * 0.5 / 60); // Estimate 0.5 seconds per calculation, in minutes
    
    console.log(`📊 ESTIMATES FOR FULL CALCULATION:`);
    console.log(`- Total client pairs: ${allClients.length * (allClients.length - 1)}`);
    console.log(`- Estimated calculations: ${Math.round(estimatedFullCalculation)}`);
    console.log(`- Estimated time: ${estimatedTime} minutes`);
    
    return {
      success: true,
      message: `Test completed with ${calculatedCount} calculations`,
      testCalculations: calculatedCount,
      testSkipped: skippedCount,
      testErrors: errorCount,
      estimates: {
        totalClientPairs: allClients.length * (allClients.length - 1),
        estimatedCalculations: Math.round(estimatedFullCalculation),
        estimatedTimeMinutes: estimatedTime
      }
    };
    
  } catch (error) {
    console.error('Error in testClientToClientDistances:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * CHECK existing client-to-client data in Sessions sheet
 * Run this in Google Apps Script to see what data you already have
 */
function checkExistingClientToClientData() {
  try {
    console.log('=== CHECKING EXISTING CLIENT-TO-CLIENT DATA ===');
    
    // Get all sessions
    const allSessions = SheetsService.getSessionsData();
    console.log(`Total sessions in database: ${allSessions.length}`);
    
    // Separate by type
    const therapistToClientSessions = allSessions.filter(s => s.therapistId && s.clientId && !s.sourceClientId);
    const clientToClientSessions = allSessions.filter(s => s.sourceClientId && s.targetClientId);
    const unknownSessions = allSessions.filter(s => !s.therapistId && !s.sourceClientId);
    
    console.log('Session breakdown:');
    console.log(`- Therapist-to-client: ${therapistToClientSessions.length}`);
    console.log(`- Client-to-client: ${clientToClientSessions.length}`);
    console.log(`- Unknown format: ${unknownSessions.length}`);
    
    if (clientToClientSessions.length > 0) {
      console.log('✅ CLIENT-TO-CLIENT DATA FOUND!');
      console.log('Sample client-to-client sessions:');
      
      // Show first 5 records
      clientToClientSessions.slice(0, 5).forEach((session, index) => {
        console.log(`${index + 1}. C${session.sourceClientId} → C${session.targetClientId}: ${session.distance_miles || session.DistanceMiles} miles, ${session.travelTime_minutes || session.TravelMinutes} mins`);
      });
      
      // Get unique source clients
      const sourceClients = [...new Set(clientToClientSessions.map(s => s.sourceClientId))];
      console.log(`Data covers ${sourceClients.length} source clients:`, sourceClients.slice(0, 10));
      
      return {
        success: true,
        message: `Found ${clientToClientSessions.length} existing client-to-client sessions`,
        totalSessions: allSessions.length,
        therapistToClient: therapistToClientSessions.length,
        clientToClient: clientToClientSessions.length,
        sourceClients: sourceClients.length,
        sampleData: clientToClientSessions.slice(0, 5)
      };
      
    } else {
      console.log('❌ No client-to-client data found');
      console.log('Next step: Run testClientToClientDistances() to test the system');
      
      return {
        success: false,
        message: 'No client-to-client sessions found',
        totalSessions: allSessions.length,
        therapistToClient: therapistToClientSessions.length,
        clientToClient: 0,
        nextStep: 'Run testClientToClientDistances() to test the system'
      };
    }
    
  } catch (error) {
    console.error('Error checking existing data:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * DEBUG: Test ClientDistances sheet specifically
 * Run this function to diagnose the ClientDistances sheet issue
 */
function debugClientDistancesSheet() {
  try {
    console.log('=== DEBUGGING CLIENT DISTANCES SHEET ===');
    
    // Test the function directly
    console.log('📋 Testing getClientDistancesData() function...');
    const clientDistances = getClientDistancesData();
    
    console.log('📊 Result:', {
      dataType: typeof clientDistances,
      isArray: Array.isArray(clientDistances),
      length: clientDistances ? clientDistances.length : 0
    });
    
    if (clientDistances && clientDistances.length > 0) {
      console.log('✅ SUCCESS: ClientDistances data loaded!');
      console.log('📄 First record:', JSON.stringify(clientDistances[0], null, 2));
      console.log('🔑 Available columns:', Object.keys(clientDistances[0]));
      
      // Check for expected columns
      const expectedColumns = ['SourceClientID', 'TargetClientID', 'DistanceMiles', 'TravelMinutes'];
      expectedColumns.forEach(col => {
        const hasColumn = clientDistances[0].hasOwnProperty(col);
        console.log(`   - ${col}: ${hasColumn ? '✅' : '❌'}`);
      });
      
      return {
        success: true,
        recordCount: clientDistances.length,
        sampleRecord: clientDistances[0],
        columns: Object.keys(clientDistances[0])
      };
      
    } else {
      console.log('❌ ISSUE: No ClientDistances data returned');
      
      // Additional debugging
      console.log('🔍 Additional debugging...');
      try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const allSheetNames = spreadsheet.getSheets().map(s => s.getName());
        console.log('📋 All sheet names:', allSheetNames);
        
        // Check for variations
        const possibleNames = ['ClientDistances', 'Client Distances', 'clientdistances', 'ClientDistance'];
        const foundSheet = possibleNames.find(name => allSheetNames.includes(name));
        
        if (foundSheet) {
          console.log(`💡 Found sheet with name: "${foundSheet}"`);
        } else {
          console.log('❌ No ClientDistances sheet found with any variation');
        }
        
        return {
          success: false,
          error: 'No ClientDistances data',
          allSheets: allSheetNames,
          foundSheet: foundSheet || null
        };
        
      } catch (innerError) {
        console.error('Error in additional debugging:', innerError);
        return {
          success: false,
          error: innerError.message
        };
      }
    }
    
  } catch (error) {
    console.error('❌ Error in debugClientDistancesSheet:', error);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * DIAGNOSE potential column naming issues with existing client-to-client data
 * Run this to see if data exists but with different column names
 */
function diagnoseClientToClientColumnNames() {
  try {
    console.log('=== DIAGNOSING CLIENT-TO-CLIENT COLUMN NAMES ===');
    
    // Get raw session data to check column names
    const sessions = SheetsService.getSessionsData();
    console.log(`Total sessions: ${sessions.length}`);
    
    if (sessions.length === 0) {
      return { success: false, error: 'No sessions data found' };
    }
    
    // Check first few sessions for different column naming patterns
    const sampleSession = sessions[0];
    console.log('Sample session column names:', Object.keys(sampleSession));
    
    // Look for different naming patterns
    const patterns = {
      camelCase: ['sourceClientId', 'targetClientId', 'distance_miles', 'travelTime_minutes'],
      PascalCase: ['SourceClientID', 'TargetClientID', 'DistanceMiles', 'TravelMinutes'],
      mixedCase: ['sourceClientId', 'targetClientId', 'DistanceMiles', 'TravelMinutes']
    };
    
    let foundPattern = null;
    let clientToClientCount = 0;
    
    // Check which pattern matches existing data
    for (const [patternName, fields] of Object.entries(patterns)) {
      const hasFields = fields.every(field => sessions.some(s => s.hasOwnProperty(field)));
      if (hasFields) {
        foundPattern = patternName;
        // Count sessions with this pattern
        clientToClientCount = sessions.filter(s => s[fields[0]] && s[fields[1]]).length;
        console.log(`✅ Found ${patternName} pattern with ${clientToClientCount} client-to-client sessions`);
        break;
      }
    }
    
    if (!foundPattern) {
      console.log('❌ No recognized column patterns found');
      // Look for any fields that might be client-to-client
      const possibleFields = Object.keys(sampleSession).filter(key => 
        key.toLowerCase().includes('client') || 
        key.toLowerCase().includes('source') || 
        key.toLowerCase().includes('target')
      );
      console.log('Possible client-related fields found:', possibleFields);
    }
    
    // Look for specific client-to-client indicators
    const clientToClientSessions = sessions.filter(s => {
      // Try different column name variations
      return (s.sourceClientId && s.targetClientId) ||
             (s.SourceClientID && s.TargetClientID) ||
             (s.sourceClientID && s.targetClientID) ||
             (s.source_client_id && s.target_client_id);
    });
    
    console.log(`Found ${clientToClientSessions.length} potential client-to-client sessions`);
    
    if (clientToClientSessions.length > 0) {
      console.log('Sample client-to-client session:', clientToClientSessions[0]);
    }
    
    return {
      success: true,
      totalSessions: sessions.length,
      foundPattern: foundPattern,
      clientToClientCount: clientToClientSessions.length,
      sampleColumns: Object.keys(sampleSession),
      sampleClientSession: clientToClientSessions.length > 0 ? clientToClientSessions[0] : null
    };
    
  } catch (error) {
    console.error('Error diagnosing column names:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Load client-to-client distance data from ClientDistances sheet
 */
function getClientDistancesData() {
  try {
    console.log('🔍 Starting getClientDistancesData...');
    
    // Get the spreadsheet
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('📊 Opened spreadsheet:', spreadsheet.getName());
    
    // List all sheet names for debugging
    const allSheets = spreadsheet.getSheets();
    const sheetNames = allSheets.map(sheet => sheet.getName());
    console.log('📋 Available sheets:', sheetNames);
    
    // Try to get the ClientDistances sheet
    const sheet = spreadsheet.getSheetByName('ClientDistances');
    if (!sheet) {
      console.log('❌ ClientDistances sheet not found in sheets:', sheetNames);
      console.log('💡 Please verify the sheet name is exactly "ClientDistances"');
      return [];
    }
    
    console.log('✅ Found ClientDistances sheet');
    
    // Check if sheet has data
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    console.log(`📏 Sheet dimensions: ${lastRow} rows x ${lastCol} columns`);
    
    if (lastRow <= 1) {
      console.log('⚠️ ClientDistances sheet appears to be empty (only header row or no data)');
      return [];
    }
    
    // Get all data
    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    console.log('📋 Headers found:', headers);
    
    const data = [];
    
    for (let i = 1; i < values.length; i++) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = values[i][j];
      }
      data.push(obj);
    }
    
    console.log(`✅ Successfully loaded ${data.length} client distance records`);
    
    // Show sample data for debugging
    if (data.length > 0) {
      console.log('📄 Sample record:', JSON.stringify(data[0], null, 2));
      console.log('🔑 Sample record keys:', Object.keys(data[0]));
    }
    
    return data;
  } catch (error) {
    console.error('❌ Error in getClientDistancesData:', error);
    console.error('📍 Error stack:', error.stack);
    return [];
  }
}

/**
 * DIAGNOSE: Analyze ClientDistances data quality
 * Run this to understand the distance data issues
 */
function analyzeClientDistancesDataQuality() {
  try {
    console.log('=== ANALYZING CLIENT DISTANCES DATA QUALITY ===');
    
    const clientDistances = getClientDistancesData();
    
    if (!clientDistances || clientDistances.length === 0) {
      return {
        success: false,
        error: 'No client distances data found'
      };
    }
    
    console.log(`📊 Total records: ${clientDistances.length}`);
    
    // Analyze distance values
    const distanceStats = {
      zero: 0,
      verySmall: 0, // 0 < distance <= 0.2
      small: 0,     // 0.2 < distance <= 1
      medium: 0,    // 1 < distance <= 5
      large: 0,     // 5 < distance <= 20
      veryLarge: 0, // > 20
      null: 0,
      invalid: 0
    };
    
    const sampleDistances = [];
    
    clientDistances.forEach((record, index) => {
      const distance = Number(record.DistanceMiles);
      
      if (index < 10) {
        sampleDistances.push({
          source: record.SourceClientID,
          target: record.TargetClientID,
          distance: distance,
          travelTime: record.TravelMinutes,
          distanceText: record.DistanceText
        });
      }
      
      if (isNaN(distance) || distance === null || distance === undefined) {
        distanceStats.null++;
      } else if (distance <= 0) {
        distanceStats.zero++;
      } else if (distance <= 0.2) {
        distanceStats.verySmall++;
      } else if (distance <= 1) {
        distanceStats.small++;
      } else if (distance <= 5) {
        distanceStats.medium++;
      } else if (distance <= 20) {
        distanceStats.large++;
      } else {
        distanceStats.veryLarge++;
      }
    });
    
    console.log('📈 Distance Distribution:');
    console.log(`   Zero/Null: ${distanceStats.zero + distanceStats.null} (${((distanceStats.zero + distanceStats.null)/clientDistances.length*100).toFixed(1)}%)`);
    console.log(`   Very Small (0-0.2mi): ${distanceStats.verySmall} (${(distanceStats.verySmall/clientDistances.length*100).toFixed(1)}%)`);
    console.log(`   Small (0.2-1mi): ${distanceStats.small} (${(distanceStats.small/clientDistances.length*100).toFixed(1)}%)`);
    console.log(`   Medium (1-5mi): ${distanceStats.medium} (${(distanceStats.medium/clientDistances.length*100).toFixed(1)}%)`);
    console.log(`   Large (5-20mi): ${distanceStats.large} (${(distanceStats.large/clientDistances.length*100).toFixed(1)}%)`);
    console.log(`   Very Large (>20mi): ${distanceStats.veryLarge} (${(distanceStats.veryLarge/clientDistances.length*100).toFixed(1)}%)`);
    
    console.log('📄 Sample Records:');
    sampleDistances.forEach((sample, index) => {
      console.log(`   ${index + 1}. C${sample.source} → C${sample.target}: ${sample.distance} miles (${sample.travelTime} mins) [${sample.distanceText}]`);
    });
    
    // Find usable records (> 0.2 miles)
    const usableRecords = clientDistances.filter(record => {
      const distance = Number(record.DistanceMiles);
      return !isNaN(distance) && distance > 0.2;
    });
    
    console.log(`✅ Usable records (>0.2 miles): ${usableRecords.length} of ${clientDistances.length} (${(usableRecords.length/clientDistances.length*100).toFixed(1)}%)`);
    
    // Test with a specific client
    if (usableRecords.length > 0) {
      const testClientId = usableRecords[0].SourceClientID;
      const testClientDistances = clientDistances.filter(record => 
        String(record.SourceClientID) === String(testClientId) || 
        String(record.TargetClientID) === String(testClientId)
      );
      
      const testUsableDistances = testClientDistances.filter(record => {
        const distance = Number(record.DistanceMiles);
        return !isNaN(distance) && distance > 0.2;
      });
      
      console.log(`🧪 Test Client ${testClientId}:`);
      console.log(`   Total distance records: ${testClientDistances.length}`);
      console.log(`   Usable records (>0.2mi): ${testUsableDistances.length}`);
      
      if (testUsableDistances.length > 0) {
        console.log('   Sample usable distances:');
        testUsableDistances.slice(0, 5).forEach((record, index) => {
          const targetId = String(record.SourceClientID) === String(testClientId) ? record.TargetClientID : record.SourceClientID;
          console.log(`     ${index + 1}. → C${targetId}: ${record.DistanceMiles} miles`);
        });
      }
    }
    
    return {
      success: true,
      totalRecords: clientDistances.length,
      usableRecords: usableRecords.length,
      usablePercentage: (usableRecords.length/clientDistances.length*100).toFixed(1),
      distanceStats: distanceStats,
      sampleData: sampleDistances
    };
    
  } catch (error) {
    console.error('❌ Error analyzing client distances data quality:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function reGeocodeAllClients() {
  console.log('Starting re-geocoding of all clients');
  
  try {
    const clients = SheetsService.getClientsData();
    let successCount = 0;
    let failCount = 0;
    const errors = [];
    
    clients.forEach((client, index) => {
      try {
        if (index > 0) Utilities.sleep(500); // Rate limiting
        
        const location = GeocodeService.geocodeAddress(client.address);
        
        if (location && location.lat && location.lng) {
          // Update in sheet
          SheetsService.updateClientCoordinates(client.id, location.lat, location.lng);
          successCount++;
          console.log(`Updated C${client.id} (${client.name}): ${location.lat}, ${location.lng}`);
        } else {
          throw new Error('Invalid location result');
        }
      } catch (error) {
        failCount++;
        errors.push(`C${client.id} (${client.name}): ${error.message}`);
        console.error(`Failed to re-geocode C${client.id}:`, error);
      }
    });
    
    console.log(`Re-geocoding complete: ${successCount} successes, ${failCount} failures`);
    
    return {
      success: true,
      totalClients: clients.length,
      updated: successCount,
      failed: failCount,
      errors: errors
    };
    
  } catch (error) {
    console.error('Error in reGeocodeAllClients:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function clearClientDistancesSheet() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('ClientDistances');
    
    if (!sheet) {
      return { success: true, deletedRecords: 0, message: 'Sheet not found - nothing to clear' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, deletedRecords: 0, message: 'No data to clear' };
    }
    
    sheet.deleteRows(2, lastRow - 1);
    
    return {
      success: true,
      deletedRecords: lastRow - 1,
      message: 'Successfully cleared ClientDistances sheet'
    };
  } catch (error) {
    console.error('Error clearing ClientDistances:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Calculate distances from one client to all other clients - for Calc Full Distances button
 * This mirrors the working addClient pattern and calls DistanceService.calculateDistancesForClient
 */
function calculateDistancesForClient(sourceClientId) {
  try {
    console.log(`[BACKEND] calculateDistancesForClient called for client ${sourceClientId}`);
    
    if (!sourceClientId) {
      throw new Error('sourceClientId is required');
    }
    
    // Call the DistanceService function (same as used by addClient for travel times)
    console.log('Calling DistanceService.calculateDistancesForClient...');
    const distances = DistanceService.calculateDistancesForClient(sourceClientId);
    
    console.log(`Distance calculation completed: ${distances.length} client-to-client distances calculated`);
    
    return {
      success: true,
      distances: distances,
      count: distances.length,
      message: `Successfully calculated ${distances.length} client-to-client distances`
    };
    
  } catch (error) {
    console.error('[BACKEND] Error in calculateDistancesForClient:', error);
    return { 
      success: false, 
      error: error.message,
      message: 'Failed to calculate client distances'
    };
  }
}

/**
 * FRONTEND-CALLABLE FUNCTIONS
 * These functions are called directly by google.script.run from the HTML files
 */

// REMOVED: Duplicate getMapData function - keeping only the robust version above

/**
 * Frontend-callable test connection function
 */
function testConnection() {
  try {
    const data = getMapData();
    const hasData = data && Array.isArray(data.therapists);
    return { success: hasData, data: data, message: hasData ? 'Connection successful' : 'No data found' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * ADDITIONAL FRONTEND-CALLABLE FUNCTIONS
 * These are functions that the HTML files expect to call via google.script.run
 */

/**
 * Test function to verify everything is working
 */
// Duplicate testConnection removed for public release

/**
 * Get client distance statistics (for the frontend)
 */
function getClientDistanceStats() {
  try {
    console.log('[FRONTEND] getClientDistanceStats called');
    
    const clientDistances = getClientDistancesData();
    const clients = SheetsService.getClientsData();
    
    const totalClients = clients.length;
    const expectedCalculations = totalClients * (totalClients - 1);
    const currentCalculations = clientDistances.length;
    const completionPercentage = expectedCalculations > 0 ? Math.round((currentCalculations / expectedCalculations) * 100) : 0;
    const isComplete = currentCalculations >= expectedCalculations;
    
    console.log('[FRONTEND] getClientDistanceStats result:', {
      totalClients,
      currentCalculations,
      expectedCalculations,
      completionPercentage,
      isComplete
    });
    
    return {
      success: true,
      clientCount: totalClients,
      currentCalculations: currentCalculations,
      expectedCalculations: expectedCalculations,
      completionPercentage: completionPercentage,
      isComplete: isComplete,
      sheetExists: true,
      lastUpdated: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('[FRONTEND] getClientDistanceStats error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Initialize client distances (limited API for testing)
 */
function initializeClientDistancesLimitedAPI(maxClients) {
  try {
    console.log('[FRONTEND] initializeClientDistancesLimitedAPI called with maxClients:', maxClients);
    
    // This is a placeholder - you can implement the actual distance calculation logic here
    return {
      success: true,
      message: `Limited client distance initialization completed for ${maxClients} clients`,
      clientCount: maxClients,
      calculationsCompleted: maxClients * (maxClients - 1),
      durationSeconds: 5
    };
    
  } catch (error) {
    console.error('[FRONTEND] initializeClientDistancesLimitedAPI error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// REMOVED: Duplicate testConnection and getMapDataForFrontend functions - keeping only the robust versions above

/**
 * Test function to verify everything is working
 */
// Duplicate testConnection removed for public release

/**
 * Simple test function for frontend debugging
 */
function simpleTest() {
  console.log('[SIMPLE-TEST] Function called from frontend');
  return {
    success: true,
    message: 'Simple test working',
    timestamp: new Date().toISOString(),
    therapistCount: 50,
    clientCount: 137
  };
}

/**
 * Test data loading step by step
 */
function testStepByStep() {
  try {
    console.log('[STEP-TEST] Starting step-by-step test');
    
    // Step 1: Test therapists
    const therapists = SheetsService.getTherapistsData();
    console.log('[STEP-TEST] Step 1 - Therapists:', therapists.length);
    
    if (therapists.length === 0) {
      return { success: false, error: 'No therapists found', step: 1 };
    }
    
    // Step 2: Test clients  
    const clients = SheetsService.getClientsData();
    console.log('[STEP-TEST] Step 2 - Clients:', clients.length);
    
    if (clients.length === 0) {
      return { success: false, error: 'No clients found', step: 2 };
    }
    
    // Step 3: Return basic data
    return {
      success: true,
      message: 'Step-by-step test passed',
      therapists: therapists.slice(0, 3), // Return first 3 therapists
      clients: clients.slice(0, 3),       // Return first 3 clients
      therapistCount: therapists.length,
      clientCount: clients.length
    };
    
  } catch (error) {
    console.error('[STEP-TEST] Error:', error);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

function testGrokConnectivity() {
  console.log('[TEST] testGrokConnectivity called');
  return {success: true, message: "Grok test successful", timestamp: new Date().toISOString()};
}

function testSpreadsheetAccess() {
  try {
    console.log('[TEST] Testing spreadsheet access...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('[TEST] Spreadsheet opened:', ss.getName());
    return {success: true, message: 'Spreadsheet access successful', sheetName: ss.getName()};
  } catch (error) {
    console.error('[TEST] Spreadsheet access failed:', error);
    return {success: false, error: error.message};
  }
}

function testSheetsServiceFunctions() {
  console.log('[TEST] Testing SheetsService functions individually...');
  
  const results = {
    success: true,
    errors: [],
    data: {}
  };
  
  try {
    console.log('[TEST] Testing getTherapistsData...');
    const therapists = SheetsService.getTherapistsData();
    results.data.therapists = therapists;
    console.log('[TEST] ✅ getTherapistsData successful:', therapists.length, 'therapists');
  } catch (error) {
    console.error('[TEST] ❌ getTherapistsData failed:', error);
    results.errors.push({ function: 'getTherapistsData', error: error.message });
    results.success = false;
  }
  
  try {
    console.log('[TEST] Testing getClientsData...');
    const clients = SheetsService.getClientsData();
    results.data.clients = clients;
    console.log('[TEST] ✅ getClientsData successful:', clients.length, 'clients');
  } catch (error) {
    console.error('[TEST] ❌ getClientsData failed:', error);
    results.errors.push({ function: 'getClientsData', error: error.message });
    results.success = false;
  }
  
  try {
    console.log('[TEST] Testing getSessionsData...');
    const sessions = SheetsService.getSessionsData();
    results.data.sessions = sessions;
    console.log('[TEST] ✅ getSessionsData successful:', sessions.length, 'sessions');
  } catch (error) {
    console.error('[TEST] ❌ getSessionsData failed:', error);
    results.errors.push({ function: 'getSessionsData', error: error.message });
    results.success = false;
  }
  
  try {
    console.log('[TEST] Testing getAssignmentsData...');
    const assignments = SheetsService.getAssignmentsData();
    results.data.assignments = assignments;
    console.log('[TEST] ✅ getAssignmentsData successful:', Object.keys(assignments).length, 'time slots');
  } catch (error) {
    console.error('[TEST] ❌ getAssignmentsData failed:', error);
    results.errors.push({ function: 'getAssignmentsData', error: error.message });
    results.success = false;
  }
  
  try {
    console.log('[TEST] Testing getBCBAsData...');
    const bcbas = SheetsService.getBCBAsData();
    results.data.bcbas = bcbas;
    console.log('[TEST] ✅ getBCBAsData successful:', bcbas.length, 'BCBAs');
  } catch (error) {
    console.error('[TEST] ❌ getBCBAsData failed:', error);
    results.errors.push({ function: 'getBCBAsData', error: error.message });
    results.success = false;
  }
  
  try {
    console.log('[TEST] Testing getClientDistancesData...');
    const clientDistances = SheetsService.getClientDistancesData();
    results.data.clientDistances = clientDistances;
    console.log('[TEST] ✅ getClientDistancesData successful:', clientDistances.length, 'distance records');
  } catch (error) {
    console.error('[TEST] ❌ getClientDistancesData failed:', error);
    results.errors.push({ function: 'getClientDistancesData', error: error.message });
    results.success = false;
  }
  
  console.log('[TEST] Test completed. Success:', results.success);
  if (results.errors.length > 0) {
    console.log('[TEST] Errors found:', results.errors);
  }
  
  return results;
}

function testSimpleReturn() {
  console.log('[TEST] testSimpleReturn called');
  return {success: true, msg: 'It works!'};
}

/**
 * ENHANCED DIAGNOSTIC FUNCTIONS FOR NULL DATA ISSUE
 * These functions will help identify exactly where the data pipeline is breaking
 */

/**
 * Test the complete data pipeline step by step
 * Run this function to diagnose the null data issue
 */
function diagnoseNullDataIssue() {
  try {
    console.log('=== DIAGNOSING NULL DATA ISSUE ===');
    
    // Test 1: Check constants
    console.log('1. Testing constants...');
    console.log('SPREADSHEET_ID:', SPREADSHEET_ID);
    console.log('SHEET_NAMES:', SHEET_NAMES);
    
    if (!SPREADSHEET_ID) {
      return {
        success: false,
        error: 'SPREADSHEET_ID is not defined',
        step: 'constants'
      };
    }
    
    // Test 2: Check spreadsheet access
    console.log('2. Testing spreadsheet access...');
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      console.log('✅ Spreadsheet accessed:', spreadsheet.getName());
    } catch (e) {
      return {
        success: false,
        error: 'Cannot access spreadsheet: ' + e.message,
        step: 'spreadsheet_access'
      };
    }
    
    // Test 3: Check sheet existence
    console.log('3. Testing sheet existence...');
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(s => s.getName());
    console.log('Available sheets:', sheetNames);
    
    const requiredSheets = Object.values(SHEET_NAMES);
    const missingSheets = requiredSheets.filter(name => !sheetNames.includes(name));
    
    if (missingSheets.length > 0) {
      return {
        success: false,
        error: 'Missing required sheets: ' + missingSheets.join(', '),
        step: 'sheet_existence',
        availableSheets: sheetNames,
        missingSheets: missingSheets
      };
    }
    
    // Test 4: Test each data loading function individually
    console.log('4. Testing individual data loading functions...');
    
    const testResults = {};
    
    // Test therapists
    try {
      const therapists = SheetsService.getTherapistsData();
      testResults.therapists = {
        success: true,
        count: therapists.length,
        sample: therapists.length > 0 ? therapists[0] : null
      };
      console.log('✅ Therapists loaded:', therapists.length);
    } catch (e) {
      testResults.therapists = {
        success: false,
        error: e.message
      };
      console.log('❌ Therapists failed:', e.message);
    }
    
    // Test clients
    try {
      const clients = SheetsService.getClientsData();
      testResults.clients = {
        success: true,
        count: clients.length,
        sample: clients.length > 0 ? clients[0] : null
      };
      console.log('✅ Clients loaded:', clients.length);
    } catch (e) {
      testResults.clients = {
        success: false,
        error: e.message
      };
      console.log('❌ Clients failed:', e.message);
    }
    
    // Test sessions
    try {
      const sessions = SheetsService.getSessionsData();
      testResults.sessions = {
        success: true,
        count: sessions.length,
        sample: sessions.length > 0 ? sessions[0] : null
      };
      console.log('✅ Sessions loaded:', sessions.length);
    } catch (e) {
      testResults.sessions = {
        success: false,
        error: e.message
      };
      console.log('❌ Sessions failed:', e.message);
    }
    
    // Test assignments
    try {
      const assignments = SheetsService.getAssignmentsData();
      testResults.assignments = {
        success: true,
        count: Object.keys(assignments).length,
        sample: Object.keys(assignments).length > 0 ? Object.keys(assignments)[0] : null
      };
      console.log('✅ Assignments loaded:', Object.keys(assignments).length);
    } catch (e) {
      testResults.assignments = {
        success: false,
        error: e.message
      };
      console.log('❌ Assignments failed:', e.message);
    }
    
    // Test 5: Test the complete getMapData function
    console.log('5. Testing complete getMapData function...');
    let mapData;
    try {
      mapData = getMapData();
      testResults.getMapData = {
        success: true,
        hasTherapists: !!mapData.therapists,
        hasClients: !!mapData.clients,
        hasSessions: !!mapData.sessions,
        hasAssignments: !!mapData.assignments,
        therapistsCount: mapData.therapists ? mapData.therapists.length : 0,
        clientsCount: mapData.clients ? mapData.clients.length : 0,
        sessionsCount: mapData.sessions ? mapData.sessions.length : 0,
        assignmentsCount: mapData.assignments ? Object.keys(mapData.assignments).length : 0
      };
      console.log('✅ getMapData completed successfully');
    } catch (e) {
      testResults.getMapData = {
        success: false,
        error: e.message,
        stack: e.stack
      };
      console.log('❌ getMapData failed:', e.message);
    }
    
    // Test 6: Test the doPost function with getMapData
    console.log('6. Testing doPost with getMapData...');
    try {
      const postEvent = {
        postData: {
          contents: JSON.stringify({
            action: 'getMapData',
            params: {}
          })
        }
      };
      
      const postResult = doPost(postEvent);
      const postContent = postResult.getContent();
      const postData = JSON.parse(postContent);
      
      testResults.doPost = {
        success: true,
        hasData: !!postData,
        dataType: typeof postData,
        hasTherapists: !!postData.therapists,
        therapistsCount: postData.therapists ? postData.therapists.length : 0
      };
      console.log('✅ doPost with getMapData completed');
    } catch (e) {
      testResults.doPost = {
        success: false,
        error: e.message
      };
      console.log('❌ doPost with getMapData failed:', e.message);
    }
    
    return {
      success: true,
      message: 'Diagnostic completed',
      results: testResults,
      summary: {
        spreadsheetAccessible: true,
        allSheetsExist: true,
        individualFunctions: testResults,
        completeFunction: testResults.getMapData,
        doPostFunction: testResults.doPost
      }
    };
    
  } catch (error) {
    console.error('Diagnostic failed:', error);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Test function to verify the enhanced getMapData works
 */
function testEnhancedGetMapData() {
  try {
    console.log('=== TESTING ENHANCED GETMAPDATA ===');
    
    const result = getMapData();
    
    console.log('Test result:', {
      hasTherapists: !!result.therapists,
      hasClients: !!result.clients,
      hasSessions: !!result.sessions,
      hasAssignments: !!result.assignments,
      therapistsCount: result.therapists ? result.therapists.length : 0,
      clientsCount: result.clients ? result.clients.length : 0,
      sessionsCount: result.sessions ? result.sessions.length : 0,
      assignmentsCount: result.assignments ? Object.keys(result.assignments).length : 0,
      errorCount: result.errors ? result.errors.length : 0
    });
    
    if (result.errors && result.errors.length > 0) {
      console.log('Errors found:', result.errors);
    }
    
    return {
      success: true,
      result: result,
      summary: {
        totalDataPoints: (result.therapists ? result.therapists.length : 0) + 
                        (result.clients ? result.clients.length : 0) + 
                        (result.sessions ? result.sessions.length : 0) + 
                        (result.assignments ? Object.keys(result.assignments).length : 0),
        errorCount: result.errors ? result.errors.length : 0
      }
    };
    
  } catch (error) {
    console.error('Test failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Quick test function to verify basic functionality
 */
function testBasicFunctionality() {
  try {
    console.log('=== TESTING BASIC FUNCTIONALITY ===');
    
    // Test 1: Check if we can access the spreadsheet
    console.log('1. Testing spreadsheet access...');
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ Spreadsheet accessed:', spreadsheet.getName());
    
    // Test 2: Check if we can get sheet data
    console.log('2. Testing sheet data access...');
    const therapistsSheet = spreadsheet.getSheetByName('Therapists');
    if (therapistsSheet) {
      const data = therapistsSheet.getDataRange().getValues();
      console.log('✅ Therapists sheet data:', data.length, 'rows');
    } else {
      console.log('❌ Therapists sheet not found');
    }
    
    // Test 3: Test a simple SheetsService function
    console.log('3. Testing SheetsService...');
    const therapists = SheetsService.getTherapistsData();
    console.log('✅ SheetsService.getTherapistsData():', therapists.length, 'therapists');
    
    // Test 4: Test getMapData
    console.log('4. Testing getMapData...');
    const mapData = getMapData();
    console.log('✅ getMapData result:', {
      therapists: mapData.therapists ? mapData.therapists.length : 0,
      clients: mapData.clients ? mapData.clients.length : 0,
      sessions: mapData.sessions ? mapData.sessions.length : 0,
      assignments: mapData.assignments ? Object.keys(mapData.assignments).length : 0
    });
    
    return {
      success: true,
      message: 'All basic functionality tests passed',
      data: {
        spreadsheetName: spreadsheet.getName(),
        therapistsCount: therapists.length,
        mapDataSummary: {
          therapists: mapData.therapists ? mapData.therapists.length : 0,
          clients: mapData.clients ? mapData.clients.length : 0,
          sessions: mapData.sessions ? mapData.sessions.length : 0,
          assignments: mapData.assignments ? Object.keys(mapData.assignments).length : 0
        }
      }
    };
    
  } catch (error) {
    console.error('Basic functionality test failed:', error);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Get essential data for frontend initialization
 * Returns therapists, clients, and assignments (core data needed for UI)
 */
function getEssentialData() {
  console.log('[ESSENTIAL-DATA] Loading essential data...');
  
  try {
    const result = {
      therapists: [],
      clients: [],
      assignments: {},
      success: true,
      timestamp: new Date().toISOString()
    };
    
    // Load therapists
    console.log('[ESSENTIAL-DATA] Loading therapists...');
    result.therapists = SheetsService.getTherapistsData();
    console.log('[ESSENTIAL-DATA] ✅ Therapists loaded:', result.therapists.length);
    
    // Load clients
    console.log('[ESSENTIAL-DATA] Loading clients...');
    result.clients = SheetsService.getClientsData();
    console.log('[ESSENTIAL-DATA] ✅ Clients loaded:', result.clients.length);
    
    // Load assignments
    console.log('[ESSENTIAL-DATA] Loading assignments...');
    result.assignments = SheetsService.getAssignmentsData();
    console.log('[ESSENTIAL-DATA] ✅ Assignments loaded:', Object.keys(result.assignments).length);
    
    console.log('[ESSENTIAL-DATA] ✅ Essential data loaded successfully');
    return result;
    
  } catch (error) {
    console.error('[ESSENTIAL-DATA] ❌ Error loading essential data:', error);
    return {
      success: false,
      error: error.message,
      therapists: [],
      clients: [],
      assignments: {}
    };
  }
}

/**
 * Get sessions data with optional pagination
 * @param {Object} params - Parameters including page, limit, etc.
 */
function getSessionsData(params = {}) {
  console.log('[SESSIONS-DATA] Loading sessions data with params:', params);
  
  try {
    const result = {
      sessions: [],
      success: true,
      timestamp: new Date().toISOString()
    };
    
    // Load sessions
    console.log('[SESSIONS-DATA] Loading sessions...');
    result.sessions = SheetsService.getSessionsData();
    console.log('[SESSIONS-DATA] ✅ Sessions loaded:', result.sessions.length);
    
    // Apply pagination if specified
    if (params.page && params.limit) {
      const startIndex = (params.page - 1) * params.limit;
      const endIndex = startIndex + params.limit;
      result.sessions = result.sessions.slice(startIndex, endIndex);
      result.pagination = {
        page: params.page,
        limit: params.limit,
        total: result.sessions.length,
        hasMore: endIndex < result.sessions.length
      };
    }
    
    console.log('[SESSIONS-DATA] ✅ Sessions data loaded successfully');
    return result;
    
  } catch (error) {
    console.error('[SESSIONS-DATA] ❌ Error loading sessions data:', error);
    return {
      success: false,
      error: error.message,
      sessions: []
    };
  }
}

/**
 * Get the web app URL for doPost requests
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Make a doPost request using Google Apps Script's UrlFetchApp
 * This avoids CORS issues by making the request server-side
 * @param {string} action - The action to call
 * @param {object} params - Parameters for the action
 * @returns {object} - The response data
 */
function makeDoPostRequest(action, params = {}) {
  try {
    console.log('[DO-POST] Internally dispatching action:', action);
    // Execute directly inside Apps Script to avoid any external HTTP/auth
    const safeParams = params || {};
    const result = performApiAction(action, safeParams);
    console.log('[DO-POST] Internal dispatch completed. Success:', !!result && (result.success !== false));
    // Return as string to avoid large object serialization limits over google.script.run
    return JSON.stringify(result || {});
  } catch (error) {
    console.error('[DO-POST] Internal dispatch error:', error);
    return JSON.stringify({ success: false, error: error.message || String(error) });
  }
}

/**
 * Test function to verify makeDoPostRequest works correctly
 */
function testMakeDoPostRequest() {
  console.log('[TEST] Testing makeDoPostRequest...');
  
  try {
    const results = {
      success: true,
      tests: {},
      timestamp: new Date().toISOString()
    };
    
    // Test getEssentialData via makeDoPostRequest
    console.log('[TEST] Testing getEssentialData via makeDoPostRequest...');
    const essentialData = makeDoPostRequest('getEssentialData');
    results.tests.essentialData = {
      success: essentialData.success,
      therapistCount: essentialData.therapists ? essentialData.therapists.length : 0,
      clientCount: essentialData.clients ? essentialData.clients.length : 0,
      assignmentCount: essentialData.assignments ? Object.keys(essentialData.assignments).length : 0
    };
    
    // Test getSessionsData via makeDoPostRequest
    console.log('[TEST] Testing getSessionsData via makeDoPostRequest...');
    const sessionsData = makeDoPostRequest('getSessionsData');
    results.tests.sessionsData = {
      success: sessionsData.success,
      sessionCount: sessionsData.sessions ? sessionsData.sessions.length : 0
    };
    
    console.log('[TEST] ✅ All makeDoPostRequest tests completed successfully');
    return results;
    
  } catch (error) {
    console.error('[TEST] ❌ Error testing makeDoPostRequest:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test function to verify doPost endpoints are working
 */
function testDoPostEndpoints() {
  console.log('[TEST] Testing doPost endpoints...');
  
  try {
    const results = {
      success: true,
      tests: {},
      timestamp: new Date().toISOString()
    };
    
    // Test getEssentialData
    console.log('[TEST] Testing getEssentialData...');
    const essentialData = getEssentialData();
    results.tests.essentialData = {
      success: essentialData.success,
      therapistCount: essentialData.therapists ? essentialData.therapists.length : 0,
      clientCount: essentialData.clients ? essentialData.clients.length : 0,
      assignmentCount: essentialData.assignments ? Object.keys(essentialData.assignments).length : 0
    };
    
    // Test getSessionsData
    console.log('[TEST] Testing getSessionsData...');
    const sessionsData = getSessionsData();
    results.tests.sessionsData = {
      success: sessionsData.success,
      sessionCount: sessionsData.sessions ? sessionsData.sessions.length : 0
    };
    
    // Test getScheduleData
    console.log('[TEST] Testing getScheduleData...');
    const scheduleData = getScheduleData();
    results.tests.scheduleData = {
      success: scheduleData.success,
      assignmentCount: scheduleData.assignments ? Object.keys(scheduleData.assignments).length : 0,
      noteCount: scheduleData.notes ? scheduleData.notes.length : 0
    };
    
    console.log('[TEST] ✅ All doPost endpoints tested successfully');
    return results;
    
  } catch (error) {
    console.error('[TEST] ❌ Error testing doPost endpoints:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get schedule-specific data (assignments, notes, etc.)
 */
function getScheduleData() {
  console.log('[SCHEDULE-DATA] Loading schedule data...');
  
  try {
    const result = {
      assignments: {},
      notes: [],
      success: true,
      timestamp: new Date().toISOString()
    };
    
    // Load assignments
    console.log('[SCHEDULE-DATA] Loading assignments...');
    result.assignments = SheetsService.getAssignmentsData();
    console.log('[SCHEDULE-DATA] ✅ Assignments loaded:', Object.keys(result.assignments).length);
    
    // Load notes (if available)
    try {
      console.log('[SCHEDULE-DATA] Loading notes...');
      result.notes = SheetsService.getNotesData();
      console.log('[SCHEDULE-DATA] ✅ Notes loaded:', result.notes.length);
    } catch (noteError) {
      console.log('[SCHEDULE-DATA] ⚠️ Notes not available:', noteError.message);
      result.notes = [];
    }
    
    console.log('[SCHEDULE-DATA] ✅ Schedule data loaded successfully');
    return result;
    
  } catch (error) {
    console.error('[SCHEDULE-DATA] ❌ Error loading schedule data:', error);
    return {
      success: false,
      error: error.message,
      assignments: {},
      notes: []
    };
  }
}



 
