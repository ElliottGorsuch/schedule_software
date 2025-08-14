/**
 * GeocodeService.gs - Geocoding service using Google Maps API
 * Replaces the Wix geocode.jsw functionality
 * Public release hardening: sensitive logs gated by DEBUG_MODE; API keys read from Script Properties.
 * Updated to use static methods for easier access
 */

// Lightweight debug logger gated by DEBUG_MODE script property or global
function geoDebugLog() {
  var isDebug = false;
  try {
    if (typeof DEBUG_MODE !== 'undefined') {
      isDebug = !!DEBUG_MODE;
    } else {
      var prop = PropertiesService.getScriptProperties().getProperty('DEBUG_MODE');
      isDebug = prop === 'true';
    }
  } catch (e) {
    isDebug = false;
  }
  if (isDebug && typeof console !== 'undefined' && console && console.log) {
    // eslint-disable-next-line prefer-rest-params
    console.log.apply(console, arguments);
  }
}

class GeocodeService {
  /**
   * Get API key from script properties
   * To set this up: go to Project Settings > Script Properties and add:
   * Key: GOOGLE_MAPS_API_KEY
   * Value: your_api_key_here
   */
  static getApiKey() {
    try {
      return PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');
    } catch (error) {
      console.error('Error getting API key:', error);
      return null;
    }
  }

  /**
   * Geocode an address using Google Maps Geocoding API
   * @param {string} address - The address to geocode
   * @param {string|number} zipCode - Optional ZIP code for validation
   * @returns {Object} Coordinates object with lat and lng properties
   */
  static geocodeAddress(address, zipCode = null) {
    geoDebugLog('Geocoding address request received');
    
    try {
      const apiKey = GeocodeService.getApiKey();
      
      if (!apiKey) {
        throw new Error('Google Maps API key not configured. Please set GOOGLE_MAPS_API_KEY in Script Properties.');
      }
      
      // Validate address format
      const validation = GeocodeService.validateAddressFormat(address);
      if (!validation.isValid) {
        throw new Error(`Invalid address format: ${validation.reason}`);
      }
      
      // Standardize the address format
      const standardizedAddress = GeocodeService.standardizeAddress(address);
      geoDebugLog('Standardized address prepared');
      
      // URL encode the address
      const encodedAddress = encodeURIComponent(standardizedAddress);
      
      // Build the API URL
      const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodedAddress}&key=${apiKey}`;
      
      geoDebugLog(`Making geocoding request to: ${url.replace(apiKey, 'API_KEY_HIDDEN')}`);
      
      // Make the API request
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());
      
      geoDebugLog(`Geocoding API response status: ${data.status}`);
      
      if (data.status !== 'OK') {
        let errorMessage = `Geocoding API error: ${data.status}`;
        
        // Provide more helpful error messages based on status
        switch(data.status) {
          case 'ZERO_RESULTS':
            errorMessage += ' - No results found. Try adding city, state, or ZIP code to the address.';
            break;
          case 'OVER_QUERY_LIMIT':
            errorMessage += ' - API quota exceeded. Please try again later.';
            break;
          case 'REQUEST_DENIED':
            errorMessage += ' - API request denied. Check your API key permissions.';
            break;
          case 'INVALID_REQUEST':
            errorMessage += ' - Invalid request. Check the address format.';
            break;
          default:
            errorMessage += ` - ${data.error_message || 'Unknown error'}`;
        }
        
        // Add suggestion for ZERO_RESULTS
        if (data.status === 'ZERO_RESULTS') {
          errorMessage += `\n\nSuggestions for "${address}":\n`;
          errorMessage += `• Try: "${address}, [City], [State]"\n`;
          errorMessage += `• Try: "${address}, [ZIP Code]"\n`;
          errorMessage += `• Verify the street name and number are correct`;
        }
        
        throw new Error(errorMessage);
      }
      
      if (!data.results || data.results.length === 0) {
        throw new Error(`No results found for address: "${address}"`);
      }
      
      // Extract coordinates from the result
      const result = data.results[0];
      const location = result.geometry.location;
      
      const coordinates = {
        lat: location.lat,
        lng: location.lng,
        formatted_address: result.formatted_address
      };
      
      geoDebugLog('Geocoding successful');
      
      // Validate coordinates against ZIP code if provided
      if (zipCode && !GeocodeService.isCoordinateInZipRegion(coordinates, zipCode)) {
        geoDebugLog('ZIP region validation warning for coordinates');
        // Don't throw error, just warn - Google Maps is usually accurate
      }
      
      return coordinates;
      
    } catch (error) {
      console.error('Geocoding error:', error);
      throw new Error(`Failed to geocode address "${address}": ${error.message}`);
    }
  }
  
  /**
   * Standardize address format by converting full words to abbreviations
   */
  static standardizeAddress(address) {
    if (!address) return address;
    
    const streetTypes = {
      "street": "St", 
      "avenue": "Ave",
      "boulevard": "Blvd",
      "drive": "Dr",
      "road": "Rd",
      "place": "Pl",
      "lane": "Ln",
      "court": "Ct",
      "circle": "Cir",
      "highway": "Hwy",
      "parkway": "Pkwy"
    };
    
    let standardized = address;
    
    // Replace full words with abbreviations (case insensitive)
    for (const [fullWord, abbrev] of Object.entries(streetTypes)) {
      const regex = new RegExp(`\\b${fullWord}\\b`, 'i');
      standardized = standardized.replace(regex, abbrev);
    }
    
    return standardized;
  }
  
  /**
   * Basic validation of address format
   */
  static validateAddressFormat(address) {
    if (!address) {
      return { isValid: false, reason: "Address is empty" };
    }
    
    // Check if address has a street number (starts with numbers)
    const hasStreetNumber = /^\d+\s+/.test(address);
    
    // Check if address has a street name
    const hasStreetName = /\d+\s+[A-Za-z]+/.test(address);
    
    // Require at least a basic "123 Main" format
    if (!hasStreetNumber || !hasStreetName) {
      return { 
        isValid: false, 
        reason: "Address should start with a street number followed by street name (e.g., '123 Main St')" 
      };
    }
    
    return { isValid: true };
  }
  
  /**
   * Check if coordinates are within the expected region for a ZIP code
   * This is a simplified validation - Google Maps API is usually accurate
   */
  static isCoordinateInZipRegion(coordinates, zipCode) {
    if (!coordinates || !zipCode) return true; // Skip validation if data is missing
    
    const zip = String(zipCode);
    
    // Simple ZIP region validation for US ZIP codes
    const zipRegions = {
      '0': { bounds: {minLat: 40, maxLat: 47, minLng: -76, maxLng: -67} }, // Northeast
      '1': { bounds: {minLat: 39, maxLat: 45, minLng: -80, maxLng: -71} }, // Northeast
      '2': { bounds: {minLat: 34, maxLat: 42, minLng: -83, maxLng: -74} }, // Mid-Atlantic
      '3': { bounds: {minLat: 24.5, maxLat: 36, minLng: -88, maxLng: -75} }, // Southeast
      '4': { bounds: {minLat: 36, maxLat: 43, minLng: -90, maxLng: -80} }, // Midwest
      '5': { bounds: {minLat: 36, maxLat: 49, minLng: -97, maxLng: -84} }, // Midwest
      '6': { bounds: {minLat: 25, maxLat: 40, minLng: -102, maxLng: -89} }, // South/Central
      '7': { bounds: {minLat: 32, maxLat: 49, minLng: -104, maxLng: -94} }, // Central
      '8': { bounds: {minLat: 31, maxLat: 49, minLng: -116, maxLng: -100} }, // Mountain
      '9': { bounds: {minLat: 32, maxLat: 49, minLng: -125, maxLng: -114} }  // West Coast
    };
    
    // Special case for Denver area (802xx ZIP codes)
    if (zip.startsWith('802')) {
      const denverBounds = {minLat: 39.5, maxLat: 40.0, minLng: -105.1, maxLng: -104.7};
      return coordinates.lat >= denverBounds.minLat && 
             coordinates.lat <= denverBounds.maxLat &&
             coordinates.lng >= denverBounds.minLng && 
             coordinates.lng <= denverBounds.maxLng;
    }
    
    // General region check using first digit
    const zipPrefix = zip.substring(0, 1);
    if (zipRegions[zipPrefix]) {
      const bounds = zipRegions[zipPrefix].bounds;
      return coordinates.lat >= bounds.minLat && 
             coordinates.lat <= bounds.maxLat &&
             coordinates.lng >= bounds.minLng && 
             coordinates.lng <= bounds.maxLng;
    }
    
    return true; // If we can't validate, assume it's correct
  }

  /**
   * Batch geocode multiple addresses with rate limiting
   */
  static batchGeocode(addresses) {
    geoDebugLog(`Starting batch geocoding of ${addresses.length} addresses`);
    const results = [];
    
    // Process addresses one by one with rate limiting
    addresses.forEach((addressObj, index) => {
      const { id, address, zipCode } = addressObj;
      
      try {
        // Add delay to respect API rate limits
        if (index > 0) {
          Utilities.sleep(100); // 100ms delay between requests
        }
        
        const location = GeocodeService.geocodeAddress(address, zipCode);
        
        results.push({
          id: id,
          success: true,
          coordinates: location
        });
        
      } catch (error) {
        console.error('Failed to geocode address in batch:', error);
        results.push({
          id: id,
          success: false,
          error: error.message
        });
      }
    });
    
    geoDebugLog(`Completed batch geocoding: ${results.filter(r => r.success).length} successes, ${results.filter(r => !r.success).length} failures`);
    return results;
  }

  /**
   * Validate and standardize address
   * Replaces: geocode.validateAddress(address)
   */
  static validateAddress(address) {
    try {
      const geocodeResult = GeocodeService.geocodeAddress(address);
      
      return {
        isValid: true,
        originalAddress: address,
        standardizedAddress: geocodeResult.formatted_address,
        coordinates: {
          latitude: geocodeResult.lat,
          longitude: geocodeResult.lng
        },
        confidence: GeocodeService.getAddressConfidence(address, geocodeResult.formatted_address)
      };
    } catch (error) {
      return {
        isValid: false,
        originalAddress: address,
        error: error.message,
        errorDetails: 'No results found'
      };
    }
  }

  /**
   * Parse Google Maps address components into a usable format
   */
  static parseAddressComponents(components) {
    const parsed = {
      streetNumber: '',
      route: '',
      city: '',
      state: '',
      zipCode: '',
      country: ''
    };

    components.forEach(component => {
      const types = component.types;
      
      if (types.includes('street_number')) {
        parsed.streetNumber = component.long_name;
      } else if (types.includes('route')) {
        parsed.route = component.long_name;
      } else if (types.includes('locality')) {
        parsed.city = component.long_name;
      } else if (types.includes('administrative_area_level_1')) {
        parsed.state = component.short_name;
      } else if (types.includes('postal_code')) {
        parsed.zipCode = component.long_name;
      } else if (types.includes('country')) {
        parsed.country = component.long_name;
      }
    });

    return parsed;
  }

  /**
   * Get confidence score for address match
   */
  static getAddressConfidence(originalAddress, formattedAddress) {
    if (!originalAddress || !formattedAddress) {
      return 0;
    }

    const original = originalAddress.toLowerCase().trim();
    const formatted = formattedAddress.toLowerCase().trim();
    
    // Simple confidence calculation based on string similarity
    if (original === formatted) {
      return 1.0; // Perfect match
    }
    
    // Check if original address is contained in formatted address
    if (formatted.includes(original)) {
      return 0.8; // High confidence
    }
    
    // Check for partial matches
    const originalWords = original.split(/\s+/);
    const formattedWords = formatted.split(/\s+/);
    
    let matchingWords = 0;
    originalWords.forEach(word => {
      if (formattedWords.some(fWord => fWord.includes(word) || word.includes(fWord))) {
        matchingWords++;
      }
    });
    
    const wordMatchRatio = matchingWords / originalWords.length;
    
    if (wordMatchRatio >= 0.7) {
      return 0.7; // Good confidence
    } else if (wordMatchRatio >= 0.5) {
      return 0.5; // Medium confidence
    } else {
      return 0.3; // Low confidence
    }
  }

  /**
   * Validate ZIP code region (useful for service area validation)
   * You can customize this based on your service areas
   */
  static validateZipCodeRegion(zipCode, allowedRegions = []) {
    if (!zipCode) {
      return {
        isValid: false,
        error: 'ZIP code is required'
      };
    }

    // Remove any non-numeric characters
    const cleanZip = zipCode.replace(/\D/g, '');
    
    if (cleanZip.length < 5) {
      return {
        isValid: false,
        error: 'ZIP code must be at least 5 digits'
      };
    }

    // If no allowed regions specified, accept any valid ZIP
    if (allowedRegions.length === 0) {
      return {
        isValid: true,
        zipCode: cleanZip.substring(0, 5)
      };
    }

    // Check against allowed regions
    const zipPrefix = cleanZip.substring(0, 3);
    const isInAllowedRegion = allowedRegions.some(region => {
      if (typeof region === 'string') {
        return zipPrefix === region;
      } else if (typeof region === 'object' && region.start && region.end) {
        return zipPrefix >= region.start && zipPrefix <= region.end;
      }
      return false;
    });

    return {
      isValid: isInAllowedRegion,
      zipCode: cleanZip.substring(0, 5),
      error: isInAllowedRegion ? null : 'ZIP code is outside service area'
    };
  }

  /**
   * Geocode therapist data and update sheet
   * Integrates with SheetsService
   */
  static geocodeAndUpdateTherapist(therapistId, address) {
    try {
      const geocodeResult = GeocodeService.geocodeAddress(address);
      
      // Since we're making this static, we'll just return the geocode result
      // The calling function can handle the sheet update
      return {
        success: true,
        geocodeResult: geocodeResult,
        message: 'Therapist geocoded successfully'
      };
    } catch (error) {
      console.error('Error geocoding therapist:', error);
      return {
        success: false,
        error: error.message,
        message: 'Failed to geocode therapist address'
      };
    }
  }

  /**
   * Geocode client data and update sheet
   * Integrates with SheetsService
   */
  static geocodeAndUpdateClient(clientId, address) {
    try {
      const geocodeResult = GeocodeService.geocodeAddress(address);
      
      // Since we're making this static, we'll just return the geocode result
      // The calling function can handle the sheet update
      return {
        success: true,
        geocodeResult: geocodeResult,
        message: 'Client geocoded successfully'
      };
    } catch (error) {
      console.error('Error geocoding client:', error);
      return {
        success: false,
        error: error.message,
        message: 'Failed to geocode client address'
      };
    }
  }

  /**
   * Test geocoding service with a sample address
   */
  static testGeocoding(testAddress = "1600 Amphitheatre Parkway, Mountain View, CA") {
    try {
      console.log('Testing geocoding service...');
      const result = GeocodeService.geocodeAddress(testAddress);
      
      console.log('Geocoding test successful!');
      console.log('Address:', result.formatted_address);
      console.log('Coordinates:', result.lat, result.lng);
      return {
        success: true,
        result: result
      };
    } catch (error) {
      console.error('Geocoding test error:', error);
      return {
        success: false,
        error: error.message,
        errorDetails: error.toString()
      };
    }
  }
}

/**
 * Quick geocode function for single addresses - now uses static method
 */
function geocodeAddress(address) {
  return GeocodeService.geocodeAddress(address);
}

/**
 * Quick validate function for addresses - now uses static method
 */
function validateAddress(address) {
  return GeocodeService.validateAddress(address);
}

/**
 * Setup function to help configure the API key
 * Run this once to set your Google Maps API key
 */
function setupGoogleMapsApiKey() {
  const apiKey = Browser.inputBox(
    'Google Maps API Key Setup',
    'Please enter your Google Maps API key:',
    Browser.Buttons.OK_CANCEL
  );
  
  if (apiKey && apiKey !== 'cancel') {
    PropertiesService.getScriptProperties().setProperty('GOOGLE_MAPS_API_KEY', apiKey);
    Browser.msgBox('Success', 'Google Maps API key has been saved!', Browser.Buttons.OK);
    
    // Test the API key using static method
    const testResult = GeocodeService.testGeocoding();
    
    if (testResult.success) {
      Browser.msgBox('Test Successful', 'Your API key is working correctly!', Browser.Buttons.OK);
    } else {
      Browser.msgBox('Test Failed', `API key test failed: ${testResult.error}`, Browser.Buttons.OK);
    }
  }
}

/**
 * Test function to verify geocoding is working
 * Run this to test your geocoding setup
 */
function testGeocodingSetup() {
  console.log('=== GEOCODING SERVICE TEST ===');
  
  const testAddress = "1234 Main Street, Denver, CO 80202";
  console.log('Testing address:', testAddress);
  
  try {
    const result = GeocodeService.geocodeAddress(testAddress);
    console.log('✅ Geocoding successful:', result);
    return {
      success: true,
      result: result
    };
  } catch (error) {
    console.error('❌ Geocoding failed:', error);
    return {
      success: false,
      error: error.message
    };
  }
} 
