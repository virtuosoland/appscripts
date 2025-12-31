/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a custom menu in the spreadsheet UI when the file is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('FUB Formatter')
    .addItem('1. Convert Realtor Raw List', 'processRealtorList')
    .addSeparator() 
    .addItem('2. Convert Neighbor Raw List', 'processNeighborList')
    .addSeparator()
    .addItem('3. Convert Propwire Export', 'processPropwireExport')
    .addToUi();
}

/**
 * Helper function to get the Active Campaign Info from Row 2 of the "Property Data" sheet.
 * If data is incomplete, it will prompt the user to fill it, save it, and then proceed.
 * @returns {object|null} An object with property info or null if no active property is found or user cancels.
 */
function getAndConfirmCampaignInfo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Property Data");

  if (!dataSheet) {
    ui.alert('CRITICAL ERROR: A sheet named "Property Data" was not found. Please create this sheet first.');
    return null;
  }

  // Define the column indices based on the required header order
  const col = {
    key: 0, // A
    campaignTag: 1, // B
    propAddress: 2, // C
    propAPN: 3,     // D
    propCounty: 4,  // E
    propState: 5,   // F
    propAcreage: 6, // G
    propPrice: 7    // H
  };

  let activeRowData;
  try {
    activeRowData = dataSheet.getRange(2, 1, 1, 8).getValues()[0];
  } catch (e) {
    ui.alert('Error reading "Property Data" sheet. Please ensure it has at least 2 rows.');
    return null;
  }
  
  let campaignInfo = {};
  const streetAddressKey = activeRowData[col.key];
  const campaignTag = activeRowData[col.campaignTag];

  // --- NEW LOGIC: Check if data is incomplete ---
  if (!streetAddressKey || !campaignTag) {
    // 3B. IF INCOMPLETE OR BLANK: Prompt user to enter the data once
    ui.alert('No active property found (or data in Row 2 is incomplete).\n\nPlease enter the property details. This will be saved to Row 2 for future use.');
    
    const getResponse = (title, prompt) => {
      const resp = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() != ui.Button.OK) throw new Error('User canceled');
      return resp.getResponseText();
    };
    
    try {
      const fullAddress = getResponse('Step 1 of 7: Full Property Address', 'Enter the full property address (e.g., 123 Main St, Winston Salem, NC 27105):');
      const parsedKey = fullAddress.split(',')[0].trim();
      
      campaignInfo = {
        streetAddressKey: parsedKey,
        campaignTag:      `Campaign: ${parsedKey}`,
        propAddress:      fullAddress,
        propAPN:          getResponse('Step 2 of 7: Property APN', 'Enter the property APN (e.g., 123-45-678):'),
        propCounty:       getResponse('Step 3 of 7: Property County', 'Enter the property county (e.g., Forsyth):'),
        propState:        getResponse('Step 4 of 7: Property State', 'Enter the property state (e.g., NC):'),
        propAcreage:      getResponse('Step 5 of 7: Property Acreage', 'Enter the property acreage (e.g., 1.5):'),
        propPrice:        getResponse('Step 6 of 7: Asking Price', 'Enter the asking price (e.g., $50,000):')
      };
      
      // Save this new data back to Row 2 of the Property Data sheet
      const newRowData = [
        campaignInfo.streetAddressKey, campaignInfo.campaignTag, campaignInfo.propAddress,
        campaignInfo.propAPN, campaignInfo.propCounty, campaignInfo.propState,
        campaignInfo.propAcreage, campaignInfo.propPrice
      ];
      dataSheet.getRange(2, 1, 1, 8).setValues([newRowData]);
      ui.alert('New property data has been saved to Row 2 of the "Property Data" sheet.');
      
    } catch (e) {
      ui.alert('Script canceled during data entry.');
      return null;
    }
    
  } else {
    // 3A. IF FOUND: Map and confirm the data with the user
    campaignInfo = {
      streetAddressKey: activeRowData[col.key],
      campaignTag:      activeRowData[col.campaignTag].startsWith('Campaign: ') ? activeRowData[col.campaignTag] : `Campaign: ${activeRowData[col.campaignTag]}`,
      propAddress:      activeRowData[col.propAddress],
      propAPN:          activeRowData[col.propAPN],
      propCounty:       activeRowData[col.propCounty],
      propState:        activeRowData[col.propState],
      propAcreage:      activeRowData[col.propAcreage],
      propPrice:        activeRowData[col.propPrice]
    };
  }
  
  // --- The Confirmation Step ---
  // This now runs for BOTH new and found data, ensuring the user always sees what is being processed.
  const confirmationMessage = `You are about to process a list for the ACTIVE property:\n\n` +
                              `Campaign: ${campaignInfo.campaignTag}\n` +
                              `Address: ${campaignInfo.propAddress}\n` +
                              `APN: ${campaignInfo.propAPN}\n` +
                              `Price: ${campaignInfo.propPrice}\n\n` +
                              `Is this correct?`;
                                
  const userResponse = ui.alert(confirmationMessage, ui.ButtonSet.OK_CANCEL);
  
  if (userResponse != ui.Button.OK) {
    ui.alert('Script canceled. Please update the "Property Data" sheet in Row 2 and try again.');
    return null;
  }

  return campaignInfo;
}


/**
 * These are the headers for the FINAL FUB-ready template.
 */
const FUB_HEADERS = [
  'First Name', 'Last Name', 'Company Name', 'Email', 
  'Phone 1', 'Phone 2', 'Phone 3', 
  'Mailing Street', 'Mailing City', 'Mailing State', 'Mailing Zip', 
  'Tags', 
  'Owned Properties', 'Realtor - Recently Sold',
  '[DISP] Property Address', 
  '[DISP] Property APN', 
  '[DISP] Property County', 
  '[DISP] Property State', 
  '[DISP] Property Acreage', 
  '[DISP] Asking Price'
];

// =======================================================================================
// === 1. SCRIPT FOR REALTOR LIST ========================================================
// =======================================================================================
function processRealtorList() {
  const campaignInfo = getAndConfirmCampaignInfo();
  if (!campaignInfo) return; // Error or cancel was handled in the helper function

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Realtors-RawList');
  const targetSheetName = 'Realtor Import (Ready to Download)';
  const ui = SpreadsheetApp.getUi();

  if (!sourceSheet) {
    ui.alert('Error: A sheet named "Realtors-RawList" was not found.'); return;
  }
  
  ui.alert(`Processing Realtor list for campaign: ${campaignInfo.campaignTag}. This may take a moment...`);
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) targetSheet.clear();
  else targetSheet = ss.insertSheet(targetSheetName);

  const sourceData = sourceSheet.getRange(3, 1, sourceSheet.getLastRow() - 2, sourceSheet.getLastColumn()).getValues();
  const headers = sourceSheet.getRange(2, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, i) => { headerMap[header.trim()] = i; });
  
  const realtors = new Map();

  sourceData.forEach(row => {
    const agentNameString = row[headerMap["Agent's Name"]];
    if (!agentNameString || agentNameString.trim() === 'Public Records') return;

    if (!realtors.has(agentNameString)) {
      const nameParts = agentNameString.split('•');
      const fullName = nameParts[0].trim();
      const companyName = nameParts.length > 1 ? nameParts[1].trim() : '';
      const nameArray = fullName.split(' ');
      const firstName = nameArray.shift() || '';
      const lastName = nameArray.join(' ') || '';
      const state = row[headerMap['STATE OR PROVINCE']];
      
      const newRealtor = {
        firstName, lastName, companyName,
        email: row[headerMap['Email Address']],
        phone: row[headerMap['Mobile Phone Number']],
        tags: new Set([campaignInfo.campaignTag, 'Type: Realtor', `County: ${campaignInfo.propCounty}`]),
        recentlySold: []
      };
      if (state) newRealtor.tags.add(`State: ${state}`);
      realtors.set(agentNameString, newRealtor);
    }

    const realtor = realtors.get(agentNameString);
    const address = row[headerMap['ADDRESS']];
    const city = row[headerMap['CITY']];
    const state = row[headerMap['STATE OR PROVINCE']];
    const zip = row[headerMap['ZIP OR POSTAL CODE']];
    
    if (address && city && state && zip) {
      realtor.recentlySold.push(`${address}, ${city}, ${state} ${zip}`);
    }
  });

  targetSheet.getRange(1, 1, 1, FUB_HEADERS.length).setValues([FUB_HEADERS]).setFontWeight('bold');

  const outputData = Array.from(realtors.values()).map(r => [
    r.firstName, r.lastName, r.companyName, r.email, r.phone, '', '', '', '', '', '',
    Array.from(r.tags).join(','), '', r.recentlySold.join('\n'),
    campaignInfo.propAddress, campaignInfo.propAPN, campaignInfo.propCounty,
    campaignInfo.propState, campaignInfo.propAcreage, campaignInfo.propPrice
  ]);

  if (outputData.length > 0) {
    targetSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  }
  
  targetSheet.autoResizeColumns(1, FUB_HEADERS.length);
  ui.alert('Realtor list processing complete!');
}

// =======================================================================================
// === 2. SCRIPT FOR NEIGHBOR LIST =======================================================
// =======================================================================================
function processNeighborList() {
  const campaignInfo = getAndConfirmCampaignInfo();
  if (!campaignInfo) return; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Neighbors-RawList');
  const targetSheetName = 'Neighbor Import (Ready to Download)';
  const ui = SpreadsheetApp.getUi();

  if (!sourceSheet) {
    ui.alert('Error: A sheet named "Neighbors-RawList" was not found.'); return;
  }
  
  ui.alert(`Processing Neighbor list for campaign: ${campaignInfo.campaignTag}. This may take a moment...`);
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) targetSheet.clear();
  else targetSheet = ss.insertSheet(targetSheetName);

  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData.shift();
  const headerMap = {};
  headers.forEach((header, i) => { headerMap[header.trim()] = i; });

  const neighbors = new Map();

  sourceData.forEach(row => {
    const nameOrCompany = row[headerMap['Company Name']] || row[headerMap['Name']];
    if (!nameOrCompany) return;
    const uniqueKey = nameOrCompany.trim();
    
    if (!neighbors.has(uniqueKey)) {
      const companyName = row[headerMap['Company Name']];
      const isCompany = !!companyName;
      let firstName = '', lastName = '';
      if (!isCompany) {
        const nameArray = (row[headerMap['Name']] || '').trim().split(' ');
        firstName = nameArray.shift() || '';
        lastName = nameArray.join(' ') || '';
      }

      let street = '', city = '', state = '', zip = '';
      const mailingAddress = row[headerMap['Mailing Address']];
      if (mailingAddress) {
        const addressParts = mailingAddress.split(',');
        if (addressParts.length >= 3) {
          street = addressParts[0].trim();
          city = addressParts[1].trim();
          const stateZipPart = addressParts[2].trim().split(' ');
          state = stateZipPart.shift() || '';
          zip = stateZipPart.join(' ') || '';
        }
      }

      const newNeighbor = {
        firstName, lastName, companyName,
        email: row[headerMap['Email']],
        phone1: row[headerMap['Phone 1']],
        phone2: row[headerMap['Phone 2']],
        mailingStreet: street, mailingCity: city, mailingState: state, mailingZip: zip,
        ownedProperty: row[headerMap['Property Address']],
        tags: new Set([campaignInfo.campaignTag, 'Type: Neighbor', `County: ${campaignInfo.propCounty}`])
      };
      
      if (state) newNeighbor.tags.add(`State: ${state}`);
      if (isCompany) newNeighbor.tags.add('Type: Company');
      
      neighbors.set(uniqueKey, newNeighbor);
    }
  });

  targetSheet.getRange(1, 1, 1, FUB_HEADERS.length).setValues([FUB_HEADERS]).setFontWeight('bold');

  const outputData = Array.from(neighbors.values()).map(n => [
    n.firstName, n.lastName, n.companyName, n.email, n.phone1, n.phone2, '', n.mailingStreet, n.mailingCity, n.mailingState, n.mailingZip, Array.from(n.tags).join(','), n.ownedProperty, '',
    campaignInfo.propAddress, campaignInfo.propAPN, campaignInfo.propCounty,
    campaignInfo.propState, campaignInfo.propAcreage, campaignInfo.propPrice
  ]);

  if (outputData.length > 0) {
    targetSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  }
  
  targetSheet.autoResizeColumns(1, FUB_HEADERS.length);
  ui.alert('Neighbor list processing complete!');
}

// =======================================================================================
// === 3. SCRIPT FOR PROPWIRE LIST (INVESTORS) ===========================================
// =======================================================================================
function processPropwireExport() {
  const campaignInfo = getAndConfirmCampaignInfo();
  if (!campaignInfo) return; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Propwire-Investors-RawList');
  const targetSheetName = 'FUB Import (Ready to Download)';
  const ui = SpreadsheetApp.getUi();

  if (!sourceSheet) {
    ui.alert('Error: A sheet named "Propwire-Investors-RawList" was not found.'); return;
  }
  
  ui.alert(`Processing Propwire export for campaign: ${campaignInfo.campaignTag}. This may take a moment...`);
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) targetSheet.clear();
  else targetSheet = ss.insertSheet(targetSheetName);

  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData.shift();
  const headerMap = {};
  headers.forEach((header, i) => { headerMap[header.trim()] = i; });
  
  const contacts = new Map();

  sourceData.forEach(row => {
    const email = row[headerMap['Email']];
    if (!email && !row[headerMap['Phone 1']]) return;
    const uniqueKey = email || row[headerMap['Phone 1']];

    if (!contacts.has(uniqueKey)) {
      const isCompany = row[headerMap['Owner Type']] === 'COMPANY';
      let firstName = row[headerMap['Owner 1 First Name']];
      let lastName = row[headerMap['Owner 1 Last Name']];
      let companyName = '';
      if (isCompany) {
        companyName = `${firstName || ''} ${lastName || ''}`.trim();
        firstName = '';
        lastName = '';
      }
      
      const mailingState = row[headerMap['Owner Mailing State']];
      const county = row[headerMap['County']];

      const newContact = {
        firstName, lastName, companyName, email,
        phone1: row[headerMap['Phone 1']],
        phone2: row[headerMap['Phone 2']],
        phone3: row[headerMap['Phone 3']],
        mailingStreet: row[headerMap['Owner Mailing Address']],
        mailingCity: row[headerMap['Owner Mailing City']],
        mailingState: mailingState,
        mailingZip: row[headerMap['Owner Mailing Zip']],
        tags: new Set([campaignInfo.campaignTag,'Type: Investor', 'Source: Propwire', `County: ${campaignInfo.propCounty}`]),
        ownedProperties: []
      };
      
      if (mailingState) newContact.tags.add(`State: ${mailingState}`);
      if (isCompany) newContact.tags.add('Type: Company');
      if (county) newContact.tags.add(`County: ${county}`);
      
      contacts.set(uniqueKey, newContact);
    }

    const contact = contacts.get(uniqueKey);
    const propAddress = row[headerMap['Address']];
    const propCity = row[headerMap['City']];
    const propState = row[headerMap['State']];
    const propZip = row[headerMap['Zip']];
    
    if (propAddress) {
      contact.ownedProperties.push(`${propAddress}, ${propCity}, ${propState} ${propZip}`);
    }
  });

  targetSheet.getRange(1, 1, 1, FUB_HEADERS.length).setValues([FUB_HEADERS]).setFontWeight('bold');

  const outputData = Array.from(contacts.values()).map(c => [
    c.firstName, c.lastName, c.companyName, c.email, c.phone1, c.phone2, c.phone3, c.mailingStreet, c.mailingCity, c.mailingState, c.mailingZip, Array.from(c.tags).join(','), c.ownedProperties.join('\n'), '',
    campaignInfo.propAddress, campaignInfo.propAPN, campaignInfo.propCounty,
    campaignInfo.propState, campaignInfo.propAcreage, campaignInfo.propPrice
  ]);

  if (outputData.length > 0) {
    targetSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
  }
  
  targetSheet.autoResizeColumns(1, FUB_HEADERS.length);
  ui.alert('Processing Complete! Your clean data is ready in the "FUB Import (Ready to Download)" sheet.');
}