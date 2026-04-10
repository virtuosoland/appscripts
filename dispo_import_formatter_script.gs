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
 */
function getAndConfirmCampaignInfo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Property Data");

  if (!dataSheet) {
    ui.alert('CRITICAL ERROR: A sheet named "Property Data" was not found. Please create this sheet first.');
    return null;
  }

  const col = {
    key: 0, campaignTag: 1, propAddress: 2, propAPN: 3, propCounty: 4, 
    propState: 5, propAcreage: 6, propPrice: 7, pebbleURL: 8
  };

  let activeRowData;
  try {
    activeRowData = dataSheet.getRange(2, 1, 1, 9).getValues()[0];
  } catch (e) {
    ui.alert('Error reading "Property Data" sheet. Please ensure it has at least 2 rows.');
    return null;
  }
  
  let campaignInfo = {};
  const streetAddressKey = activeRowData[col.key];
  const campaignTag = activeRowData[col.campaignTag];

  if (!streetAddressKey || !campaignTag) {
    ui.alert('No active property found in Row 2 of "Property Data".\n\nPlease fill out the details.');
    
    const getResponse = (title, prompt) => {
      const resp = ui.prompt(title, prompt, ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() != ui.Button.OK) throw new Error('User canceled');
      return resp.getResponseText();
    };
    
    try {
      const fullAddress = getResponse('Step 1: Address', 'Full property address:');
      const parsedKey = fullAddress.split(',')[0].trim();
      
      campaignInfo = {
        streetAddressKey: parsedKey,
        campaignTag:      `Campaign: ${parsedKey}`,
        propAddress:      fullAddress,
        propAPN:          getResponse('Step 2: APN', 'Property APN:'),
        propCounty:       getResponse('Step 3: County', 'County:'),
        propState:        getResponse('Step 4: State', 'State:'),
        propAcreage:      getResponse('Step 5: Acreage', 'Acreage:'),
        propPrice:        getResponse('Step 6: Price', 'Asking Price:'),
        pebbleURL:        getResponse('Step 7: URL', 'Pebble URL (optional):')
      };
      
      const newRowData = [
        campaignInfo.streetAddressKey, campaignInfo.campaignTag, campaignInfo.propAddress,
        campaignInfo.propAPN, campaignInfo.propCounty, campaignInfo.propState,
        campaignInfo.propAcreage, campaignInfo.propPrice, campaignInfo.pebbleURL
      ];
      dataSheet.getRange(2, 1, 1, 9).setValues([newRowData]);
    } catch (e) {
      ui.alert('Script canceled.');
      return null;
    }
  } else {
    campaignInfo = {
      streetAddressKey: activeRowData[col.key],
      campaignTag:      activeRowData[col.campaignTag].toString().startsWith('Campaign: ') ? activeRowData[col.campaignTag] : `Campaign: ${activeRowData[col.campaignTag]}`,
      propAddress:      activeRowData[col.propAddress],
      propAPN:          activeRowData[col.propAPN],
      propCounty:       activeRowData[col.propCounty],
      propState:        activeRowData[col.propState],
      propAcreage:      activeRowData[col.propAcreage],
      propPrice:        activeRowData[col.propPrice],
      pebbleURL:        activeRowData[col.pebbleURL]
    };
  }
  
  const confirmationMessage = `Process list for:\n\nCampaign: ${campaignInfo.campaignTag}\nAddress: ${campaignInfo.propAddress}\n\nIs this correct?`;
  const userResponse = ui.alert(confirmationMessage, ui.ButtonSet.OK_CANCEL);
  
  return (userResponse == ui.Button.OK) ? campaignInfo : null;
}

const FUB_HEADERS = [
  'First Name', 'Last Name', 'Company Name', 'Email', 
  'Phone 1', 'Phone 2', 'Phone 3', 
  'Mailing Street', 'Mailing City', 'Mailing State', 'Mailing Zip', 
  'Tags', 'Owned Properties', 'Realtor - Recently Sold',
  '[DISP] Property Address', '[DISP] Property APN', '[DISP] Property County', 
  '[DISP] Property State', '[DISP] Property Acreage', '[DISP] Asking Price', 'Pebble Deal URL'
];

/**
 * 1. SCRIPT FOR REALTOR LIST
 * Dynamically finds headers and skips empty/title rows.
 */
function processRealtorList() {
  const campaignInfo = getAndConfirmCampaignInfo();
  if (!campaignInfo) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Realtors-RawList');
  const ui = SpreadsheetApp.getUi();

  if (!sourceSheet) {
    ui.alert('Error: "Realtors-RawList" sheet not found.'); return;
  }
  
  const outreachResponse = ui.alert('Outreach Type', 'Is this for BPO Outreach?\n\nYes = BPO\nNo = Buyer', ui.ButtonSet.YES_NO);
  const realtorTriggerTag = outreachResponse == ui.Button.YES ? '[DISP] Trigger: Start BPO Outreach' : '[DISP] Trigger: Start Realtors (Buyer)';
  
  const fullData = sourceSheet.getDataRange().getValues();
  let headerRowIndex = -1;
  
  // Dynamic header detection (scans first 10 rows)
  for (let i = 0; i < Math.min(fullData.length, 10); i++) {
    const rowStr = fullData[i].join("|");
    if (rowStr.includes("Agent's Name") || rowStr.includes("Mobile Phone Number")) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    ui.alert('Headers not found. Ensure "Agent\'s Name" exists in the headers.');
    return;
  }

  const headers = fullData[headerRowIndex];
  const map = {};
  headers.forEach((h, i) => { map[h.toString().trim()] = i; });
  
  const realtors = new Map();
  const dataRows = fullData.slice(headerRowIndex + 1);

  dataRows.forEach(row => {
    const agentName = (row[map["Agent's Name"]] || "").toString().trim();
    // Skip if empty, "Public Records", or placeholder text
    if (!agentName || agentName === '' || agentName === 'Public Records') return;

    if (!realtors.has(agentName)) {
      const nameParts = agentName.split('•');
      const fullName = nameParts[0].trim();
      const company = nameParts.length > 1 ? nameParts[1].trim() : '';
      const nameArray = fullName.split(' ');
      const fName = nameArray.shift() || '';
      const lName = nameArray.join(' ') || '';
      const state = row[map['STATE OR PROVINCE']];
      
      realtors.set(agentName, {
        firstName: fName, lastName: lName, companyName: company,
        email: (row[map['Email Address']] || "").toString().trim(),
        phone: (row[map['Mobile Phone Number']] || "").toString().trim(),
        tags: new Set([campaignInfo.campaignTag, 'Type: Realtor', `County: ${campaignInfo.propCounty}`, realtorTriggerTag]),
        recentlySold: []
      });
      if (state) realtors.get(agentName).tags.add(`State: ${state}`);
    }

    const r = realtors.get(agentName);
    const addr = row[map['ADDRESS']], city = row[map['CITY']], st = row[map['STATE OR PROVINCE']], zip = row[map['ZIP OR POSTAL CODE']];
    if (addr && city) {
      r.recentlySold.push(`${addr}, ${city}, ${st} ${zip}`);
    }
  });

  const output = Array.from(realtors.values()).map(r => [
    r.firstName, r.lastName, r.companyName, r.email, r.phone, '', '', '', '', '', '',
    Array.from(r.tags).join(','), '', r.recentlySold.join('\n'),
    campaignInfo.propAddress, campaignInfo.propAPN, campaignInfo.propCounty,
    campaignInfo.propState, campaignInfo.propAcreage, campaignInfo.propPrice,
    campaignInfo.pebbleURL
  ]);

  writeToTargetSheet(ss, 'Realtor Import (Ready to Download)', output);
  ui.alert('Realtor list complete!');
}

/**
 * 2. SCRIPT FOR NEIGHBOR LIST
 */
function processNeighborList() {
  const info = getAndConfirmCampaignInfo();
  if (!info) return; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Neighbors-RawList');
  if (!sheet) { SpreadsheetApp.getUi().alert('Error: "Neighbors-RawList" missing.'); return; }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const map = {};
  headers.forEach((h, i) => { map[h.toString().trim()] = i; });

  const neighbors = new Map();
  data.forEach(row => {
    const key = (row[map['Company Name']] || row[map['Name']] || "").toString().trim();
    if (!key || neighbors.has(key)) return;
    
    const isCo = !!row[map['Company Name']];
    let f = '', l = '';
    if (!isCo) {
      const arr = key.split(' ');
      f = arr.shift(); l = arr.join(' ');
    }

    let street = '', city = '', state = '', zip = '';
    const mAddr = (row[map['Mailing Address']] || "").toString();
    if (mAddr.includes(',')) {
      const p = mAddr.split(',');
      street = p[0].trim(); city = (p[1] || "").trim();
      const sz = (p[2] || "").trim().split(' ');
      state = sz.shift(); zip = sz.join(' ');
    }

    const tags = [info.campaignTag, 'Type: Neighbor', `County: ${info.propCounty}`, '[DISP] Trigger: Neighbor Outreach'];
    if (state) tags.push(`State: ${state}`);
    if (isCo) tags.push('Type: Company');
    
    neighbors.set(key, [
      f, l, isCo ? key : '', row[map['Email']] || '', row[map['Phone 1']] || '', row[map['Phone 2']] || '', '',
      street, city, state, zip, tags.join(','), row[map['Property Address']] || '', '',
      info.propAddress, info.propAPN, info.propCounty, info.propState, info.propAcreage, info.propPrice, info.pebbleURL
    ]);
  });

  writeToTargetSheet(ss, 'Neighbor Import (Ready to Download)', Array.from(neighbors.values()));
  SpreadsheetApp.getUi().alert('Neighbor list complete!');
}

/**
 * 3. SCRIPT FOR PROPWIRE LIST
 */
function processPropwireExport() {
  const info = getAndConfirmCampaignInfo();
  if (!info) return; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Propwire-Investors-RawList');
  if (!sheet) { SpreadsheetApp.getUi().alert('Error: "Propwire-Investors-RawList" missing.'); return; }
  
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const map = {};
  headers.forEach((h, i) => { map[h.toString().trim()] = i; });
  
  const contacts = new Map();
  data.forEach(row => {
    const email = row[map['Email']], phone = row[map['Phone 1']];
    const key = email || phone;
    if (!key) return;

    if (!contacts.has(key)) {
      const isCo = row[map['Owner Type']] === 'COMPANY';
      const f = row[map['Owner 1 First Name']] || '', l = row[map['Owner 1 Last Name']] || '';
      const tags = [info.campaignTag, 'Type: Investor', 'Source: Propwire', `County: ${info.propCounty}`, '[DISP] Trigger: Investor Outreach'];
      if (isCo) tags.push('Type: Company');
      if (row[map['Owner Mailing State']]) tags.push(`State: ${row[map['Owner Mailing State']]}`);

      contacts.set(key, {
        f: isCo ? '' : f, l: isCo ? '' : l, co: isCo ? `${f} ${l}`.trim() : '',
        email: email || '', p1: phone || '', p2: row[map['Phone 2']] || '', p3: row[map['Phone 3']] || '',
        ms: row[map['Owner Mailing Address']], mc: row[map['Owner Mailing City']], mst: row[map['Owner Mailing State']], mz: row[map['Owner Mailing Zip']],
        tags: tags, owned: []
      });
    }
    const c = contacts.get(key);
    if (row[map['Address']]) c.owned.push(`${row[map['Address']]}, ${row[map['City']]}, ${row[map['State']]}`);
  });

  const output = Array.from(contacts.values()).map(c => [
    c.f, c.l, c.co, c.email, c.p1, c.p2, c.p3, c.ms, c.mc, c.mst, c.mz, c.tags.join(','), c.owned.join('\n'), '',
    info.propAddress, info.propAPN, info.propCounty, info.propState, info.propAcreage, info.propPrice, info.pebbleURL
  ]);

  writeToTargetSheet(ss, 'FUB Import (Ready to Download)', output);
  SpreadsheetApp.getUi().alert('Propwire complete!');
}

/**
 * Shared Helper: Writes data to a target sheet.
 */
function writeToTargetSheet(ss, sheetName, data) {
  let ts = ss.getSheetByName(sheetName);
  if (ts) ts.clear(); else ts = ss.insertSheet(sheetName);
  ts.getRange(1, 1, 1, FUB_HEADERS.length).setValues([FUB_HEADERS]).setFontWeight('bold');
  if (data.length > 0) ts.getRange(2, 1, data.length, data[0].length).setValues(data);
  ts.autoResizeColumns(1, FUB_HEADERS.length);
}