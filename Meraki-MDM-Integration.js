/**
 * Google Apps Script: Meraki MDM Tag Provisioning
 * 
 * Automates the assignment of Meraki MDM tags to iPads based on approvals tracked in a Google Sheet.
 * Handles form submissions, maps requested apps to Meraki tags, and performs the API call to update the device.
 */

const API_KEY = '<YOUR_MERAKI_API_KEY>';
const NETWORK_ID = '<YOUR_MERAKI_NETWORK_ID>';
const SHEET_NAME = 'Form Responses';

function onOpen(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const dataRange = sheet.getRange(2, 1, lastRow - 1, 10);
    const values = dataRange.getValues();
    let changesMade = false;
   
    for (let i = 0; i < values.length; i++) {
      const colA = values[i][0];
      const colB = values[i][1];
      const colJ = values[i][9]; // Approved column

      if (colA && colB && colJ === '') {
        values[i][9] = false; 
        changesMade = true;
      }
    }

    if (changesMade) {
      dataRange.setValues(values);
    }
  } catch (err) {
    Logger.log(`Error in onOpen trigger: ${err.message}`);
  }
}

function handleEdit(e) {
  try {
    if (!e || !e.source || !e.range) return;

    const sheet = e.source.getSheetByName(SHEET_NAME);
    const range = e.range;
    const editedRow = range.getRow();
    const editedCol = range.getColumn();

    // Only act if the "Approved" column is edited
    if (editedCol !== 10 || editedRow === 1) return;

    const serial = sheet.getRange(editedRow, 6).getValue(); 
    const appNameString = sheet.getRange(editedRow, 9).getValue(); 
    const approved = sheet.getRange(editedRow, 10).getValue(); 

    if (!serial || !appNameString) return;

    const tags = mapAppToTags(appNameString);
    if (tags.length === 0) return;

    const action = approved === true ? 'add' : 'delete';

    for (const tag of tags) {
      modifyTag(serial, tag, action, e);
    }
  } catch (err) {
    Logger.log(`Unexpected error in handleEdit: ${err.message}`);
  }
}

function mapAppToTags(appNameString) {
  const foundTags = [];
  const normalized = appNameString.toLowerCase().trim();

  if (normalized.includes('lamp')) foundTags.push('lamp');
  if (normalized.includes('td snap')) foundTags.push('tdsnap');
  if (normalized.includes('proloquo')) foundTags.push('proloquo');
 
  return foundTags;
}

function modifyTag(serial, tag, action, e) {
  const url = `https://api.meraki.com/api/v1/networks/${NETWORK_ID}/sm/devices/modifyTags`;
  const payload = {
    serials: [serial],
    tags: [tag],
    updateAction: action
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Cisco-Meraki-API-Key': API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const ss = e.source;
    const sheet = ss.getSheetByName(SHEET_NAME);
    const editedRow = e.range.getRow();

    if (statusCode >= 200 && statusCode < 300) {
      if (action === 'add') {
        sheet.getRange(editedRow, 11).setValue("Completed");
        ss.toast(`Successfully added tag '${tag}' for ${serial}`, "Tag Added", 5);
      } else {
        sheet.getRange(editedRow, 11).setValue("Revoked");
        ss.toast(`Successfully removed tag '${tag}' for ${serial}`, "Tag Revoked", 5);
      }
    } else {
      ss.toast(`API Error: ${statusCode}.`, "Meraki Error", 10);
    }
  } catch (err) {
    e.source.toast("Connection Error: " + err.message, "Critical Error", 10);
  }
}
