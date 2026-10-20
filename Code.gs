/**
 * SIMSALACLEAN — WAITLIST INTEGRATION
 * Google Apps Script Web App
 *
 * This script receives form submissions from the Simsalaclean landing page
 * and saves them to the connected Google Sheet.
 *
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet (or use an existing one).
 * 2. Go to Extensions → Apps Script.
 * 3. Delete any code in Code.gs, then paste ALL of this code.
 * 4. Save (Ctrl/Cmd + S).
 * 5. Select "setup" from the function dropdown and click Run (one-time setup).
 *    → Approve the permissions when prompted.
 * 6. Click Deploy → New deployment.
 * 7. Choose type: Web app.
 * 8. Settings:
 *    - Description: Simsalaclean Waitlist
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 9. Click Deploy, authorize if asked, then copy the Web App URL.
 * 10. Paste that URL into index.html → APPS_SCRIPT_URL constant.
 */

// ============================================================
//  SETUP — Run this once to create sheet headers
// ============================================================

function setup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Rename the sheet tab for clarity
  sheet.setName('Waitlist');

  // Clear and write headers
  sheet.clear();
  sheet.appendRow(['Timestamp', 'Nama', 'Username Instagram']);

  // Style the header row
  var headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#F26522'); // Simsalaclean orange
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('left');

  // Column widths
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 240); // Nama
  sheet.setColumnWidth(3, 200); // Instagram

  // Freeze header row
  sheet.setFrozenRows(1);

  Logger.log('✅ Setup complete! Sheet headers added.');
}

// ============================================================
//  POST HANDLER — Called when the landing page form is submitted
// ============================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Parse JSON body sent from the landing page
    var data = JSON.parse(e.postData.contents);
    var name = sanitize(data.name);
    var instagram = sanitize(data.instagram);

    // Validate required fields
    if (!name || !instagram) {
      return jsonResponse({ success: false, error: 'Nama dan Instagram wajib diisi.' });
    }

    // Check for duplicate Instagram username (case-insensitive)
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][2] && String(existing[i][2]).toLowerCase() === instagram.toLowerCase()) {
        return jsonResponse({ success: false, error: 'duplicate' });
      }
    }

    // Append the new entry
    sheet.appendRow([
      new Date(),     // Timestamp
      name,           // Nama
      instagram       // Instagram
    ]);

    return jsonResponse({ success: true, message: 'Pendaftaran berhasil!' });

  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ============================================================
//  GET HANDLER — For testing the endpoint in a browser
// ============================================================

function doGet() {
  return jsonResponse({
    success: true,
    message: 'Simsalaclean waitlist endpoint is running. 🎉',
    count: getWaitlistCount()
  });
}

// ============================================================
//  HELPER FUNCTIONS
// ============================================================

/**
 * Trim and sanitize user input (max 100 chars).
 */
function sanitize(value) {
  if (!value) return '';
  return String(value).trim().slice(0, 100);
}

/**
 * Return a JSON HTTP response.
 */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Count total waitlist entries (excluding header row).
 */
function getWaitlistCount() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return Math.max(0, sheet.getLastRow() - 1);
}
