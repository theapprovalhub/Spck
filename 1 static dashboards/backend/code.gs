// ============================================================
// CHOICE PROPERTIES - RENTAL APPLICATION BACKEND (FULLY ALIGNED)
// ============================================================
// Company: Choice Properties
// Email: choicepropertygroup@hotmail.com
// Phone: 707-706-3137 (TEXT ONLY)
// Address: 2265 Livernois, Suite 500, Troy, MI 48083
// ============================================================

// Sheet configuration
const SHEET_NAME = 'Applications';
const SETTINGS_SHEET = 'Settings';
const LOG_SHEET = 'EmailLogs';
const ADMIN_EMAILS_RANGE = 'AdminEmails';

// ============================================================
// Helper: get or create spreadsheet
// ============================================================
function getSpreadsheet() {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      const scriptProperties = PropertiesService.getScriptProperties();
      const savedSheetId = scriptProperties.getProperty('SPREADSHEET_ID');
      if (savedSheetId) {
        try {
          ss = SpreadsheetApp.openById(savedSheetId);
          console.log('Opened existing spreadsheet by ID');
        } catch (e) {
          console.log('Saved sheet ID invalid, creating new');
          ss = SpreadsheetApp.create('Choice Properties Rental Applications');
          scriptProperties.setProperty('SPREADSHEET_ID', ss.getId());
          console.log('Created new spreadsheet: ' + ss.getUrl());
        }
      } else {
        ss = SpreadsheetApp.create('Choice Properties Rental Applications');
        scriptProperties.setProperty('SPREADSHEET_ID', ss.getId());
        console.log('Created new spreadsheet: ' + ss.getUrl());
      }
    }
    return ss;
  } catch (error) {
    console.log('Error getting spreadsheet: ' + error);
    const ss = SpreadsheetApp.create('Choice Properties Rental Applications');
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('SPREADSHEET_ID', ss.getId());
    return ss;
  }
}

// ============================================================
// Initialize sheets and named ranges (headers match HTML names)
// ============================================================
function initializeSheets() {
  const ss = getSpreadsheet();

  // Applications sheet – headers exactly as they appear in HTML name attributes
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      'Timestamp', 'App ID', 'Status', 'Payment Status', 'Payment Date', 'Admin Notes',
      'First Name', 'Last Name', 'Email', 'Phone', 'Property Address', 'Requested Move-in Date',
      'Desired Lease Term', 'DOB', 'SSN', 'Current Address', 'Residency Duration',
      'Current Rent Amount', 'Reason for leaving', 'Current Landlord Name', 'Landlord Phone',
      'Employment Status', 'Employer', 'Job Title', 'Employment Duration',
      'Supervisor Name', 'Supervisor Phone', 'Monthly Income', 'Other Income',
      'Reference 1 Name', 'Reference 1 Phone', 'Reference 2 Name', 'Reference 2 Phone',
      'Emergency Contact Name', 'Emergency Contact Phone', 'Primary Payment Method', 'Primary Payment Method Other',
      'Alternative Payment Method', 'Alternative Payment Method Other', 'Third Choice Payment Method', 'Third Choice Payment Method Other',
      'Has Pets', 'Pet Details', 'Total Occupants', 'Additional Occupants',
      'Ever Evicted', 'Smoker', 'Document URL',
      'Has Co-Applicant', 'Additional Person Role',
      'Co-Applicant First Name', 'Co-Applicant Last Name', 'Co-Applicant Email', 'Co-Applicant Phone',
      'Co-Applicant DOB', 'Co-Applicant SSN', 'Co-Applicant Employer', 'Co-Applicant Job Title',
      'Co-Applicant Monthly Income', 'Co-Applicant Employment Duration', 'Co-Applicant Consent',
      'Vehicle Make', 'Vehicle Model', 'Vehicle Year', 'Vehicle License Plate',
      'Emergency Contact Relationship', 'Preferred Contact Method', 'Preferred Time', 'Preferred Time Specific'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1a5276').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  // Settings sheet with default admin emails
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET);
    settingsSheet.getRange('A1').setValue('Setting').setFontWeight('bold');
    settingsSheet.getRange('B1').setValue('Value').setFontWeight('bold');
    settingsSheet.getRange('A2').setValue('AdminEmails');
    settingsSheet.getRange('B2').setValue('choicepropertygroup@hotmail.com,theapprovalh@gmail.com,jamesdouglaspallock@gmail.com');

    const range = settingsSheet.getRange('B2');
    ss.setNamedRange(ADMIN_EMAILS_RANGE, range);
  }

  // Email logs sheet
  let logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
    logSheet.getRange(1, 1, 1, 6).setValues([[
      'Timestamp', 'Type', 'Recipient', 'Status', 'App ID', 'Error'
    ]]).setFontWeight('bold').setBackground('#1a5276').setFontColor('#ffffff');
  }
}

// ============================================================
// Helper: get column map (dynamic)
// ============================================================
function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index + 1;
  });
  return map;
}

// ============================================================
// Helper: combine checkbox arrays into comma‑separated string
// ============================================================
function getCheckboxValues(formData, fieldName) {
  const val = formData[fieldName];
  if (Array.isArray(val)) {
    return val.join(', ');
  }
  return val || '';
}

// ============================================================
// doGet() – Serve web pages and JSON endpoints
// ============================================================
function doGet(e) {
  initializeSheets();
  const params = e || { parameter: {} };

  // --- EXISTING HTML ENDPOINTS (keep as is) ---
  if (params.parameter.path === 'admin') {
    return renderAdminPanel();
  }
  if (params.parameter.path === 'dashboard' && params.parameter.id) {
    return renderApplicantDashboard(params.parameter.id);
  }
  if (params.parameter.path === 'login') {
    return renderLoginPage();
  }

  // --- NEW JSON ENDPOINTS ---
  if (params.parameter.action === 'getApplication') {
    const id = params.parameter.id;
    const result = getApplication(id);
    // Remove sensitive fields before sending
    if (result.success && result.application) {
      delete result.application['SSN'];
      delete result.application['Co-Applicant SSN'];
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*'); // Allow your static domain later
  }

  if (params.parameter.action === 'listApplications') {
    const filter = params.parameter.filter || 'all';
    const result = getAllApplications(filter);
    if (result.success && result.applications) {
      result.applications.forEach(app => {
        delete app['SSN'];
        delete app['Co-Applicant SSN'];
      });
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin', '*');
  }

  // --- DEFAULT: serve the main application form (or a landing page) ---
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head><title>Choice Properties</title></head>
      <body>
        <h1>Choice Properties API</h1>
        <p>Use ?path=login or ?path=admin for the web interface.</p>
      </body>
    </html>
  `).setTitle('Choice Properties');
}

// ============================================================
// doPost() – Handle form submissions (including JSON actions)
// ============================================================
function doPost(e) {
  try {
    // --- EXISTING multipart/form-data handling (for the main application) ---
    if (e.postData && e.postData.type && e.postData.type.indexOf('multipart/form-data') === 0) {
      // Your existing file upload parsing code – keep it exactly as is.
      let formData = {};
      let fileBlob = null;

      const boundary = e.postData.type.split('boundary=')[1];
      const parts = e.postData.contents.split('--' + boundary);

      parts.forEach(part => {
        if (part.trim() === '' || part === '--') return;

        const headerEnd = part.indexOf('\r\n\r\n');
        if (headerEnd === -1) return;

        const headers = part.substring(0, headerEnd);
        const content = part.substring(headerEnd + 4, part.length - 2);

        const filenameMatch = headers.match(/filename="(.+?)"/);
        if (filenameMatch) {
          const filename = filenameMatch[1];
          const contentTypeMatch = headers.match(/Content-Type: (.+)/);
          const contentType = contentTypeMatch ? contentTypeMatch[1] : 'application/octet-stream';
          fileBlob = Utilities.newBlob(content, contentType, filename);
        } else {
          const nameMatch = headers.match(/name="(.+?)"/);
          if (nameMatch) {
            const fieldName = nameMatch[1];
            if (formData.hasOwnProperty(fieldName)) {
              if (!Array.isArray(formData[fieldName])) {
                formData[fieldName] = [formData[fieldName]];
              }
              formData[fieldName].push(content);
            } else {
              formData[fieldName] = content;
            }
          }
        }
      });

      const result = processApplication(formData, fileBlob);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // --- NEW JSON handling for dashboard actions ---
    else if (e.postData && e.postData.type === 'application/json') {
      const data = JSON.parse(e.postData.contents);
      const action = data.action;
      const appId = data.appId;
      const notes = data.notes || '';

      if (action === 'markPaid') {
        const result = markAsPaid(appId, notes);
        return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (action === 'updateStatus') {
        const newStatus = data.status;
        const result = updateStatus(appId, newStatus, notes);
        return ContentService.createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    // Fallback for other POST data
    else {
      const result = { success: false, error: 'Unsupported request' };
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// processApplication() – Save to sheet and send emails
// ============================================================
function processApplication(formData, fileBlob = null) {
  try {
    // Validate required fields
    const requiredFields = ['First Name', 'Last Name', 'Email', 'Phone'];
    for (let field of requiredFields) {
      if (!formData[field] || formData[field].trim() === '') {
        throw new Error(`Missing required field: ${field}`);
      }
    }

    const ss = getSpreadsheet();
    initializeSheets();

    const sheet = ss.getSheetByName(SHEET_NAME);
    const col = getColumnMap(sheet);

    const appId = formData.appId || generateAppId();

    let fileUrl = '';
    if (fileBlob) {
      try {
        const file = DriveApp.createFile(fileBlob);
        fileUrl = file.getUrl();
      } catch (err) {
        console.error('File upload error:', err);
      }
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = [];
    headers.forEach(header => {
      switch (header) {
        case 'Timestamp': rowData.push(new Date()); break;
        case 'App ID': rowData.push(appId); break;
        case 'Status': rowData.push('pending'); break;
        case 'Payment Status': rowData.push('unpaid'); break;
        case 'Payment Date': rowData.push(''); break;
        case 'Admin Notes': rowData.push(''); break;
        case 'Document URL': rowData.push(fileUrl); break;
        case 'Preferred Contact Method': rowData.push(getCheckboxValues(formData, 'Preferred Contact Method')); break;
        case 'Preferred Time': rowData.push(getCheckboxValues(formData, 'Preferred Time')); break;
        case 'Preferred Time Specific': rowData.push(formData['Preferred Time Specific'] || ''); break;
        default:
          rowData.push(formData[header] || '');
      }
    });

    sheet.appendRow(rowData);

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
      range.setBorder(true, true, true, true, true, true);
      for (let i = 2; i <= lastRow; i++) {
        if (i % 2 === 0) {
          sheet.getRange(i, 1, 1, headers.length).setBackground('#f8f9fa');
        }
      }
    }

    sendApplicantConfirmation(formData, appId);
    sendAdminNotification(formData, appId);
    logEmail('application_submitted', formData['Email'], 'success', appId);

    return {
      success: true,
      appId: appId,
      message: 'Application submitted successfully'
    };

  } catch (error) {
    console.error('processApplication error:', error);
    logEmail('application_submitted', formData['Email'] || 'unknown', 'failed', null, error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================================
// generateAppId() – Unique ID
// ============================================================
function generateAppId() {
  const date = new Date();
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const random = Math.random().toString(36).substring(2, 8).toUpperCase();
  const ms = String(date.getMilliseconds()).padStart(3, '0');
  return `CP-${year}${month}${day}-${random}${ms}`;
}

// ============================================================
// Email Templates (updated to use new header names)
// ============================================================
const EmailTemplates = {
  applicantConfirmation: (data, appId, loginLink, paymentMethods) => `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background: #f5f7fa; }
        .container { max-width: 600px; margin: 20px auto; background: white; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #1a5276 0%, #3498db 100%); color: white; padding: 30px; text-align: center; }
        .logo { font-size: 32px; font-weight: 800; margin-bottom: 10px; letter-spacing: 1px; }
        .tagline { font-size: 14px; opacity: 0.9; }
        .content { padding: 30px; }
        .summary-box { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #3498db; }
        .app-id { background: #e8f4fc; padding: 15px; border-radius: 8px; font-family: monospace; font-size: 20px; text-align: center; margin: 20px 0; border: 1px dashed #3498db; }
        .payment-preferences { background: #fff3cd; padding: 20px; border-radius: 8px; margin: 20px 0; border: 1px solid #f39c12; }
        .next-steps { background: #d4edda; padding: 20px; border-radius: 8px; margin: 20px 0; border: 1px solid #27ae60; }
        .step { display: flex; align-items: center; margin: 15px 0; }
        .step-number { width: 30px; height: 30px; background: #1a5276; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: bold; margin-right: 15px; }
        .cta-button { display: inline-block; background: #1a5276; color: white; padding: 15px 30px; text-decoration: none; border-radius: 50px; font-weight: bold; margin: 10px 0; transition: background 0.3s; }
        .cta-button:hover { background: #3498db; }
        .contact-info { background: #f8f9fa; padding: 20px; border-radius: 8px; margin-top: 20px; text-align: center; }
        .footer { background: #1a5276; color: white; padding: 20px; text-align: center; font-size: 14px; }
        .footer a { color: white; text-decoration: underline; }
        .divider { height: 1px; background: #ddd; margin: 25px 0; }
        .highlight { color: #1a5276; font-weight: bold; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <div class="logo">🏢 CHOICE PROPERTIES</div>
          <div class="tagline">Professional Property Management</div>
        </div>
        
        <div class="content">
          <h2 style="color: #1a5276; margin-top: 0;">Dear ${data['First Name'] || 'Applicant'},</h2>
          
          <p>Thank you for choosing Choice Properties for your rental application. We have successfully received your application and our team is ready to assist you.</p>
          
          <div class="app-id">
            <strong>Application ID:</strong> ${appId}
          </div>
          
          <div class="summary-box">
            <h3 style="margin-top: 0; color: #1a5276;">📋 Application Summary</h3>
            <table style="width: 100%; border-collapse: collapse;">
              <tr><td style="padding: 8px 0;"><strong>Name:</strong></td><td>${data['First Name']} ${data['Last Name']}</td></tr>
              <tr><td style="padding: 8px 0;"><strong>Property:</strong></td><td>${data['Property Address'] || 'Not specified'}</td></tr>
              <tr><td style="padding: 8px 0;"><strong>Move-in Date:</strong></td><td>${data['Requested Move-in Date'] || 'Not specified'}</td></tr>
              <tr><td style="padding: 8px 0;"><strong>Email:</strong></td><td>${data['Email']}</td></tr>
              <tr><td style="padding: 8px 0;"><strong>Phone:</strong></td><td>${data['Phone']}</td></tr>
            </table>
          </div>
          
          <div class="payment-preferences">
            <h3 style="margin-top: 0; color: #856404;">💳 Your Selected Payment Methods</h3>
            ${paymentMethods.map(m => `<p style="margin: 8px 0;">${m}</p>`).join('')}
          </div>
          
          <div class="next-steps">
            <h3 style="margin-top: 0; color: #27ae60;">⏳ What Happens Next?</h3>
            
            <div class="step">
              <div class="step-number">1</div>
              <div><strong>Payment Team Contact</strong> - Our payment specialist will TEXT you at <span class="highlight">${data['Phone']}</span> within 24 hours to arrange the $50 application fee.</div>
            </div>
            
            <div class="step">
              <div class="step-number">2</div>
              <div><strong>Payment Confirmation</strong> - Once payment is received, we'll update your status and notify you via email.</div>
            </div>
            
            <div class="step">
              <div class="step-number">3</div>
              <div><strong>Application Review</strong> - Your application will be reviewed within 2-3 business days after payment.</div>
            </div>
            
            <div class="step">
              <div class="step-number">4</div>
              <div><strong>Final Decision</strong> - You'll receive our decision via email and text.</div>
            </div>
          </div>
          
          <div style="text-align: center; margin: 30px 0;">
            <a href="${loginLink}" class="cta-button">📱 TRACK YOUR APPLICATION</a>
            <p style="font-size: 13px; color: #666; margin-top: 10px;">Click above to view your application status.</p>
          </div>
          
          <div class="contact-info">
            <h4 style="margin-top: 0;">📱 Important - Save Our Number</h4>
            <p style="font-size: 24px; margin: 10px 0;"><strong>707-706-3137</strong></p>
            <p>Our team will TEXT you from this number. Please save it in your contacts.</p>
            <p>📧 Email: choicepropertygroup@hotmail.com</p>
          </div>
          
          <div class="divider"></div>
          
          <p style="font-size: 14px; color: #666;">We're here to help! If you have any questions before our team contacts you, feel free to TEXT us at the number above.</p>
        </div>
        
        <div class="footer">
          <p>
            <strong>Choice Properties</strong><br>
            2265 Livernois, Suite 500<br>
            Troy, MI 48083<br>
            📱 Text: 707-706-3137<br>
            📧 choicepropertygroup@hotmail.com
          </p>
          <p style="font-size: 12px; opacity: 0.8; margin-top: 15px;">This is an automated message. Our team will contact you via text.</p>
        </div>
      </div>
    </body>
    </html>
  `,

  adminNotification: (data, appId, baseUrl, loginLink, paymentMethods) => `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; background: #f5f7fa; }
        .container { max-width: 600px; margin: 20px auto; background: white; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #1a5276 0%, #3498db 100%); color: white; padding: 20px; text-align: center; border-radius: 12px 12px 0 0; }
        .content { padding: 25px; }
        .section { margin: 20px 0; padding: 15px; background: #f8f9fa; border-radius: 8px; border-left: 4px solid #3498db; }
        .payment-box { background: #fff3cd; border: 2px solid #f39c12; padding: 15px; border-radius: 8px; margin: 20px 0; }
        .button-group { display: flex; gap: 10px; flex-wrap: wrap; margin: 20px 0; }
        .button { display: inline-block; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; color: white; }
        .button-primary { background: #27ae60; }
        .button-warning { background: #f39c12; }
        .button-info { background: #3498db; }
        .button-danger { background: #e74c3c; }
        .detail-grid { display: grid; grid-template-columns: 1fr 2fr; gap: 10px; margin: 10px 0; }
        .footer { background: #1a5276; color: white; padding: 15px; text-align: center; border-radius: 0 0 12px 12px; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1 style="margin:0;">🔔 NEW APPLICATION</h1>
          <p style="margin:5px 0 0; opacity:0.9;">${appId}</p>
        </div>
        
        <div class="content">
          <div class="payment-box">
            <h3 style="margin-top:0; color:#856404;">💰 PAYMENT PREFERENCES</h3>
            <ul>
              ${paymentMethods.map(p => `<li><strong>${p}</strong></li>`).join('')}
            </ul>
            <p style="color:#856404;"><strong>⏳ Action Required:</strong> Contact applicant to arrange payment</p>
          </div>
          
          <div class="section">
            <h3 style="margin-top:0;">Applicant Information</h3>
            <div class="detail-grid">
              <div><strong>Name:</strong></div><div>${data['First Name']} ${data['Last Name']}</div>
              <div><strong>Email:</strong></div><div>${data['Email']}</div>
              <div><strong>Phone:</strong></div><div>${data['Phone']} (TEXT PREFERRED)</div>
              <div><strong>Property:</strong></div><div>${data['Property Address']}</div>
              <div><strong>Move-in:</strong></div><div>${data['Requested Move-in Date']}</div>
            </div>
          </div>
          
          <div class="button-group">
            <a href="${baseUrl}?path=admin" class="button button-primary">📊 GO TO DASHBOARD</a>
            <a href="${loginLink}" class="button button-info">👁️ VIEW AS APPLICANT</a>
            <a href="sms:7077063137?body=Hi%20${data['First Name']}%2C%20this%20is%20Choice%20Properties%20regarding%20your%20application%20${appId}" class="button button-warning">📱 TEXT APPLICANT</a>
          </div>
          
          <div class="section">
            <h3 style="margin-top:0;">Quick Summary</h3>
            <div class="detail-grid">
              <div><strong>Employer:</strong></div><div>${data['Employer'] || 'N/A'}</div>
              <div><strong>Income:</strong></div><div>$${data['Monthly Income'] || '0'}/month</div>
              <div><strong>Status:</strong></div><div>${data['Employment Status'] || 'N/A'}</div>
              <div><strong>Pets:</strong></div><div>${data['Has Pets'] || 'No'}</div>
            </div>
          </div>
          
          <p><em>Application submitted: ${new Date().toLocaleString()}</em></p>
        </div>
        
        <div class="footer">
          <p style="margin:0;">Choice Properties Admin | 707-706-3137</p>
        </div>
      </div>
    </body>
    </html>
  `,

  paymentConfirmation: (appId, applicantName, phone, loginLink) => `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f7fa; }
        .container { max-width: 600px; margin: 20px auto; background: white; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%); color: white; padding: 30px; text-align: center; }
        .content { padding: 30px; }
        .success-icon { font-size: 64px; margin-bottom: 15px; }
        .info-box { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #27ae60; }
        .cta-button { display: inline-block; background: #1a5276; color: white; padding: 15px 30px; text-decoration: none; border-radius: 50px; font-weight: bold; margin: 10px 0; }
        .footer { background: #1a5276; color: white; padding: 20px; text-align: center; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <div class="success-icon">✅</div>
          <h1 style="margin:0;">PAYMENT CONFIRMED</h1>
        </div>
        
        <div class="content">
          <h2>Dear ${applicantName},</h2>
          
          <p>Great news! Your payment of <strong>$50.00</strong> for application <strong>${appId}</strong> has been confirmed.</p>
          
          <div class="info-box">
            <p><strong>Payment Status:</strong> <span style="color: #27ae60;">PAID ✓</span></p>
            <p><strong>Application Status:</strong> UNDER REVIEW</p>
            <p><strong>Next Step:</strong> Our team will review your application within 24-48 hours.</p>
          </div>
          
          <div style="text-align: center; margin: 30px 0;">
            <a href="${loginLink}" class="cta-button">📱 VIEW APPLICATION STATUS</a>
            <p style="font-size: 13px; color: #666; margin-top: 10px;">Log in with your email or Application ID</p>
          </div>
          
          <p>We'll TEXT you at <strong>${phone}</strong> once the review is complete.</p>
          
          <p>Questions? TEXT us at <strong>707-706-3137</strong></p>
          
          <p>Best regards,<br>Choice Properties Leasing Team</p>
        </div>
        
        <div class="footer">
          <p style="margin:0;">
            <strong>Choice Properties</strong><br>
            2265 Livernois, Suite 500 | Troy, MI 48083<br>
            📱 707-706-3137
          </p>
        </div>
      </div>
    </body>
    </html>
  `,

  statusUpdate: (appId, firstName, status, reason, loginLink) => {
    const isApproved = status === 'approved';
    const headerColor = isApproved ? '#27ae60' : '#e74c3c';
    const icon = isApproved ? '✅' : '📋';
    return `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f7fa; }
          .container { max-width: 600px; margin: 20px auto; background: white; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); overflow: hidden; }
          .header { background: ${headerColor}; color: white; padding: 30px; text-align: center; }
          .content { padding: 30px; }
          .info-box { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid ${headerColor}; }
          .cta-button { display: inline-block; background: #1a5276; color: white; padding: 15px 30px; text-decoration: none; border-radius: 50px; font-weight: bold; }
          .footer { background: #1a5276; color: white; padding: 20px; text-align: center; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1 style="margin:0;">${icon} APPLICATION ${status.toUpperCase()}</h1>
          </div>
          
          <div class="content">
            <h2>Dear ${firstName},</h2>
            
            <p>Your application <strong>${appId}</strong> has been <strong>${status}</strong>.</p>
            
            <div class="info-box">
              ${reason ? `<p><strong>Reason:</strong> ${reason}</p>` : ''}
              ${isApproved ? '<p><strong>Next Steps:</strong> Our leasing team will contact you within 24 hours to schedule lease signing.</p>' : ''}
            </div>
            
            <div style="text-align: center; margin: 30px 0;">
              <a href="${loginLink}" class="cta-button">📱 VIEW STATUS</a>
              <p style="font-size: 13px; color: #666; margin-top: 10px;">Log in with your email or Application ID</p>
            </div>
            
            <p>Questions? TEXT us at <strong>707-706-3137</strong></p>
            
            <p>Best regards,<br>Choice Properties Leasing Team</p>
          </div>
          
          <div class="footer">
            <p style="margin:0;">
              <strong>Choice Properties</strong><br>
              2265 Livernois, Suite 500 | Troy, MI 48083<br>
              📱 707-706-3137
            </p>
          </div>
        </div>
      </body>
      </html>
    `;
  }
};

// ============================================================
// sendApplicantConfirmation() – uses new header names
// ============================================================
function sendApplicantConfirmation(data, appId) {
  try {
    const subject = `Choice Properties - Application Received (Ref: ${appId})`;

    const paymentMethods = [];
    const primary = data['Primary Payment Method'] || '';
    const primaryOther = data['Primary Payment Method Other'] || '';
    const secondary = data['Alternative Payment Method'] || '';
    const secondaryOther = data['Alternative Payment Method Other'] || '';
    const third = data['Third Choice Payment Method'] || '';
    const thirdOther = data['Third Choice Payment Method Other'] || '';

    if (primary) {
      if (primary === 'Other' && primaryOther) {
        paymentMethods.push(`<strong>Primary:</strong> ${primaryOther}`);
      } else {
        paymentMethods.push(`<strong>Primary:</strong> ${primary}`);
      }
    }
    if (secondary) {
      if (secondary === 'Other' && secondaryOther) {
        paymentMethods.push(`<strong>Secondary:</strong> ${secondaryOther}`);
      } else {
        paymentMethods.push(`<strong>Secondary:</strong> ${secondary}`);
      }
    }
    if (third) {
      if (third === 'Other' && thirdOther) {
        paymentMethods.push(`<strong>Third Choice:</strong> ${thirdOther}`);
      } else {
        paymentMethods.push(`<strong>Third Choice:</strong> ${third}`);
      }
    }

    const baseUrl = ScriptApp.getService().getUrl();
    // --- UPDATED LINK TO STATIC DASHBOARD ---
    const loginLink = 'https://yourdomain.com/applicant-dashboard.html?id=' + appId; // Replace with your actual domain

    const htmlBody = EmailTemplates.applicantConfirmation(data, appId, loginLink, paymentMethods);

    const plainBody = `
CHOICE PROPERTIES - Application Received

Dear ${data['First Name'] || 'Applicant'},

Thank you for choosing Choice Properties. We have received your application.

Application ID: ${appId}

WHAT HAPPENS NEXT:
1. Our payment team will TEXT you at ${data['Phone']} within 24 hours to arrange the $50 application fee.
2. Once payment is confirmed, your application will move to review.
3. You'll receive updates via email and can track status online.

Track your application: ${loginLink}
(You will need your email or Application ID to log in)

Questions? TEXT us at 707-706-3137

Choice Properties
2265 Livernois, Suite 500
Troy, MI 48083
    `;

    MailApp.sendEmail({
      to: data['Email'],
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody,
      name: 'Choice Properties Leasing'
    });

    return true;
  } catch (error) {
    console.error('sendApplicantConfirmation error:', error);
    return false;
  }
}

// ============================================================
// sendAdminNotification() – uses new header names
// ============================================================
function sendAdminNotification(data, appId) {
  try {
    const ss = getSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
    let adminEmails = ['choicepropertygroup@hotmail.com'];
    const namedRange = ss.getRangeByName(ADMIN_EMAILS_RANGE);
    if (namedRange) {
      const emailsString = namedRange.getValue();
      adminEmails = emailsString.split(',').map(e => e.trim());
    }

    const paymentMethods = [];
    const primary = data['Primary Payment Method'] || '';
    const primaryOther = data['Primary Payment Method Other'] || '';
    const secondary = data['Alternative Payment Method'] || '';
    const secondaryOther = data['Alternative Payment Method Other'] || '';
    const third = data['Third Choice Payment Method'] || '';
    const thirdOther = data['Third Choice Payment Method Other'] || '';

    if (primary) {
      if (primary === 'Other' && primaryOther) {
        paymentMethods.push(`🥇 ${primaryOther}`);
      } else {
        paymentMethods.push(`🥇 ${primary}`);
      }
    }
    if (secondary) {
      if (secondary === 'Other' && secondaryOther) {
        paymentMethods.push(`🥈 ${secondaryOther}`);
      } else {
        paymentMethods.push(`🥈 ${secondary}`);
      }
    }
    if (third) {
      if (third === 'Other' && thirdOther) {
        paymentMethods.push(`🥉 ${thirdOther}`);
      } else {
        paymentMethods.push(`🥉 ${third}`);
      }
    }

    const baseUrl = ScriptApp.getService().getUrl();
    const loginLink = baseUrl + '?path=login'; // Keep old link for admin email? Or update to static? We'll keep as is for now.
    const subject = `🔔 NEW APPLICATION: ${appId} - ${data['First Name']} ${data['Last Name']}`;

    const htmlBody = EmailTemplates.adminNotification(data, appId, baseUrl, loginLink, paymentMethods);

    adminEmails.forEach(email => {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody,
        name: 'Choice Properties System'
      });
    });

    return true;
  } catch (error) {
    console.error('sendAdminNotification error:', error);
    return false;
  }
}

// ============================================================
// sendPaymentConfirmation() – unchanged (uses appId, email, name, phone)
// ============================================================
function sendPaymentConfirmation(appId, applicantEmail, applicantName, phone) {
  try {
    const subject = `✅ Payment Confirmed - Application ${appId}`;
    const baseUrl = ScriptApp.getService().getUrl();
    const loginLink = baseUrl + '?path=login'; // Could also update to static, but not critical
    const htmlBody = EmailTemplates.paymentConfirmation(appId, applicantName, phone, loginLink);
    MailApp.sendEmail({
      to: applicantEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: 'Choice Properties'
    });
    return true;
  } catch (error) {
    console.error('sendPaymentConfirmation error:', error);
    return false;
  }
}

// ============================================================
// sendStatusUpdateEmail() – unchanged
// ============================================================
function sendStatusUpdateEmail(appId, email, firstName, status, reason) {
  try {
    const baseUrl = ScriptApp.getService().getUrl();
    const loginLink = baseUrl + '?path=login';
    const subject = status === 'approved' ? `✅ Application Approved - ${appId}` : `Application Update - ${appId}`;
    const htmlBody = EmailTemplates.statusUpdate(appId, firstName, status, reason, loginLink);
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      name: 'Choice Properties'
    });
    return true;
  } catch (error) {
    console.error('sendStatusUpdateEmail error:', error);
    return false;
  }
}

// ============================================================
// markAsPaid() – updated to use dynamic column map
// ============================================================
function markAsPaid(appId, notes) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Applications sheet not found');

    const col = getColumnMap(sheet);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][col['App ID'] - 1] === appId) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Application not found');

    const currentPaymentStatus = sheet.getRange(rowIndex, col['Payment Status']).getValue();
    if (currentPaymentStatus === 'paid') {
      throw new Error('Application already marked as paid');
    }

    sheet.getRange(rowIndex, col['Payment Status']).setValue('paid');
    sheet.getRange(rowIndex, col['Payment Date']).setValue(new Date());

    if (notes) {
      const currentNotes = sheet.getRange(rowIndex, col['Admin Notes']).getValue();
      const newNote = `[${new Date().toLocaleString()}] Payment marked as paid. ${notes}`;
      sheet.getRange(rowIndex, col['Admin Notes']).setValue(currentNotes ? currentNotes + '\n' + newNote : newNote);
    }

    const email = sheet.getRange(rowIndex, col['Email']).getValue();
    const firstName = sheet.getRange(rowIndex, col['First Name']).getValue();
    const lastName = sheet.getRange(rowIndex, col['Last Name']).getValue();
    const phone = sheet.getRange(rowIndex, col['Phone']).getValue();

    sendPaymentConfirmation(appId, email, firstName + ' ' + lastName, phone);
    logEmail('payment_confirmation', email, 'success', appId);

    return { success: true, message: 'Application marked as paid' };
  } catch (error) {
    console.error('markAsPaid error:', error);
    logEmail('payment_confirmation', 'admin', 'failed', appId, error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// updateStatus() – updated to use dynamic column map
// ============================================================
function updateStatus(appId, newStatus, notes) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Applications sheet not found');

    const col = getColumnMap(sheet);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][col['App ID'] - 1] === appId) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Application not found');

    const paymentStatus = sheet.getRange(rowIndex, col['Payment Status']).getValue();
    if (paymentStatus !== 'paid') {
      throw new Error('Cannot change status until payment is received');
    }

    const currentStatus = sheet.getRange(rowIndex, col['Status']).getValue();
    if (currentStatus === newStatus) {
      throw new Error(`Application already ${newStatus}`);
    }

    sheet.getRange(rowIndex, col['Status']).setValue(newStatus);

    if (notes) {
      const currentNotes = sheet.getRange(rowIndex, col['Admin Notes']).getValue();
      const newNote = `[${new Date().toLocaleString()}] Status changed to ${newStatus}. ${notes}`;
      sheet.getRange(rowIndex, col['Admin Notes']).setValue(currentNotes ? currentNotes + '\n' + newNote : newNote);
    }

    const email = sheet.getRange(rowIndex, col['Email']).getValue();
    const firstName = sheet.getRange(rowIndex, col['First Name']).getValue();

    sendStatusUpdateEmail(appId, email, firstName, newStatus, notes);
    logEmail('status_update', email, 'success', appId);

    return { success: true, message: `Status updated to ${newStatus}` };
  } catch (error) {
    console.error('updateStatus error:', error);
    logEmail('status_update', 'admin', 'failed', appId, error.toString());
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// getApplication() – retrieve by ID or email
// ============================================================
function getApplication(query) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Applications sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const appId = row[1];
      const email = row[8];

      if (appId === query || email === query) {
        const result = {};
        headers.forEach((header, index) => {
          result[header] = row[index];
        });
        delete result['SSN'];
        delete result['Co-Applicant SSN'];
        return { success: true, application: result };
      }
    }

    return { success: false, error: 'Application not found' };
  } catch (error) {
    console.error('getApplication error:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// getAllApplications() – for admin panel
// ============================================================
function getAllApplications(filterStatus) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Applications sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const applications = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[2];
      const paymentStatus = row[3];

      let displayStatus = paymentStatus === 'unpaid' ? 'pending' :
                         (status === 'approved' ? 'approved' :
                         (status === 'denied' ? 'denied' :
                         (paymentStatus === 'paid' ? 'reviewing' : 'pending')));

      if (filterStatus && filterStatus !== 'all') {
        if (filterStatus === 'pending' && displayStatus !== 'pending') continue;
        if (filterStatus === 'paid' && paymentStatus !== 'paid') continue;
        if (filterStatus === 'approved' && status !== 'approved') continue;
        if (filterStatus === 'denied' && status !== 'denied') continue;
      }

      const app = {};
      headers.forEach((header, index) => {
        if (!header.includes('SSN')) {
          app[header] = row[index];
        }
      });
      app['DisplayStatus'] = displayStatus;
      applications.push(app);
    }

    applications.sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']));

    return { success: true, applications: applications };
  } catch (error) {
    console.error('getAllApplications error:', error);
    return { success: false, error: error.toString() };
  }
}

// ============================================================
// logEmail() – track email sending
// ============================================================
function logEmail(type, recipient, status, appId, errorMsg) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET);
      sheet.getRange(1, 1, 1, 6).setValues([[
        'Timestamp', 'Type', 'Recipient', 'Status', 'App ID', 'Error'
      ]]).setFontWeight('bold').setBackground('#1a5276').setFontColor('#ffffff');
    }
    sheet.appendRow([new Date(), type, recipient, status, appId || '', errorMsg || '']);
  } catch (error) {
    console.error('logEmail error:', error);
  }
}

// ============================================================
// renderApplicantDashboard() – kept as fallback (unchanged)
// ============================================================
function renderApplicantDashboard(appId) {
  const result = getApplication(appId);

  if (!result.success) {
    return renderLoginPage('Invalid application ID or email. Please try again.');
  }

  const app = result.application;

  let statusColor, statusText, statusIcon;
  if (app['Payment Status'] === 'unpaid') {
    statusColor = '#f39c12';
    statusText = 'PENDING PAYMENT';
    statusIcon = '⏳';
  } else if (app['Status'] === 'approved') {
    statusColor = '#27ae60';
    statusText = 'APPROVED';
    statusIcon = '✅';
  } else if (app['Status'] === 'denied') {
    statusColor = '#e74c3c';
    statusText = 'DENIED';
    statusIcon = '❌';
  } else if (app['Payment Status'] === 'paid') {
    statusColor = '#3498db';
    statusText = 'UNDER REVIEW';
    statusIcon = '🔄';
  } else {
    statusColor = '#7f8c8d';
    statusText = 'PENDING';
    statusIcon = '📝';
  }

  const paymentMethods = [];
  if (app['Primary Payment Method']) {
    if (app['Primary Payment Method'] === 'Other' && app['Primary Payment Method Other']) {
      paymentMethods.push(`<strong>Primary:</strong> ${app['Primary Payment Method Other']}`);
    } else {
      paymentMethods.push(`<strong>Primary:</strong> ${app['Primary Payment Method']}`);
    }
  }
  if (app['Alternative Payment Method']) {
    if (app['Alternative Payment Method'] === 'Other' && app['Alternative Payment Method Other']) {
      paymentMethods.push(`<strong>Secondary:</strong> ${app['Alternative Payment Method Other']}`);
    } else {
      paymentMethods.push(`<strong>Secondary:</strong> ${app['Alternative Payment Method']}`);
    }
  }
  if (app['Third Choice Payment Method']) {
    if (app['Third Choice Payment Method'] === 'Other' && app['Third Choice Payment Method Other']) {
      paymentMethods.push(`<strong>Third Choice:</strong> ${app['Third Choice Payment Method Other']}`);
    } else {
      paymentMethods.push(`<strong>Third Choice:</strong> ${app['Third Choice Payment Method']}`);
    }
  }

  let extraHtml = '';

  if (app['Has Co-Applicant'] && app['Has Co-Applicant'] !== '' && app['Co-Applicant First Name']) {
    extraHtml += '<h4 style="margin:20px 0 10px; color:#1a5276;">👥 Co-Applicant / Guarantor</h4><div style="background:#f8f9fa; padding:15px; border-radius:8px;">';
    extraHtml += '<p><strong>Role:</strong> ' + (app['Additional Person Role'] || 'Not specified') + '</p>';
    extraHtml += '<p><strong>Name:</strong> ' + (app['Co-Applicant First Name'] || '') + ' ' + (app['Co-Applicant Last Name'] || '') + '</p>';
    extraHtml += '<p><strong>Email:</strong> ' + (app['Co-Applicant Email'] || '') + '</p>';
    extraHtml += '<p><strong>Phone:</strong> ' + (app['Co-Applicant Phone'] || '') + '</p>';
    extraHtml += '<p><strong>DOB:</strong> ' + (app['Co-Applicant DOB'] || '') + '</p>';
    extraHtml += '<p><strong>Employer:</strong> ' + (app['Co-Applicant Employer'] || '') + '</p>';
    extraHtml += '<p><strong>Job Title:</strong> ' + (app['Co-Applicant Job Title'] || '') + '</p>';
    extraHtml += '<p><strong>Monthly Income:</strong> $' + (app['Co-Applicant Monthly Income'] || '0') + '</p>';
    extraHtml += '<p><strong>Employment Duration:</strong> ' + (app['Co-Applicant Employment Duration'] || '') + '</p>';
    extraHtml += '</div>';
  }

  if (app['Vehicle Make'] && app['Vehicle Make'] !== '') {
    extraHtml += '<h4 style="margin:20px 0 10px; color:#1a5276;">🚗 Vehicle Information</h4><div style="background:#f8f9fa; padding:15px; border-radius:8px;">';
    extraHtml += '<p><strong>Make:</strong> ' + (app['Vehicle Make'] || '') + '</p>';
    extraHtml += '<p><strong>Model:</strong> ' + (app['Vehicle Model'] || '') + '</p>';
    extraHtml += '<p><strong>Year:</strong> ' + (app['Vehicle Year'] || '') + '</p>';
    extraHtml += '<p><strong>License Plate:</strong> ' + (app['Vehicle License Plate'] || '') + '</p>';
    extraHtml += '</div>';
  }

  if (app['Emergency Contact Relationship']) {
    extraHtml += '<h4 style="margin:20px 0 10px; color:#1a5276;">📞 Emergency Contact Relationship</h4><div style="background:#f8f9fa; padding:15px; border-radius:8px;">';
    extraHtml += '<p><strong>Relationship:</strong> ' + app['Emergency Contact Relationship'] + '</p>';
    extraHtml += '</div>';
  }

  if (app['Preferred Contact Method'] || app['Preferred Time'] || app['Preferred Time Specific']) {
    extraHtml += '<h4 style="margin:20px 0 10px; color:#1a5276;">📱 Contact Preferences</h4><div style="background:#f8f9fa; padding:15px; border-radius:8px;">';
    extraHtml += '<p><strong>Preferred Methods:</strong> ' + (app['Preferred Contact Method'] || 'Not specified') + '</p>';
    extraHtml += '<p><strong>Preferred Times:</strong> ' + (app['Preferred Time'] || 'Not specified') + '</p>';
    extraHtml += '<p><strong>Additional Notes:</strong> ' + (app['Preferred Time Specific'] || 'None') + '</p>';
    extraHtml += '</div>';
  }

  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Application Status - Choice Properties</title>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <style>
        body { background: #f5f7fa; }
        .status-badge { background: ${statusColor}; color: white; padding: 15px; border-radius: 8px; text-align: center; font-size: 20px; font-weight: bold; display: flex; align-items: center; justify-content: center; gap: 10px; }
        .app-id { background: #e8f4fc; padding: 15px; font-family: monospace; font-size: 18px; text-align: center; border-radius: 8px; border: 1px dashed #1a5276; }
        .payment-card { background: #fff3cd; border: 1px solid #f39c12; border-radius: 8px; padding: 20px; }
        .info-card { background: #f8f9fa; border-left: 4px solid #3498db; border-radius: 8px; padding: 20px; }
        .toggle-btn { background: #ecf0f1; border: 1px solid #bdc3c7; color: #2c3e50; padding: 8px 16px; border-radius: 30px; font-size: 14px; cursor: pointer; transition: all 0.2s; }
        .toggle-btn:hover { background: #bdc3c7; }
        .extra-details { display: none; margin-top: 20px; }
      </style>
    </head>
    <body>
      <div class="container py-4">
        <div class="row justify-content-center">
          <div class="col-lg-8">
            <div class="card shadow-lg border-0 rounded-4">
              <div class="card-header bg-primary text-white text-center py-4 rounded-top-4">
                <h2 class="h4 mb-0">🏢 Choice Properties</h2>
                <p class="mb-0 small">Application Status</p>
              </div>
              <div class="card-body p-4">
                <div class="status-badge mb-4">
                  <span>${statusIcon}</span> ${statusText}
                </div>

                <div class="app-id mb-4">
                  <strong>Application ID:</strong> ${app['App ID']}
                </div>

                ${app['Payment Status'] === 'unpaid' ? `
                  <div class="payment-card mb-4">
                    <h5 class="text-warning">⏳ PAYMENT PENDING</h5>
                    <p>Your application is on hold until the $50 fee is paid.</p>
                    <p>Our payment team will TEXT you at <strong>${app['Phone']}</strong> within 24 hours to arrange payment.</p>
                    <div class="bg-light p-3 rounded">
                      <h6 class="mb-2">Your Selected Payment Methods:</h6>
                      ${paymentMethods.map(m => `<p class="mb-1">${m}</p>`).join('')}
                    </div>
                  </div>
                ` : ''}

                <div class="info-card mb-4">
                  <h5>Application Details</h5>
                  <div class="row g-3">
                    <div class="col-sm-6">
                      <div class="bg-white p-3 rounded shadow-sm">
                        <strong class="text-primary d-block">Property</strong>
                        ${app['Property Address'] || 'Not specified'}
                      </div>
                    </div>
                    <div class="col-sm-6">
                      <div class="bg-white p-3 rounded shadow-sm">
                        <strong class="text-primary d-block">Move-in Date</strong>
                        ${app['Requested Move-in Date'] || 'Not specified'}
                      </div>
                    </div>
                    <div class="col-sm-6">
                      <div class="bg-white p-3 rounded shadow-sm">
                        <strong class="text-primary d-block">Name</strong>
                        ${app['First Name']} ${app['Last Name']}
                      </div>
                    </div>
                    <div class="col-sm-6">
                      <div class="bg-white p-3 rounded shadow-sm">
                        <strong class="text-primary d-block">Email</strong>
                        ${app['Email']}
                      </div>
                    </div>
                    <div class="col-sm-6">
                      <div class="bg-white p-3 rounded shadow-sm">
                        <strong class="text-primary d-block">Phone</strong>
                        ${app['Phone']}
                      </div>
                    </div>
                    <div class="col-sm-6">
                      <div class="bg-white p-3 rounded shadow-sm">
                        <strong class="text-primary d-block">Lease Term</strong>
                        ${app['Desired Lease Term'] || 'Not specified'}
                      </div>
                    </div>
                  </div>
                </div>

                <button class="toggle-btn btn btn-outline-secondary w-100 mb-3" onclick="toggleDetails()">📋 SHOW FULL APPLICATION DETAILS</button>

                <div id="extraDetails" class="extra-details">
                  ${extraHtml || '<p class="text-muted">No additional details provided.</p>'}
                </div>

                <div class="bg-primary text-white p-4 rounded-4 mt-4 text-center">
                  <h5 class="mb-3">📱 QUESTIONS? TEXT US</h5>
                  <p class="display-6 mb-2">707-706-3137</p>
                  <p class="mb-0">📧 choicepropertygroup@hotmail.com</p>
                  <p class="mt-3 small">2265 Livernois, Suite 500<br>Troy, MI 48083</p>
                </div>

                <hr class="my-4">
                <a href="?path=login" class="btn btn-outline-primary w-100">← Check Another Application</a>
              </div>
              <div class="card-footer text-center text-muted small py-3">
                Choice Properties - Professional Property Management
              </div>
            </div>
          </div>
        </div>
      </div>

      <script>
        function toggleDetails() {
          const details = document.getElementById('extraDetails');
          const btn = document.querySelector('.toggle-btn');
          if (details.style.display === 'none' || details.style.display === '') {
            details.style.display = 'block';
            btn.textContent = '📋 HIDE FULL DETAILS';
          } else {
            details.style.display = 'none';
            btn.textContent = '📋 SHOW FULL APPLICATION DETAILS';
          }
        }
      </script>
    </body>
    </html>
  `).setTitle(`Application ${app['App ID']} - Choice Properties`);
}

// ============================================================
// renderAdminPanel() – kept as fallback (unchanged)
// ============================================================
function renderAdminPanel() {
  initializeSheets();
  const ss = getSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);
  let authorizedEmails = [];
  const namedRange = ss.getRangeByName(ADMIN_EMAILS_RANGE);
  if (namedRange) {
    const emailsString = namedRange.getValue();
    authorizedEmails = emailsString.split(',').map(e => e.trim());
  } else {
    authorizedEmails = ['choicepropertygroup@hotmail.com', 'theapprovalh@gmail.com', 'jamesdouglaspallock@gmail.com'];
  }

  const userEmail = Session.getActiveUser().getEmail();
  if (!authorizedEmails.includes(userEmail)) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Access Denied</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      </head>
      <body>
        <div class="container mt-5">
          <div class="alert alert-danger text-center">
            <h2>⛔ Access Denied</h2>
            <p>You are not authorized to access the admin panel.</p>
            <p><strong>Logged in as:</strong> ${userEmail}</p>
          </div>
        </div>
      </body>
      </html>
    `).setTitle('Access Denied');
  }

  const result = getAllApplications();
  const applications = result.success ? result.applications : [];

  const pendingPayment = applications.filter(a => a['Payment Status'] === 'unpaid').length;
  const paid = applications.filter(a => a['Payment Status'] === 'paid' && a['Status'] !== 'approved' && a['Status'] !== 'denied').length;
  const approved = applications.filter(a => a['Status'] === 'approved').length;
  const denied = applications.filter(a => a['Status'] === 'denied').length;
  const total = applications.length;

  const baseUrl = ScriptApp.getService().getUrl();

  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Admin Dashboard - Choice Properties</title>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <style>
        .status-pending { background: #fff3cd; color: #856404; }
        .status-paid { background: #d4edda; color: #155724; }
        .status-approved { background: #c3e6cb; color: #155724; }
        .status-denied { background: #f8d7da; color: #721c24; }
        .contact-pref { background: #e7f1ff; padding: 8px; border-radius: 20px; display: inline-block; font-size: 12px; margin-right: 5px; }
        .modal { display: none; position: fixed; z-index: 1050; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.5); }
        .modal-content { background-color: #fff; margin: 10% auto; padding: 20px; border-radius: 12px; max-width: 500px; }
      </style>
    </head>
    <body>
      <div class="container-fluid py-4">
        <div class="d-flex justify-content-between align-items-center mb-4">
          <h1 class="h3">🏢 Admin Dashboard</h1>
          <span class="badge bg-secondary">${userEmail}</span>
        </div>

        <div class="row g-3 mb-4">
          <div class="col"><div class="card text-center p-3"><span class="fs-3 fw-bold">${pendingPayment}</span><span class="text-muted">Pending Payment</span></div></div>
          <div class="col"><div class="card text-center p-3"><span class="fs-3 fw-bold">${paid}</span><span class="text-muted">Paid</span></div></div>
          <div class="col"><div class="card text-center p-3"><span class="fs-3 fw-bold">${approved}</span><span class="text-muted">Approved</span></div></div>
          <div class="col"><div class="card text-center p-3"><span class="fs-3 fw-bold">${denied}</span><span class="text-muted">Denied</span></div></div>
          <div class="col"><div class="card text-center p-3"><span class="fs-3 fw-bold">${total}</span><span class="text-muted">Total</span></div></div>
        </div>

        <input type="text" class="form-control mb-3" id="searchInput" placeholder="🔍 Search by name, email, ID, or property...">

        <div class="btn-group mb-3" id="filterButtons">
          <button class="btn btn-outline-primary active" onclick="filterApps('all', this)">All</button>
          <button class="btn btn-outline-primary" onclick="filterApps('pending', this)">Pending Payment</button>
          <button class="btn btn-outline-primary" onclick="filterApps('paid', this)">Paid</button>
          <button class="btn btn-outline-primary" onclick="filterApps('approved', this)">Approved</button>
          <button class="btn btn-outline-primary" onclick="filterApps('denied', this)">Denied</button>
        </div>

        <div id="applicationsList">
          ${applications.length === 0 ? `
            <div class="alert alert-info">No applications found</div>
          ` : applications.map(app => {
            const statusClass = app['Payment Status'] === 'unpaid' ? 'status-pending' :
                               (app['Status'] === 'approved' ? 'status-approved' :
                               (app['Status'] === 'denied' ? 'status-denied' :
                               (app['Payment Status'] === 'paid' ? 'status-paid' : 'status-pending')));

            const statusText = app['Payment Status'] === 'unpaid' ? '⏳ PENDING PAYMENT' :
                              (app['Status'] === 'approved' ? '✅ APPROVED' :
                              (app['Status'] === 'denied' ? '❌ DENIED' :
                              (app['Payment Status'] === 'paid' ? '🔄 UNDER REVIEW' : '📝 PENDING')));

            const contactMethod = app['Preferred Contact Method'] || 'Not specified';
            const contactTimes = app['Preferred Time'] || 'Any';
            const contactDisplay = `<span class="contact-pref">📱 ${contactMethod}</span><span class="contact-pref">🕒 ${contactTimes}</span>`;

            const paymentDisplay = [];
            if (app['Primary Payment Method']) {
              if (app['Primary Payment Method'] === 'Other' && app['Primary Payment Method Other']) {
                paymentDisplay.push(`🥇 ${app['Primary Payment Method Other']}`);
              } else {
                paymentDisplay.push(`🥇 ${app['Primary Payment Method']}`);
              }
            }
            if (app['Alternative Payment Method']) {
              if (app['Alternative Payment Method'] === 'Other' && app['Alternative Payment Method Other']) {
                paymentDisplay.push(`🥈 ${app['Alternative Payment Method Other']}`);
              } else {
                paymentDisplay.push(`🥈 ${app['Alternative Payment Method']}`);
              }
            }
            if (app['Third Choice Payment Method']) {
              if (app['Third Choice Payment Method'] === 'Other' && app['Third Choice Payment Method Other']) {
                paymentDisplay.push(`🥉 ${app['Third Choice Payment Method Other']}`);
              } else {
                paymentDisplay.push(`🥉 ${app['Third Choice Payment Method']}`);
              }
            }

            const searchTerms = (app['First Name'] + ' ' + app['Last Name'] + ' ' + app['Email'] + ' ' + app['App ID'] + ' ' + app['Property Address']).toLowerCase();

            return `
              <div class="card mb-3 application-card" data-status="${app['Payment Status'] === 'unpaid' ? 'pending' : (app['Status'] === 'approved' ? 'approved' : (app['Status'] === 'denied' ? 'denied' : (app['Payment Status'] === 'paid' ? 'paid' : 'pending')))}" data-search="${searchTerms}">
                <div class="card-body">
                  <div class="d-flex justify-content-between align-items-start">
                    <div>
                      <h5 class="card-title mb-1">${app['First Name']} ${app['Last Name']}</h5>
                      <p class="text-muted small mb-2">${app['App ID']} · ${new Date(app['Timestamp']).toLocaleDateString()}</p>
                      <div class="mb-2">${contactDisplay}</div>
                    </div>
                    <span class="badge ${statusClass} p-2">${statusText}</span>
                  </div>
                  <p class="mb-2">📧 ${app['Email']} | 📱 ${app['Phone']}</p>
                  <p class="mb-2">🏠 ${app['Property Address'] || 'Not specified'}</p>

                  ${app['Payment Status'] === 'unpaid' ? `
                    <div class="bg-light p-2 rounded small">
                      <strong>💰 Payment Preferences:</strong> ${paymentDisplay.join(', ')}
                    </div>
                  ` : ''}

                  <div class="mt-3">
                    <button class="btn btn-sm btn-warning" onclick="showConfirmModal('markPaid', '${app['App ID']}', '${app['First Name']} ${app['Last Name']}', '${contactMethod}', '${contactTimes}')" ${app['Payment Status'] !== 'unpaid' ? 'disabled' : ''}>💰 Mark Paid</button>
                    <button class="btn btn-sm btn-success" onclick="showConfirmModal('approve', '${app['App ID']}', '${app['First Name']} ${app['Last Name']}', '${contactMethod}', '${contactTimes}')" ${app['Payment Status'] !== 'paid' || app['Status'] !== 'pending' ? 'disabled' : ''}>✅ Approve</button>
                    <button class="btn btn-sm btn-danger" onclick="showConfirmModal('deny', '${app['App ID']}', '${app['First Name']} ${app['Last Name']}', '${contactMethod}', '${contactTimes}')" ${app['Payment Status'] !== 'paid' || app['Status'] !== 'pending' ? 'disabled' : ''}>❌ Deny</button>
                    <a href="${baseUrl}?path=dashboard&id=${app['App ID']}" class="btn btn-sm btn-info" target="_blank">👁️ View</a>
                    <a href="sms:7077063137?body=Hi%20${app['First Name']}%2C%20this%20is%20Choice%20Properties%20regarding%20application%20${app['App ID']}" class="btn btn-sm btn-secondary">📱 Text</a>
                  </div>
                </div>
              </div>
            `;
          }).join('')}
        </div>

        <!-- Confirmation Modal -->
        <div id="confirmModal" class="modal">
          <div class="modal-content">
            <div class="modal-header">
              <h5 class="modal-title" id="modalTitle">Confirm Action</h5>
              <button type="button" class="btn-close" onclick="closeModal()"></button>
            </div>
            <div class="modal-body" id="modalBody">
              <p id="modalMessage"></p>
              <div class="alert alert-info" id="contactInfo"></div>
              <div class="mb-3" id="notesField" style="display:none;">
                <label for="denialReason" class="form-label">Reason (optional):</label>
                <textarea class="form-control" id="denialReason" rows="2"></textarea>
              </div>
            </div>
            <div class="modal-footer">
              <button class="btn btn-secondary" onclick="closeModal()">Cancel</button>
              <button class="btn btn-primary" id="modalConfirmBtn">Confirm</button>
            </div>
          </div>
        </div>

        <div class="bg-primary text-white p-4 rounded-4 mt-4 text-center">
          <p class="mb-1"><strong>Choice Properties</strong> · 2265 Livernois, Suite 500 · Troy, MI 48083</p>
          <p class="mb-0">📱 707-706-3137 · 📧 choicepropertygroup@hotmail.com</p>
        </div>
      </div>

      <script>
        let currentAction = '';
        let currentAppId = '';

        function filterApps(status, btn) {
          document.querySelectorAll('.btn-group .btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');

          const cards = document.querySelectorAll('.application-card');
          const searchTerm = document.getElementById('searchInput').value.toLowerCase();

          cards.forEach(card => {
            const cardStatus = card.dataset.status;
            const matchesFilter = status === 'all' || cardStatus === status;
            const matchesSearch = searchTerm === '' || card.dataset.search.includes(searchTerm);
            card.style.display = (matchesFilter && matchesSearch) ? 'block' : 'none';
          });
        }

        function filterSearch() {
          const searchTerm = document.getElementById('searchInput').value.toLowerCase();
          const activeFilter = document.querySelector('.btn-group .btn.active')?.innerText.toLowerCase() || 'all';
          let filterValue = 'all';
          if (activeFilter === 'pending payment') filterValue = 'pending';
          else if (activeFilter === 'paid') filterValue = 'paid';
          else if (activeFilter === 'approved') filterValue = 'approved';
          else if (activeFilter === 'denied') filterValue = 'denied';

          const cards = document.querySelectorAll('.application-card');
          cards.forEach(card => {
            const cardStatus = card.dataset.status;
            const matchesFilter = filterValue === 'all' || cardStatus === filterValue;
            const matchesSearch = searchTerm === '' || card.dataset.search.includes(searchTerm);
            card.style.display = (matchesFilter && matchesSearch) ? 'block' : 'none';
          });
        }

        function showConfirmModal(action, appId, applicantName, contactMethod, contactTimes) {
          currentAction = action;
          currentAppId = appId;
          const modal = document.getElementById('confirmModal');
          const title = document.getElementById('modalTitle');
          const message = document.getElementById('modalMessage');
          const contactInfo = document.getElementById('contactInfo');
          const notesField = document.getElementById('notesField');

          if (action === 'markPaid') {
            title.innerText = 'Mark as Paid';
            message.innerText = \`Are you sure you want to mark application \${appId} (\${applicantName}) as paid? This will send a payment confirmation email.\`;
            notesField.style.display = 'none';
          } else if (action === 'approve') {
            title.innerText = 'Approve Application';
            message.innerText = \`Approve application \${appId} (\${applicantName})? This will send an approval email.\`;
            notesField.style.display = 'none';
          } else if (action === 'deny') {
            title.innerText = 'Deny Application';
            message.innerText = \`Deny application \${appId} (\${applicantName})? You can provide a reason below.\`;
            notesField.style.display = 'block';
          }

          contactInfo.innerHTML = \`<strong>Contact Preference:</strong> \${contactMethod} · Preferred times: \${contactTimes}\`;
          modal.style.display = 'block';
        }

        function closeModal() {
          document.getElementById('confirmModal').style.display = 'none';
        }

        document.getElementById('modalConfirmBtn').onclick = function() {
          if (currentAction === 'markPaid') {
            google.script.run
              .withSuccessHandler(() => { alert('Application marked as paid!'); location.reload(); })
              .withFailureHandler(err => alert('Error: ' + err))
              .markAsPaid(currentAppId, '');
          } else if (currentAction === 'approve') {
            google.script.run
              .withSuccessHandler(() => { alert('Application approved!'); location.reload(); })
              .withFailureHandler(err => alert('Error: ' + err))
              .updateStatus(currentAppId, 'approved', '');
          } else if (currentAction === 'deny') {
            const reason = document.getElementById('denialReason').value;
            google.script.run
              .withSuccessHandler(() => { alert('Application denied!'); location.reload(); })
              .withFailureHandler(err => alert('Error: ' + err))
              .updateStatus(currentAppId, 'denied', reason);
          }
          closeModal();
        };

        document.getElementById('searchInput').addEventListener('keyup', filterSearch);
      </script>
    </body>
    </html>
  `).setTitle('Admin Dashboard - Choice Properties');
}

// ============================================================
// renderLoginPage() – unchanged
// ============================================================
function renderLoginPage(errorMsg) {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Applicant Login - Choice Properties</title>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
      <style>
        body { background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ed 100%); min-height: 100vh; display: flex; align-items: center; }
        .card { border: none; border-radius: 20px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); }
        .form-control:focus { border-color: #1a5276; box-shadow: 0 0 0 0.2rem rgba(26,82,118,0.25); }
        .btn-primary { background: #1a5276; border: none; padding: 12px; font-weight: 600; }
        .btn-primary:hover { background: #3498db; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="row justify-content-center">
          <div class="col-md-5">
            <div class="card p-4">
              <div class="text-center mb-4">
                <i class="fas fa-key fa-3x text-primary"></i>
                <h2 class="h4 mt-3">Access Your Application</h2>
                <p class="text-muted">Enter your email or Application ID</p>
              </div>
              ${errorMsg ? `<div class="alert alert-danger">${errorMsg}</div>` : ''}
              <form id="loginForm" onsubmit="event.preventDefault(); login();">
                <div class="mb-3">
                  <label for="query" class="form-label">Email or Application ID</label>
                  <input type="text" class="form-control form-control-lg" id="query" placeholder="e.g., CP-20250315-ABCDEF or email@example.com" required>
                </div>
                <button type="submit" class="btn btn-primary w-100 btn-lg">View My Application</button>
              </form>
              <hr class="my-4">
              <p class="text-center text-muted small">
                <i class="fas fa-phone-alt me-1"></i> Need help? Text us at <strong>707-706-3137</strong>
              </p>
            </div>
          </div>
        </div>
      </div>
      <script>
        function login() {
          const query = document.getElementById('query').value.trim();
          if (!query) return;
          window.location.href = '?path=dashboard&id=' + encodeURIComponent(query);
        }
      </script>
    </body>
    </html>
  `).setTitle('Applicant Login - Choice Properties');
}

// ============================================================
// Test function (optional)
// ============================================================
function runCompleteBackendTest() {
  console.log("🚀 TEST FUNCTION - kept for development");
}