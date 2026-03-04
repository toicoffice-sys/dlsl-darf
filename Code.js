// ============================================
// DATA ACCESS REQUEST FORM (DARF) SYSTEM v2.2
// De La Salle Lipa - Data Privacy Office
// Enhanced: Single/Bulk Request routing + Success Modal
// ============================================

// =============================================
// CONFIGURATION
// Values loaded from Script Properties — never hardcoded.
// To update: Apps Script editor → ⚙️ Project Settings → Script Properties
// =============================================
const _props = PropertiesService.getScriptProperties();

// Always use stored production /exec URL in emails — never /dev
const _webAppUrl = _props.getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl();

const CONFIG = {
  // Non-sensitive — stays in code
  ALLOWED_DOMAIN:             'dlsl.edu.ph',
  APP_NAME:                   'DARF System - DLSL',
  LASALLE_GREEN:              '#1B5E20',
  PROCESSING_DAYS_AVAILABLE:  '1 working day',
  PROCESSING_DAYS_NEW_REPORT: '5 working days',

  // Sensitive — loaded from Script Properties
  DPO_EMAIL:         _props.getProperty('DPO_EMAIL'),
  SHEET_NAME:        _props.getProperty('SHEET_NAME'),
  APPROVAL_SHEET:    _props.getProperty('APPROVAL_SHEET'),
  DATA_OWNERS_SHEET: _props.getProperty('DATA_OWNERS_SHEET'),

  SINGLE_REQUEST: {
    DATA_PROCESSOR_NAME:  _props.getProperty('SINGLE_PROCESSOR_NAME'),
    DATA_PROCESSOR_EMAIL: _props.getProperty('SINGLE_PROCESSOR_EMAIL'),
  },
  BULK_REQUEST: {
    DATA_PROCESSOR_NAME:  _props.getProperty('BULK_PROCESSOR_NAME'),
    DATA_PROCESSOR_EMAIL: _props.getProperty('BULK_PROCESSOR_EMAIL'),
  },

  // Kept for backward compatibility — mirrors SINGLE_REQUEST
  DATA_PROCESSOR_NAME:  _props.getProperty('SINGLE_PROCESSOR_NAME'),
  DATA_PROCESSOR_EMAIL: _props.getProperty('SINGLE_PROCESSOR_EMAIL'),
};

function getProcessorForType(requestType, dataSource) {
  if (requestType === 'bulk') {
    return { name: CONFIG.BULK_REQUEST.DATA_PROCESSOR_NAME, email: CONFIG.BULK_REQUEST.DATA_PROCESSOR_EMAIL };
  }
  // Single request: use per-data-source mapping if available
  if (dataSource && SINGLE_PROCESSOR_MAP[dataSource]) {
    return SINGLE_PROCESSOR_MAP[dataSource];
  }
  return { name: CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_NAME, email: CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_EMAIL };
}

// =============================================
// PUBLIC CONFIG — safe non-sensitive display data for client HTML
// =============================================
function getPublicConfig() {
  return {
    singleProcessorName:  CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_NAME || 'Data Processor',
    bulkProcessorName:    CONFIG.BULK_REQUEST.DATA_PROCESSOR_NAME   || 'Data Processor',
    bulkProcessorEmail:   CONFIG.BULK_REQUEST.DATA_PROCESSOR_EMAIL  || '',
    singleProcessorMap:   SINGLE_PROCESSOR_MAP,
  };
}

// =============================================
// HTML ESCAPING — prevent injection in email templates
// =============================================
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

// =============================================
// DATA OWNER ROUTING MAP
// Loaded from Script Properties (key: DATA_OWNER_MAP) as JSON.
// To update owners: edit the DATA_OWNER_MAP property in ⚙️ Script Properties.
// =============================================
const DATA_OWNER_MAP = JSON.parse(_props.getProperty('DATA_OWNER_MAP') || '{}');

// Per-data-source processor routing for single requests.
// Stored in Script Properties as SINGLE_PROCESSOR_MAP (JSON).
// Example entry: { "Employment Records": { "name": "Data Processor", "email": "hrd.office@dlsl.edu.ph" } }
const SINGLE_PROCESSOR_MAP = JSON.parse(_props.getProperty('SINGLE_PROCESSOR_MAP') || '{}');

// =============================================
// WEB APP ENTRY POINT
// =============================================
function doGet(e) {
  const page = e.parameter.page || 'landing';
  const token = e.parameter.token;
  const action = e.parameter.action;

  if (action === 'approve' && token) {
    return handleApprovalAction(token, 'approve');
  }
  if (action === 'reject' && token) {
    return handleApprovalAction(token, 'reject');
  }
  if (action === 'pending' && token) {
    if (e.parameter.submit === '1') {
      const recommendation = e.parameter.recommendation || '';
      const tokenData = decodeApprovalToken(token);
      if (tokenData && recommendation) {
        return processPendingRecommendation(tokenData, recommendation);
      }
      return createResponsePage('❌ Missing Recommendation', 'Please enter a recommendation before submitting.', 'error');
    }
    return handlePendingAction(token);
  }
  if (action === 'confirm-recommendation' && token) {
    const submitAction   = e.parameter.submit   || '';
    const requesterReply = e.parameter.reply    || '';
    return handleRequesterConfirmation(token, submitAction, requesterReply);
  }
  if (action === 'grant-access' && token) {
    return handleGrantAccess(token);
  }

  if (page === 'status' && e.parameter.id) {
    return showStatus(e.parameter.id);
  }

  if (page === 'dashboard') {
    return serveDashboard();
  }

  if (page === 'form' || page === 'single') {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle(CONFIG.APP_NAME + ' - Single Request')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'bulk') {
    return HtmlService.createTemplateFromFile('IndexBulk')
      .evaluate()
      .setTitle(CONFIG.APP_NAME + ' - Bulk Request')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Landing')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================
// SPREADSHEET INITIALIZATION
// =============================================
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = [
      'Request ID',                   // A - 1
      'Timestamp',                    // B - 2
      'Status',                       // C - 3
      'Requester Name',               // D - 4
      'Requester Email',              // E - 5
      'Requester Office/Course',      // F - 6
      'Date of Request',              // G - 7
      'Data Source Category',         // H - 8
      'Personal Data Requested',      // I - 9
      'File Type',                    // J - 10
      'Purpose',                      // K - 11
      'Internal Recipients',          // L - 12
      'External Recipients',          // M - 13
      'Access Start Date',            // N - 14
      'Access End Date',              // O - 15
      'Security Measures',            // P - 16
      'Immediate Head Name',          // Q - 17
      'Immediate Head Email',         // R - 18
      'Immediate Head Status',        // S - 19
      'Immediate Head Date',          // T - 20
      'Immediate Head Remarks',       // U - 21
      'DPO Status',                   // V - 22
      'DPO Date',                     // W - 23
      'DPO Remarks',                  // X - 24
      'Data Owner Title',             // Y - 25
      'Data Owner Name',              // Z - 26
      'Data Owner Email',             // AA - 27
      'Data Owner Status',            // AB - 28
      'Data Owner Date',              // AC - 29
      'Data Owner Remarks',           // AD - 30
      'Data Processor Name',          // AE - 31
      'Data Processor Email',         // AF - 32
      'Data Processor Status',        // AG - 33
      'Data Processor Date',          // AH - 34
      'Access Method',                // AI - 35
      'Processor Comments',           // AJ - 36
      'Expected Timeline',            // AK - 37
      'Recipients Acknowledged',      // AL - 38
      'Requestee Acknowledged',       // AM - 39
      'Supporting Documents',         // AN - 40
      'Last Updated',                 // AO - 41
      'Request Type'                  // AP - 42
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(CONFIG.LASALLE_GREEN)
      .setFontColor('white')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  let approvalSheet = ss.getSheetByName(CONFIG.APPROVAL_SHEET);
  if (!approvalSheet) {
    approvalSheet = ss.insertSheet(CONFIG.APPROVAL_SHEET);
    const headers = ['Request ID', 'Timestamp', 'Approver Email', 'Action', 'Role', 'Comments'];
    approvalSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground(CONFIG.LASALLE_GREEN)
      .setFontColor('white')
      .setFontWeight('bold');
    approvalSheet.setFrozenRows(1);
  }
}

// =============================================
// EMAIL VALIDATION
// =============================================
function validateDLSLEmail(email) {
  if (!email) return { valid: false, message: 'Email is required' };
  const emailPattern = /^[a-zA-Z0-9._%+-]+@dlsl\.edu\.ph$/;
  if (!emailPattern.test(email.toLowerCase())) {
    return {
      valid: false,
      message: 'Only DLSL email addresses (@dlsl.edu.ph) are allowed for security purposes'
    };
  }
  return { valid: true };
}

// =============================================
// GET DATA CATEGORIES
// =============================================
function getDataCategories() {
  return Object.keys(DATA_OWNER_MAP).map(key => ({
    value: key,
    label: key,
    description: DATA_OWNER_MAP[key].description,
    ownerTitle: DATA_OWNER_MAP[key].ownerTitle
  }));
}

// =============================================
// SUBMIT NEW DARF REQUEST
// =============================================
function submitDARFRequest(formData) {
  try {
    const emailValidation = validateDLSLEmail(formData.requesterEmail);
    if (!emailValidation.valid) {
      return { success: false, message: emailValidation.message };
    }

    const approverValidation = validateDLSLEmail(formData.approverEmail);
    if (!approverValidation.valid) {
      return { success: false, message: 'Approver ' + approverValidation.message };
    }

    if (formData.internalRecipients && formData.internalRecipients.length > 0) {
      for (let recipient of formData.internalRecipients) {
        const recipientValidation = validateDLSLEmail(recipient.email);
        if (!recipientValidation.valid) {
          return {
            success: false,
            message: `Internal recipient ${recipient.name}: ${recipientValidation.message}`
          };
        }
      }
    }

    if (!formData.dataSourceCategory || !DATA_OWNER_MAP[formData.dataSourceCategory]) {
      return { success: false, message: 'Please select a valid data source category.' };
    }

    // Length validation — prevent oversized inputs
    if (formData.requesterName    && formData.requesterName.length    > 200) return { success: false, message: 'Requester name is too long.' };
    if (formData.requesterOffice  && formData.requesterOffice.length  > 300) return { success: false, message: 'Office/course field is too long.' };
    if (formData.approverName     && formData.approverName.length     > 200) return { success: false, message: 'Approver name is too long.' };
    if (formData.approverPosition && formData.approverPosition.length > 300) return { success: false, message: 'Approver position is too long.' };
    if (formData.securityMeasures && formData.securityMeasures.length > 3000) return { success: false, message: 'Security measures field is too long.' };
    if (formData.dataItems        && formData.dataItems.length        > 20)   return { success: false, message: 'Too many data items (max 20).' };
    if (formData.internalRecipients && formData.internalRecipients.length > 20) return { success: false, message: 'Too many internal recipients (max 20).' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      initializeSpreadsheet();
      sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    }

    const requestType = formData.requestType || 'single';
    const processor = getProcessorForType(requestType, formData.dataSourceCategory);
    const typePrefix = requestType === 'bulk' ? 'DARF-B-' : 'DARF-';

    const requestId = typePrefix + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    const timestamp = new Date();

    const dataOwnerInfo = DATA_OWNER_MAP[formData.dataSourceCategory];

    const dataItems = formData.dataItems.map(item =>
      `${item.description} (${item.fileType}) - Purpose: ${item.purpose}`
    ).join('\n');

    const internalRecipients = formData.internalRecipients.map(r =>
      `${r.name} (${r.position}) - ${r.email} - ${r.justification}`
    ).join('\n');

    const externalRecipients = formData.externalRecipients.map(r =>
      `${r.name} - Purpose: ${r.purpose} - MOA/NDA: ${r.hasAgreement}`
    ).join('\n');

    const rowData = [
      requestId,                              // A  - 1
      timestamp,                              // B  - 2
      'Pending - Immediate Head Approval',    // C  - 3
      formData.requesterName,                 // D  - 4
      formData.requesterEmail,                // E  - 5
      formData.requesterOffice,               // F  - 6
      formData.dateOfRequest,                 // G  - 7
      formData.dataSourceCategory,            // H  - 8
      dataItems,                              // I  - 9
      '',                                     // J  - 10
      '',                                     // K  - 11
      internalRecipients,                     // L  - 12
      externalRecipients,                     // M  - 13
      formData.accessStartDate,               // N  - 14
      formData.accessEndDate,                 // O  - 15
      formData.securityMeasures,              // P  - 16
      formData.approverName,                  // Q  - 17
      formData.approverEmail,                 // R  - 18
      'Pending',                              // S  - 19
      '',                                     // T  - 20
      '',                                     // U  - 21
      'Pending',                              // V  - 22
      '',                                     // W  - 23
      '',                                     // X  - 24
      dataOwnerInfo.ownerTitle,               // Y  - 25
      dataOwnerInfo.ownerName,                // Z  - 26
      dataOwnerInfo.ownerEmail,               // AA - 27
      'Pending',                              // AB - 28
      '',                                     // AC - 29
      '',                                     // AD - 30
      processor.name,                         // AE - 31
      processor.email,                        // AF - 32
      'Pending',                              // AG - 33
      '',                                     // AH - 34
      '',                                     // AI - 35
      '',                                     // AJ - 36
      '',                                     // AK - 37
      'Pending',                              // AL - 38
      'Pending',                              // AM - 39
      formData.supportingDocs || '',          // AN - 40
      timestamp,                              // AO - 41
      requestType                             // AP - 42
    ];

    sheet.appendRow(rowData);

    sendApprovalEmail(requestId, formData.approverEmail, formData.approverName, 'Immediate Head', formData, dataOwnerInfo, requestType);
    sendRequesterConfirmation(requestId, formData.requesterEmail, formData.requesterName, formData.dataSourceCategory, dataOwnerInfo, requestType);

    const typeLabel = requestType === 'bulk' ? 'Bulk Request' : 'Single Request';

    return {
      success: true,
      message: 'Request submitted successfully!',
      requestId: requestId,
      requestType: typeLabel
    };

  } catch (error) {
    Logger.log('Error in submitDARFRequest: ' + error.toString());
    return { success: false, message: 'Error submitting request: ' + error.toString() };
  }
}

// =============================================
// TOKEN MANAGEMENT
// =============================================
function generateApprovalToken(requestId, approverEmail, role) {
  const data = {
    requestId: requestId,
    approverEmail: approverEmail,
    role: role,
    timestamp: new Date().getTime()
  };
  return Utilities.base64Encode(JSON.stringify(data));
}

function generatePendingToken(requestId, approverEmail, role, recommendation) {
  const data = {
    requestId: requestId,
    approverEmail: approverEmail,
    role: role,
    recommendation: recommendation,
    timestamp: new Date().getTime(),
    type: 'pending'
  };
  return Utilities.base64Encode(JSON.stringify(data));
}

function decodeApprovalToken(token) {
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    return JSON.parse(decoded);
  } catch (error) {
    return null;
  }
}

// =============================================
// EMAIL TEMPLATES
// =============================================
function getEmailHeader(title, subtitle) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }
        .header { background-color: ${CONFIG.LASALLE_GREEN}; color: white; padding: 20px; text-align: center; }
        .content { padding: 20px; background-color: #f9f9f9; }
        .section { background-color: white; padding: 15px; margin: 10px 0; border-left: 4px solid ${CONFIG.LASALLE_GREEN}; border-radius: 4px; }
        .button { display: inline-block; padding: 12px 24px; margin: 10px 5px; text-decoration: none; border-radius: 5px; font-weight: bold; }
        .approve { background-color: #4CAF50; color: white; }
        .reject { background-color: #f44336; color: white; }
        .pending-rec { background-color: #FF9800; color: white; }
        .info { background-color: #2196F3; color: white; }
        .grant { background-color: #FF9800; color: white; }
        .confirm { background-color: #9C27B0; color: white; }
        .footer { text-align: center; padding: 20px; font-size: 12px; color: #666; background-color: #f0f0f0; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        .label { font-weight: bold; width: 200px; color: #555; }
        .step { text-align: center; padding: 8px 16px; margin: 0 2px; border-radius: 4px; font-size: 12px; }
        .step-active { background-color: ${CONFIG.LASALLE_GREEN}; color: white; font-weight: bold; }
        .step-done { background-color: #81C784; color: white; }
        .step-pending { background-color: #E0E0E0; color: #999; }
        .type-badge { display: inline-block; padding: 3px 10px; border-radius: 10px; font-size: 11px; font-weight: bold; }
        .type-single { background: #E3F2FD; color: #1565C0; }
        .type-bulk { background: #F3E5F5; color: #7B1FA2; }
      </style>
    </head>
    <body>
      <div class="header">
        <h2>${title}</h2>
        <p>${subtitle || 'De La Salle Lipa - Data Privacy Office'}</p>
      </div>
      <div class="content">`;
}

function getEmailFooter() {
  return `
      </div>
      <div class="footer">
        <p>This is an automated message from the DLSL Data Privacy Office.<br>
        For questions, contact: ${CONFIG.DPO_EMAIL}</p>
        <p style="font-size:10px; color:#999;">This email contains secure approval links. Do not forward this email.</p>
      </div>
    </body>
    </html>`;
}

function getProgressTracker(currentStep) {
  const steps = [
    { label: 'Submitted', key: 'submitted' },
    { label: 'Immediate Head', key: 'head' },
    { label: 'DPO', key: 'dpo' },
    { label: 'Data Owner', key: 'owner' },
    { label: 'Data Processor', key: 'processor' },
    { label: 'Complete', key: 'complete' }
  ];

  const stepIndex = steps.findIndex(s => s.key === currentStep);

  let html = '<div style="text-align:center; margin:20px 0;">';
  steps.forEach((step, i) => {
    let className = 'step-pending';
    if (i < stepIndex) className = 'step-done';
    if (i === stepIndex) className = 'step-active';
    html += `<span class="step ${className}">${step.label}</span>`;
    if (i < steps.length - 1) html += '<span style="color:#ccc;"> → </span>';
  });
  html += '</div>';
  return html;
}

function getRequestTypeBadge(requestType) {
  if (requestType === 'bulk') {
    return '<span class="type-badge type-bulk">📦 Bulk Request</span>';
  }
  return '<span class="type-badge type-single">📄 Single Request</span>';
}

// =============================================
// BUILD APPROVAL EMAIL BUTTONS
// =============================================
function buildApprovalButtons(approveUrl, rejectUrl, pendingUrl, statusUrl) {
  return `
    <div style="text-align: center; margin: 30px 0;">
      <a href="${approveUrl}" class="button approve">✅ ENDORSE REQUEST</a>
      <a href="${rejectUrl}" class="button reject">❌ DISAPPROVE REQUEST</a>
      <br><br>
      <a href="${pendingUrl}" class="button pending-rec">⏳ PENDING — WITH RECOMMENDATION</a>
      <br><br>
      <a href="${statusUrl}" class="button info">📊 VIEW FULL DETAILS</a>
    </div>
    <div class="section" style="border-left-color: #FF9800; background-color: #FFF3E0;">
      <p><strong>⏳ Pending with Recommendation:</strong> Use this option to send a recommendation to the requester and pause the workflow until they confirm the recommendation has been followed.</p>
    </div>`;
}

// =============================================
// STEP 1: SEND APPROVAL TO IMMEDIATE HEAD
// =============================================
function sendApprovalEmail(requestId, approverEmail, approverName, role, formData, dataOwnerInfo, requestType) {
  // Escape all user-provided fields before embedding in HTML
  const fd = {
    requesterName:      escapeHtml(formData.requesterName),
    requesterOffice:    escapeHtml(formData.requesterOffice),
    requesterEmail:     escapeHtml(formData.requesterEmail),
    dateOfRequest:      escapeHtml(formData.dateOfRequest),
    dataSourceCategory: escapeHtml(formData.dataSourceCategory),
    accessStartDate:    escapeHtml(formData.accessStartDate),
    accessEndDate:      escapeHtml(formData.accessEndDate),
    securityMeasures:   escapeHtml(formData.securityMeasures),
    dataItems: (formData.dataItems || []).map(function(item) { return {
      description: escapeHtml(item.description),
      fileType:    escapeHtml(item.fileType),
      purpose:     escapeHtml(item.purpose)
    }; }),
    internalRecipients: (formData.internalRecipients || []).map(function(r) { return {
      name:          escapeHtml(r.name),
      position:      escapeHtml(r.position),
      email:         escapeHtml(r.email),
      justification: escapeHtml(r.justification)
    }; }),
    externalRecipients: (formData.externalRecipients || []).map(function(r) { return {
      name:        escapeHtml(r.name),
      purpose:     escapeHtml(r.purpose),
      hasAgreement: escapeHtml(r.hasAgreement)
    }; })
  };
  const token = generateApprovalToken(requestId, approverEmail, role);
  const webAppUrl = _webAppUrl;
  const approveUrl = `${webAppUrl}?action=approve&token=${encodeURIComponent(token)}`;
  const rejectUrl = `${webAppUrl}?action=reject&token=${encodeURIComponent(token)}`;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;
  const pendingUrl = `${webAppUrl}?action=pending&token=${encodeURIComponent(token)}`;

  const processor = getProcessorForType(requestType || 'single');
  const typeLabel = (requestType === 'bulk') ? '📦 Bulk Request' : '📄 Single Request';

  const subject = `[DARF] Data Access Request Approval Required - ${requestId}`;

  const htmlBody = getEmailHeader('🔐 Data Access Request Form', 'Approval Required - ' + role) + `
        ${getProgressTracker('head')}

        <div class="section">
          <h3>Dear ${approverName},</h3>
          <p>A new Data Access Request requires your endorsement as <strong>${role}</strong>.</p>
          <p style="margin-top:8px;">${getRequestTypeBadge(requestType)}</p>

          <table>
            <tr><td class="label">Request ID:</td><td><strong>${requestId}</strong></td></tr>
            <tr><td class="label">Request Type:</td><td>${typeLabel}</td></tr>
            <tr><td class="label">Requester:</td><td>${fd.requesterName}</td></tr>
            <tr><td class="label">Office/Course:</td><td>${fd.requesterOffice}</td></tr>
            <tr><td class="label">Email:</td><td>${fd.requesterEmail}</td></tr>
            <tr><td class="label">Date of Request:</td><td>${fd.dateOfRequest}</td></tr>
            <tr><td class="label">Data Source:</td><td>${fd.dataSourceCategory}</td></tr>
            <tr><td class="label">Data Owner:</td><td>${dataOwnerInfo.ownerTitle}</td></tr>
            <tr><td class="label">Assigned Processor:</td><td>${processor.name} (${processor.email})</td></tr>
          </table>
        </div>

        <div class="section">
          <h4>📋 Data Being Requested:</h4>
          ${fd.dataItems.map(item => `
            <p><strong>• ${item.description}</strong><br>
            File Type: ${item.fileType} | Purpose: ${item.purpose}</p>
          `).join('')}
        </div>

        <div class="section">
          <h4>👥 Internal Recipients:</h4>
          ${fd.internalRecipients.map(r => `
            <p><strong>• ${r.name}</strong> (${r.position})<br>
            Email: ${r.email}<br>
            Justification: ${r.justification}</p>
          `).join('')}
        </div>

        ${fd.externalRecipients.length > 0 ? `
        <div class="section">
          <h4>🌐 External Recipients:</h4>
          ${fd.externalRecipients.map(r => `
            <p><strong>• ${r.name}</strong><br>
            Purpose: ${r.purpose}<br>
            MOA/NDA: ${r.hasAgreement}</p>
          `).join('')}
        </div>` : ''}

        <div class="section">
          <h4>📅 Access Period:</h4>
          <p><strong>Start:</strong> ${fd.accessStartDate} | <strong>End:</strong> ${fd.accessEndDate}</p>
        </div>

        <div class="section">
          <h4>🔒 Security Measures:</h4>
          <p>${fd.securityMeasures}</p>
        </div>

        ${buildApprovalButtons(approveUrl, rejectUrl, pendingUrl, statusUrl)}

        <div class="section">
          <p><em>Upon your endorsement, this request will be forwarded to the <strong>Data Protection Officer</strong> for review.</em></p>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: approverEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// STEP 2: SEND TO DPO
// =============================================
function sendDPOApprovalEmail(requestId, rowData) {
  const token = generateApprovalToken(requestId, CONFIG.DPO_EMAIL, 'Data Protection Officer');
  const webAppUrl = _webAppUrl;
  const approveUrl = `${webAppUrl}?action=approve&token=${encodeURIComponent(token)}`;
  const rejectUrl = `${webAppUrl}?action=reject&token=${encodeURIComponent(token)}`;
  const pendingUrl = `${webAppUrl}?action=pending&token=${encodeURIComponent(token)}`;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;

  const reqType = rowData[41] || 'single';

  const subject = `[DARF] DPO Review Required - ${requestId}`;

  const htmlBody = getEmailHeader('🔍 DPO Review Required', 'Data Protection Officer Review') + `
        ${getProgressTracker('dpo')}

        <div class="section">
          <h3>Data Protection Officer Review</h3>
          <p>The following data access request has been endorsed by the Immediate Head and now requires your review.</p>
          <p style="margin-top:8px;">${getRequestTypeBadge(reqType)}</p>

          <table>
            <tr><td class="label">Request ID:</td><td><strong>${requestId}</strong></td></tr>
            <tr><td class="label">Request Type:</td><td>${reqType === 'bulk' ? '📦 Bulk Request' : '📄 Single Request'}</td></tr>
            <tr><td class="label">Requester:</td><td>${rowData[3]} (${rowData[4]})</td></tr>
            <tr><td class="label">Office/Course:</td><td>${rowData[5]}</td></tr>
            <tr><td class="label">Data Source:</td><td>${rowData[7]}</td></tr>
            <tr><td class="label">Immediate Head:</td><td>✅ Endorsed by ${rowData[16]}</td></tr>
            <tr><td class="label">Assigned Processor:</td><td>${rowData[30]} (${rowData[31]})</td></tr>
          </table>
        </div>

        <div class="section">
          <h4>📋 Data Being Requested:</h4>
          <p>${rowData[8]}</p>
        </div>

        <div class="section">
          <h4>👥 Internal Recipients:</h4>
          <p>${rowData[11] || 'None specified'}</p>
        </div>

        ${rowData[12] ? `
        <div class="section">
          <h4>🌐 External Recipients:</h4>
          <p>${rowData[12]}</p>
        </div>` : ''}

        <div class="section">
          <h4>📅 Access Period & Security:</h4>
          <p><strong>Start:</strong> ${rowData[13]} | <strong>End:</strong> ${rowData[14]}</p>
          <p><strong>Security Measures:</strong> ${rowData[15]}</p>
        </div>

        ${buildApprovalButtons(approveUrl, rejectUrl, pendingUrl, statusUrl)}

        <div class="section">
          <p><strong>Evaluation Criteria:</strong></p>
          <ul>
            <li>Transparency, legitimate purpose, and proportionality</li>
            <li>Adequate safeguards for data integrity, confidentiality, and availability</li>
            <li>Compliance with the Data Privacy Act of 2012 and its IRR</li>
          </ul>
          <p><em>Upon your approval, this request will be forwarded to the <strong>Data Owner (${rowData[24]})</strong> for endorsement.</em></p>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: CONFIG.DPO_EMAIL, subject: subject, htmlBody: htmlBody });
}

// =============================================
// STEP 3: SEND TO DATA OWNER
// =============================================
function sendDataOwnerApprovalEmail(requestId, rowData) {
  const dataOwnerEmail = rowData[26];
  const dataOwnerName = rowData[25];
  const dataOwnerTitle = rowData[24];

  if (!dataOwnerEmail) {
    Logger.log('WARNING: No Data Owner email configured for request ' + requestId);
    sendDataProcessorNotification(requestId, rowData);
    return;
  }

  const token = generateApprovalToken(requestId, dataOwnerEmail, 'Data Owner - ' + dataOwnerTitle);
  const webAppUrl = _webAppUrl;
  const approveUrl = `${webAppUrl}?action=approve&token=${encodeURIComponent(token)}`;
  const rejectUrl = `${webAppUrl}?action=reject&token=${encodeURIComponent(token)}`;
  const pendingUrl = `${webAppUrl}?action=pending&token=${encodeURIComponent(token)}`;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;

  const reqType = rowData[41] || 'single';

  const subject = `[DARF] Data Owner Endorsement Required - ${requestId}`;

  const htmlBody = getEmailHeader('📋 Data Owner Endorsement Required', dataOwnerTitle + ' Review') + `
        ${getProgressTracker('owner')}

        <div class="section">
          <h3>Dear ${dataOwnerName || dataOwnerTitle},</h3>
          <p>A Data Access Request has been endorsed by the Immediate Head and reviewed by the Data Protection Officer. It now requires your endorsement as <strong>Data Owner (${dataOwnerTitle})</strong>.</p>
          <p style="margin-top:8px;">${getRequestTypeBadge(reqType)}</p>

          <table>
            <tr><td class="label">Request ID:</td><td><strong>${requestId}</strong></td></tr>
            <tr><td class="label">Request Type:</td><td>${reqType === 'bulk' ? '📦 Bulk Request' : '📄 Single Request'}</td></tr>
            <tr><td class="label">Requester:</td><td>${rowData[3]} (${rowData[4]})</td></tr>
            <tr><td class="label">Office/Course:</td><td>${rowData[5]}</td></tr>
            <tr><td class="label">Data Source:</td><td>${rowData[7]}</td></tr>
            <tr><td class="label">Immediate Head:</td><td>✅ Endorsed by ${rowData[16]}</td></tr>
            <tr><td class="label">DPO:</td><td>✅ Approved</td></tr>
          </table>
        </div>

        <div class="section">
          <h4>📋 Data Being Requested:</h4>
          <p>${rowData[8]}</p>
        </div>

        <div class="section">
          <h4>👥 Internal Recipients:</h4>
          <p>${rowData[11] || 'None specified'}</p>
        </div>

        <div class="section">
          <h4>📅 Access Period:</h4>
          <p><strong>Start:</strong> ${rowData[13]} | <strong>End:</strong> ${rowData[14]}</p>
        </div>

        <div class="section">
          <h4>🔒 Security Measures:</h4>
          <p>${rowData[15]}</p>
        </div>

        ${buildApprovalButtons(approveUrl, rejectUrl, pendingUrl, statusUrl)}

        <div class="section">
          <p><strong>As Data Owner, please verify:</strong></p>
          <ul>
            <li>The requested data falls within your domain</li>
            <li>The purpose and justification are valid</li>
            <li>The access period and security measures are appropriate</li>
          </ul>
          <p><em>Upon your endorsement, this request will be forwarded to the <strong>Data Processor (${rowData[30]})</strong>.</em></p>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: dataOwnerEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// STEP 4: NOTIFY DATA PROCESSOR
// =============================================
function sendDataProcessorNotification(requestId, rowData) {
  const processorEmail = rowData[31];
  const processorName = rowData[30];

  if (!processorEmail) {
    Logger.log('WARNING: No Data Processor email for request ' + requestId);
    return;
  }

  const reqType = rowData[41] || 'single';
  const webAppUrl = _webAppUrl;
  const processorEmailEncoded = encodeURIComponent(processorEmail);
  const dashboardUrl = `${webAppUrl}?page=dashboard&authuser=${processorEmailEncoded}`;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}&authuser=${processorEmailEncoded}`;

  const subject = `[DARF] Action Required: Provide Data Access - ${requestId}`;

  const htmlBody = getEmailHeader('⚙️ Data Access Processing Required', 'Data Processor Action Needed') + `
        ${getProgressTracker('processor')}

        <div class="section">
          <h3>Dear ${processorName || 'Data Processor'},</h3>
          <p>A Data Access Request has been <strong>fully approved</strong> and requires you to provide access to the requested data.</p>
          <p style="margin-top:8px;">${getRequestTypeBadge(reqType)}</p>

          <table>
            <tr><td class="label">Request ID:</td><td><strong>${requestId}</strong></td></tr>
            <tr><td class="label">Request Type:</td><td>${reqType === 'bulk' ? '📦 Bulk Request' : '📄 Single Request'}</td></tr>
            <tr><td class="label">Requester:</td><td>${rowData[3]} (${rowData[4]})</td></tr>
            <tr><td class="label">Office/Course:</td><td>${rowData[5]}</td></tr>
            <tr><td class="label">Data Source:</td><td>${rowData[7]}</td></tr>
          </table>
        </div>

        <div class="section" style="border-left-color: #FF9800;">
          <h4>✅ Approval Chain Complete:</h4>
          <table>
            <tr><td class="label">Immediate Head:</td><td>✅ ${rowData[16]}</td></tr>
            <tr><td class="label">DPO:</td><td>✅ Approved</td></tr>
            <tr><td class="label">Data Owner (${rowData[24]}):</td><td>✅ Endorsed</td></tr>
          </table>
        </div>

        <div class="section">
          <h4>📋 Data to Provide:</h4>
          <p>${rowData[8]}</p>
        </div>

        <div class="section">
          <h4>👤 Provide Access To:</h4>
          <p><strong>Requester:</strong> ${rowData[3]} (${rowData[4]})</p>
          <p><strong>Internal Recipients:</strong></p>
          <p>${rowData[11] || 'Requester only'}</p>
        </div>

        <div class="section">
          <h4>📅 Terms & Conditions:</h4>
          <p><strong>Access Period:</strong> ${rowData[13]} to ${rowData[14]}</p>
          <p><strong>Security Measures:</strong> ${rowData[15]}</p>
        </div>

        <div class="section" style="border-left-color: #FF9800; background-color: #FFF3E0;">
          <h4>⏱️ Please set the expected timeline in the Dashboard:</h4>
          <p><strong>${CONFIG.PROCESSING_DAYS_AVAILABLE}</strong> — if the data is readily available</p>
          <p><strong>${CONFIG.PROCESSING_DAYS_NEW_REPORT}</strong> — if a new report needs to be generated</p>
        </div>

        <div style="text-align: center; margin: 30px 0;">
          <a href="${dashboardUrl}" class="button grant">📊 OPEN DARF DASHBOARD</a>
          <br><br>
          <a href="${statusUrl}" class="button info">📄 VIEW REQUEST DETAILS</a>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: processorEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// SEND REQUESTER CONFIRMATION
// =============================================
function sendRequesterConfirmation(requestId, requesterEmail, requesterName, dataSourceCategory, dataOwnerInfo, requestType) {
  const webAppUrl = _webAppUrl;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;
  const processor = getProcessorForType(requestType || 'single');
  const typeLabel = (requestType === 'bulk') ? '📦 Bulk Request' : '📄 Single Request';

  const subject = `[DARF] Request Submitted Successfully - ${requestId}`;

  const htmlBody = getEmailHeader('✅ Request Submitted Successfully') + `
        ${getProgressTracker('submitted')}

        <div class="section">
          <h3>Dear ${requesterName},</h3>
          <p>Your Data Access Request has been successfully submitted and is now being processed.</p>

          <p><strong>Request ID:</strong> <span style="color: ${CONFIG.LASALLE_GREEN}; font-size: 18px;">${requestId}</span></p>
          <p><strong>Request Type:</strong> ${typeLabel}</p>
          <p><strong>Data Source:</strong> ${dataSourceCategory}</p>
          <p><strong>Data Owner:</strong> ${dataOwnerInfo.ownerTitle}</p>
          <p><strong>Assigned Processor:</strong> ${processor.name}</p>
        </div>

        <div class="section">
          <h4>📋 Approval Workflow:</h4>
          <ol>
            <li><strong>Immediate Head / Adviser</strong> — will review and endorse the request</li>
            <li><strong>Data Protection Officer</strong> — will evaluate for compliance</li>
            <li><strong>${dataOwnerInfo.ownerTitle} (Data Owner)</strong> — will verify and endorse the request</li>
            <li><strong>${processor.name} (Data Processor)</strong> — will provide access to the requested data</li>
          </ol>
        </div>

        <div class="section" style="border-left-color: #FF9800;">
          <h4>⏱️ Expected Timeline:</h4>
          <p><strong>${CONFIG.PROCESSING_DAYS_AVAILABLE}</strong> — if the data is readily available</p>
          <p><strong>${CONFIG.PROCESSING_DAYS_NEW_REPORT}</strong> — for newly created reports</p>
        </div>

        <div style="text-align: center; margin: 30px 0;">
          <a href="${statusUrl}" class="button info">📊 TRACK REQUEST STATUS</a>
        </div>

        <div class="section">
          <p><strong>Important Reminders:</strong></p>
          <ul>
            <li>You will receive email notifications as your request progresses</li>
            <li>If an approver places your request as "Pending with Recommendation", you will receive instructions to follow before the process continues</li>
            <li>Keep your Request ID for future reference</li>
          </ul>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: requesterEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// HANDLE PENDING WITH RECOMMENDATION
// =============================================
function handlePendingAction(token) {
  try {
    const tokenData = decodeApprovalToken(token);
    if (!tokenData) {
      return createResponsePage('❌ Invalid Token', 'This link is invalid or has expired.', 'error');
    }

    const webAppUrl = _webAppUrl;

    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f5f5; padding: 40px 20px; }
          .container { background: white; padding: 40px; border-radius: 15px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); max-width: 640px; margin: 0 auto; }
          h2 { color: ${CONFIG.LASALLE_GREEN}; margin-bottom: 8px; }
          p { color: #555; margin-bottom: 16px; font-size: 14px; }
          label { display: block; font-weight: 600; color: #333; margin-bottom: 6px; font-size: 14px; }
          textarea { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 8px; font-size: 14px; font-family: inherit; resize: vertical; min-height: 120px; box-sizing: border-box; }
          textarea:focus { outline: none; border-color: ${CONFIG.LASALLE_GREEN}; }
          .btn { display: block; width: 100%; padding: 14px; margin-top: 16px; background: #FF9800; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; }
          .btn:hover { background: #F57C00; }
          .btn:disabled { background: #ccc; cursor: not-allowed; }
          .req-info { background: #f8f9fa; border-left: 4px solid ${CONFIG.LASALLE_GREEN}; padding: 12px 16px; border-radius: 4px; margin-bottom: 20px; font-size: 13px; }
          .req-info strong { color: ${CONFIG.LASALLE_GREEN}; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>⏳ Pending with Recommendation</h2>
          <div class="req-info">
            <strong>Request ID:</strong> ${tokenData.requestId}<br>
            <strong>Your Role:</strong> ${tokenData.role}
          </div>
          <p>Please enter your recommendation or the action the requester must take before this request can proceed.</p>
          <label for="recommendation">Recommendation / Required Action <span style="color:#f44336;">*</span></label>
          <textarea id="recommendation" placeholder="e.g., Please provide a signed consent form from all data subjects before this request can be endorsed."></textarea>
          <button class="btn" id="submitBtn" onclick="submitRecommendation()">⏳ Submit Recommendation</button>
        </div>
        <script>
          function submitRecommendation() {
            var rec = document.getElementById('recommendation').value.trim();
            if (!rec) { alert('Please enter your recommendation.'); return; }
            var btn = document.getElementById('submitBtn');
            btn.disabled = true; btn.textContent = '⏳ Submitting...';
            var token = '${encodeURIComponent(token)}';
            var url = '${webAppUrl}?action=pending&token=' + token + '&recommendation=' + encodeURIComponent(rec) + '&submit=1';
            window.location.href = url;
          }
        <\/script>
      </body>
      </html>
    `).setTitle('Pending with Recommendation');

  } catch (error) {
    Logger.log('Error in handlePendingAction: ' + error.toString());
    return createResponsePage('❌ Error', 'An error occurred: ' + error.toString(), 'error');
  }
}

// =============================================
// PROCESS PENDING RECOMMENDATION
// =============================================
function processPendingRecommendation(tokenData, recommendation) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tokenData.requestId) { rowIndex = i; break; }
    }

    if (rowIndex === -1) {
      return createResponsePage('❌ Request Not Found', `Request ID ${tokenData.requestId} was not found.`, 'error');
    }

    const rowData = data[rowIndex];
    const timestamp = new Date();

    if (tokenData.role === 'Immediate Head') {
      if (rowData[18] !== 'Pending') {
        return createResponsePage('ℹ️ Already Processed', 'This request has already been acted upon.', 'info');
      }
      sheet.getRange(rowIndex + 1, 19).setValue('Pending - With Recommendation');
      sheet.getRange(rowIndex + 1, 20).setValue(timestamp);
      sheet.getRange(rowIndex + 1, 21).setValue(recommendation);
      sheet.getRange(rowIndex + 1, 3).setValue('Pending - Awaiting Requester Action (Immediate Head)');
      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);

    } else if (tokenData.role === 'Data Protection Officer') {
      if (rowData[21] !== 'Pending') {
        return createResponsePage('ℹ️ Already Processed', 'This request has already been acted upon.', 'info');
      }
      sheet.getRange(rowIndex + 1, 22).setValue('Pending - With Recommendation');
      sheet.getRange(rowIndex + 1, 23).setValue(timestamp);
      sheet.getRange(rowIndex + 1, 24).setValue(recommendation);
      sheet.getRange(rowIndex + 1, 3).setValue('Pending - Awaiting Requester Action (DPO)');
      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);

    } else if (tokenData.role.startsWith('Data Owner')) {
      if (rowData[27] !== 'Pending') {
        return createResponsePage('ℹ️ Already Processed', 'This request has already been acted upon.', 'info');
      }
      sheet.getRange(rowIndex + 1, 28).setValue('Pending - With Recommendation');
      sheet.getRange(rowIndex + 1, 29).setValue(timestamp);
      sheet.getRange(rowIndex + 1, 30).setValue(recommendation);
      sheet.getRange(rowIndex + 1, 3).setValue('Pending - Awaiting Requester Action (Data Owner)');
      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);
    }

    logApprovalAction(tokenData.requestId, tokenData.approverEmail, 'pending-with-recommendation', tokenData.role, recommendation);

    const confirmToken = generateApprovalToken(tokenData.requestId, rowData[4], 'Requester-Confirmation-' + tokenData.role);

    sendRecommendationToRequester(tokenData.requestId, rowData, tokenData.role, recommendation, confirmToken);

    return createResponsePage(
      '⏳ Recommendation Sent',
      `Your recommendation has been sent to the requester.<br><br>
       <strong>Request ID:</strong> ${tokenData.requestId}<br>
       <strong>Recommendation:</strong> ${recommendation}<br><br>
       The workflow is paused until the requester confirms they have followed your recommendation.`,
      'info'
    );

  } catch (error) {
    Logger.log('Error in processPendingRecommendation: ' + error.toString());
    return createResponsePage('❌ Error', 'An error occurred: ' + error.toString(), 'error');
  }
}

// =============================================
// SEND RECOMMENDATION EMAIL TO REQUESTER
// =============================================
function sendRecommendationToRequester(requestId, rowData, approverRole, recommendation, confirmToken) {
  const requesterEmail = rowData[4];
  const requesterName  = rowData[3];
  const webAppUrl      = _webAppUrl;
  const responsePageUrl = `${webAppUrl}?action=confirm-recommendation&token=${encodeURIComponent(confirmToken)}`;
  const statusUrl       = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;

  const subject = `[DARF] Action Required: Recommendation from ${approverRole} - ${requestId}`;

  const htmlBody = getEmailHeader('⚠️ Action Required', 'Recommendation from ' + approverRole) + `
        <div class="section" style="border-left-color: #FF9800; background-color: #FFF8E1;">
          <h3>Dear ${requesterName},</h3>
          <p>Your Data Access Request <strong>${requestId}</strong> has been placed <strong>Pending with Recommendation</strong> by the <strong>${approverRole}</strong>.</p>
          <p>Please read the recommendation carefully, then click the button below to respond.</p>
        </div>

        <div class="section" style="border-left-color: #FF9800;">
          <h4>📝 Recommendation from ${approverRole}:</h4>
          <blockquote style="border-left: 4px solid #FF9800; margin: 10px 0; padding: 10px 20px; background: #FFF3E0; font-style: italic; color: #333;">
            ${recommendation}
          </blockquote>
        </div>

        <div class="section">
          <h4>📋 Your Request Details:</h4>
          <table>
            <tr><td class="label">Request ID:</td><td><strong>${requestId}</strong></td></tr>
            <tr><td class="label">Data Source:</td><td>${rowData[7]}</td></tr>
            <tr><td class="label">Current Stage:</td><td>${approverRole} Review</td></tr>
          </table>
        </div>

        <div class="section" style="border-left-color: #9C27B0; background-color: #F3E5F5;">
          <h4>📋 On the response page you can:</h4>
          <ul>
            <li><strong>Write your reply</strong> — answer the inquiry or describe the action you have taken</li>
            <li><strong>Confirm</strong> — submit your reply and resume the approval workflow</li>
            <li><strong>Abort</strong> — cancel the entire request if you no longer wish to proceed</li>
          </ul>
        </div>

        <div style="text-align: center; margin: 30px 0;">
          <a href="${responsePageUrl}" class="button confirm">📝 OPEN RESPONSE PAGE</a>
          <br><br>
          <a href="${statusUrl}" class="button info">📊 VIEW REQUEST STATUS</a>
        </div>

        <div class="section">
          <p><em>⚠️ Important: Your reply will be visible to the ${approverRole} when they re-review the request.</em></p>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: requesterEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// HANDLE REQUESTER CONFIRMATION
// =============================================
function handleRequesterConfirmation(token, submitAction, requesterReply) {
  try {
    const tokenData = decodeApprovalToken(token);
    if (!tokenData) {
      return createResponsePage('❌ Invalid Token', 'This confirmation link is invalid or has expired.', 'error');
    }

    if (submitAction === 'confirm' || submitAction === 'abort') {
      return processRequesterResponse(tokenData, submitAction, requesterReply || '');
    }

    const approverRole   = tokenData.role.replace('Requester-Confirmation-', '');
    const webAppUrl      = _webAppUrl;
    const encodedToken   = encodeURIComponent(token);
    const confirmPostUrl = `${webAppUrl}?action=confirm-recommendation&token=${encodedToken}&submit=confirm`;
    const abortPostUrl   = `${webAppUrl}?action=confirm-recommendation&token=${encodedToken}&submit=abort`;

    let recommendation = '';
    let requesterName  = '';
    let requestId      = tokenData.requestId;
    try {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === requestId) {
            requesterName = data[i][3];
            if (approverRole === 'Immediate Head')            recommendation = data[i][20];
            else if (approverRole === 'Data Protection Officer') recommendation = data[i][23];
            else if (approverRole.startsWith('Data Owner'))   recommendation = data[i][29];
            break;
          }
        }
      }
    } catch(e) { Logger.log('Could not fetch recommendation: ' + e); }

    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
          *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
          body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f5f5; padding: 30px 16px; }
          .card { background: white; border-radius: 14px; box-shadow: 0 4px 24px rgba(0,0,0,0.10); max-width: 680px; margin: 0 auto; overflow: hidden; }
          .card-header { background: linear-gradient(135deg, ${CONFIG.LASALLE_GREEN}, #2E7D32); color: white; padding: 24px 28px; }
          .card-header h2 { font-size: 20px; margin-bottom: 4px; }
          .card-header p  { font-size: 13px; opacity: 0.88; }
          .card-body { padding: 28px; }
          .meta { background: #f8f9fa; border-left: 4px solid ${CONFIG.LASALLE_GREEN}; padding: 12px 16px; border-radius: 0 6px 6px 0; margin-bottom: 20px; font-size: 13px; color: #444; }
          .meta strong { color: ${CONFIG.LASALLE_GREEN}; }
          .rec-box { background: #FFF8E1; border: 2px solid #FFB300; border-radius: 8px; padding: 16px 20px; margin-bottom: 24px; }
          .rec-box h4 { color: #F57F17; font-size: 13px; text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 10px; }
          .rec-box blockquote { border-left: 4px solid #FF9800; padding: 10px 16px; background: #FFF3E0; font-style: italic; color: #333; border-radius: 0 4px 4px 0; font-size: 14px; line-height: 1.6; }
          .field-label { display: block; font-weight: 600; color: #333; margin-bottom: 6px; font-size: 14px; }
          .field-label span { color: #f44336; }
          textarea { width: 100%; padding: 12px 14px; border: 2px solid #ddd; border-radius: 8px; font-size: 14px; font-family: inherit; resize: vertical; min-height: 130px; transition: border-color 0.2s; }
          textarea:focus { outline: none; border-color: ${CONFIG.LASALLE_GREEN}; }
          .char-count { font-size: 11px; color: #aaa; text-align: right; margin-top: 4px; }
          .divider { border: none; border-top: 1px solid #eee; margin: 24px 0; }
          .actions { display: flex; gap: 12px; flex-wrap: wrap; }
          .btn { flex: 1; min-width: 140px; padding: 13px 18px; border: none; border-radius: 8px; font-size: 15px; font-weight: 700; cursor: pointer; transition: all 0.18s; display: flex; align-items: center; justify-content: center; gap: 8px; }
          .btn:disabled { opacity: 0.55; cursor: not-allowed; }
          .btn-confirm { background: ${CONFIG.LASALLE_GREEN}; color: white; }
          .btn-confirm:hover:not(:disabled) { background: #2E7D32; }
          .btn-abort { background: white; color: #C62828; border: 2px solid #f44336; }
          .btn-abort:hover:not(:disabled) { background: #FFEBEE; }
          .overlay { display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 100; align-items: center; justify-content: center; padding: 20px; }
          .overlay.show { display: flex; }
          .dialog { background: white; border-radius: 12px; padding: 28px; max-width: 440px; width: 100%; box-shadow: 0 8px 32px rgba(0,0,0,0.2); }
          .dialog h3 { color: #C62828; margin-bottom: 12px; }
          .dialog p { color: #555; font-size: 14px; margin-bottom: 20px; line-height: 1.6; }
          .dialog-actions { display: flex; gap: 10px; }
          .dialog-actions .btn { font-size: 14px; padding: 10px 16px; }
          .btn-cancel-dialog { background: #f5f5f5; color: #333; border: 1px solid #ddd; }
          .notice { background: #E8F5E9; border-left: 4px solid ${CONFIG.LASALLE_GREEN}; padding: 10px 14px; border-radius: 0 6px 6px 0; font-size: 13px; color: #2E7D32; margin-bottom: 20px; }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="card-header">
            <h2>📝 Respond to Recommendation</h2>
            <p>De La Salle Lipa — Data Privacy Office</p>
          </div>
          <div class="card-body">
            <div class="meta">
              <strong>Request ID:</strong> ${requestId} &nbsp;|&nbsp;
              <strong>Pending with:</strong> ${approverRole}
              ${requesterName ? ' &nbsp;|&nbsp; <strong>Requester:</strong> ' + requesterName : ''}
            </div>
            <div class="rec-box">
              <h4>📝 Recommendation / Inquiry from ${approverRole}</h4>
              <blockquote>${recommendation || '(No recommendation text found — please check your email for the original message.)'}</blockquote>
            </div>
            <div class="notice">
              Your reply will be sent to the <strong>${approverRole}</strong> together with your confirmation.
            </div>
            <label class="field-label" for="replyText">Your Reply / Action Taken <span>*</span></label>
            <textarea id="replyText" placeholder="Describe what action you have taken..." oninput="updateChar()"></textarea>
            <div class="char-count"><span id="charCount">0</span> characters</div>
            <hr class="divider">
            <div class="actions">
              <button class="btn btn-confirm" id="confirmBtn" onclick="handleConfirm()">✅ Confirm & Resume Workflow</button>
              <button class="btn btn-abort" onclick="showAbortDialog()">🚫 Abort Request</button>
            </div>
          </div>
        </div>
        <div class="overlay" id="abortOverlay">
          <div class="dialog">
            <h3>🚫 Abort this Request?</h3>
            <p>This will <strong>permanently cancel</strong> your Data Access Request <strong>${requestId}</strong>.</p>
            <div class="dialog-actions">
              <button class="btn btn-abort" id="abortConfirmBtn" onclick="handleAbort()">Yes, Abort Request</button>
              <button class="btn btn-cancel-dialog" onclick="hideAbortDialog()">Cancel</button>
            </div>
          </div>
        </div>
        <script>
          function updateChar() { document.getElementById('charCount').textContent = document.getElementById('replyText').value.length; }
          function handleConfirm() {
            var reply = document.getElementById('replyText').value.trim();
            if (!reply) { document.getElementById('replyText').style.borderColor = '#f44336'; document.getElementById('replyText').focus(); return; }
            var btn = document.getElementById('confirmBtn');
            btn.disabled = true; btn.textContent = '⏳ Submitting...';
            window.location.href = '${confirmPostUrl}&reply=' + encodeURIComponent(reply);
          }
          function showAbortDialog() { document.getElementById('abortOverlay').classList.add('show'); }
          function hideAbortDialog() { document.getElementById('abortOverlay').classList.remove('show'); }
          function handleAbort() {
            document.getElementById('abortConfirmBtn').disabled = true;
            document.getElementById('abortConfirmBtn').textContent = '⏳ Aborting...';
            window.location.href = '${abortPostUrl}';
          }
        <\/script>
      </body>
      </html>
    `).setTitle('Respond to Recommendation — ' + requestId);

  } catch (error) {
    Logger.log('Error in handleRequesterConfirmation: ' + error.toString());
    return createResponsePage('❌ Error', 'An error occurred: ' + error.toString(), 'error');
  }
}

// =============================================
// PROCESS REQUESTER RESPONSE
// =============================================
function processRequesterResponse(tokenData, decision, requesterReply) {
  try {
    const approverRole = tokenData.role.replace('Requester-Confirmation-', '');
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data  = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tokenData.requestId) { rowIndex = i; break; }
    }

    if (rowIndex === -1) {
      return createResponsePage('❌ Request Not Found', `Request ID ${tokenData.requestId} was not found.`, 'error');
    }

    const rowData  = data[rowIndex];
    const timestamp = new Date();

    if (decision === 'abort') {
      const currentStatus = rowData[2];
      if (!currentStatus.includes('Awaiting Requester')) {
        return createResponsePage('ℹ️ Already Processed', 'This request is no longer awaiting your action.', 'info');
      }
      sheet.getRange(rowIndex + 1, 3).setValue('Withdrawn by Requester');
      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);
      logApprovalAction(tokenData.requestId, tokenData.approverEmail, 'withdrawn-by-requester', approverRole, 'Requester chose to abort');
      notifyApproverOfWithdrawal(tokenData.requestId, rowData, approverRole);
      return createResponsePage('🚫 Request Withdrawn', `Your Data Access Request <strong>${tokenData.requestId}</strong> has been withdrawn.`, 'error');
    }

    if (decision === 'confirm') {
      let approverEmail = '';
      let approverName  = '';

      if (approverRole === 'Immediate Head') {
        if (rowData[18] !== 'Pending - With Recommendation') {
          return createResponsePage('ℹ️ Already Processed', 'This confirmation has already been submitted.', 'info');
        }
        const existingRemark = rowData[20] || '';
        sheet.getRange(rowIndex + 1, 19).setValue('Pending - Requester Confirmed');
        sheet.getRange(rowIndex + 1, 21).setValue(existingRemark + '\n\n── Requester Reply ──\n' + requesterReply);
        sheet.getRange(rowIndex + 1, 3).setValue('Pending - Immediate Head Re-review');
        approverEmail = rowData[17]; approverName = rowData[16];

      } else if (approverRole === 'Data Protection Officer') {
        if (rowData[21] !== 'Pending - With Recommendation') {
          return createResponsePage('ℹ️ Already Processed', 'This confirmation has already been submitted.', 'info');
        }
        const existingRemark = rowData[23] || '';
        sheet.getRange(rowIndex + 1, 22).setValue('Pending - Requester Confirmed');
        sheet.getRange(rowIndex + 1, 24).setValue(existingRemark + '\n\n── Requester Reply ──\n' + requesterReply);
        sheet.getRange(rowIndex + 1, 3).setValue('Pending - DPO Re-review');
        approverEmail = CONFIG.DPO_EMAIL; approverName = 'Data Protection Officer';

      } else if (approverRole.startsWith('Data Owner')) {
        if (rowData[27] !== 'Pending - With Recommendation') {
          return createResponsePage('ℹ️ Already Processed', 'This confirmation has already been submitted.', 'info');
        }
        const existingRemark = rowData[29] || '';
        sheet.getRange(rowIndex + 1, 28).setValue('Pending - Requester Confirmed');
        sheet.getRange(rowIndex + 1, 30).setValue(existingRemark + '\n\n── Requester Reply ──\n' + requesterReply);
        sheet.getRange(rowIndex + 1, 3).setValue('Pending - Data Owner Re-review');
        approverEmail = rowData[26]; approverName = rowData[25] || rowData[24];
      }

      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);

      if (approverEmail) {
        sendReReviewEmail(tokenData.requestId, rowData, approverRole, approverEmail, approverName, requesterReply);
      }

      logApprovalAction(tokenData.requestId, tokenData.approverEmail, 'requester-confirmed-with-reply', approverRole, 'Requester reply: ' + requesterReply);

      return createResponsePage('✅ Reply Submitted', `Thank you! Your reply has been sent to the <strong>${approverRole}</strong>.<br><br><strong>Request ID:</strong> ${tokenData.requestId}`, 'success');
    }

    return createResponsePage('❌ Unknown Action', 'An unrecognized action was submitted.', 'error');

  } catch (error) {
    Logger.log('Error in processRequesterResponse: ' + error.toString());
    return createResponsePage('❌ Error', 'An error occurred: ' + error.toString(), 'error');
  }
}

// =============================================
// NOTIFY APPROVER OF WITHDRAWAL
// =============================================
function notifyApproverOfWithdrawal(requestId, rowData, approverRole) {
  let approverEmail = '';
  if (approverRole === 'Immediate Head')            approverEmail = rowData[17];
  else if (approverRole === 'Data Protection Officer') approverEmail = CONFIG.DPO_EMAIL;
  else if (approverRole.startsWith('Data Owner'))   approverEmail = rowData[26];
  if (!approverEmail) return;

  const subject = `[DARF] Request Withdrawn by Requester - ${requestId}`;
  const htmlBody = getEmailHeader('🚫 Request Withdrawn', 'Requester Withdrew the Request') + `
        <div class="section" style="border-left-color: #f44336;">
          <p>The requester has chosen to <strong>withdraw</strong> Data Access Request <strong>${requestId}</strong>.</p>
          <table>
            <tr><td class="label">Request ID:</td><td><strong>${requestId}</strong></td></tr>
            <tr><td class="label">Requester:</td><td>${rowData[3]} (${rowData[4]})</td></tr>
            <tr><td class="label">Data Source:</td><td>${rowData[7]}</td></tr>
          </table>
          <p style="margin-top:12px;">No further action is required.</p>
        </div>
  ` + getEmailFooter();
  MailApp.sendEmail({ to: approverEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// SEND RE-REVIEW EMAIL TO APPROVER
// =============================================
function sendReReviewEmail(requestId, rowData, approverRole, approverEmail, approverName, requesterReply) {
  const webAppUrl = _webAppUrl;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;

  let tokenRole = 'Immediate Head';
  if (approverRole === 'Data Protection Officer') tokenRole = 'Data Protection Officer';
  else if (approverRole.startsWith('Data Owner')) tokenRole = approverRole;

  const token      = generateApprovalToken(requestId, approverEmail, tokenRole);
  const approveUrl = `${webAppUrl}?action=approve&token=${encodeURIComponent(token)}`;
  const rejectUrl  = `${webAppUrl}?action=reject&token=${encodeURIComponent(token)}`;
  const pendingUrl = `${webAppUrl}?action=pending&token=${encodeURIComponent(token)}`;

  let recommendation = '';
  if (approverRole === 'Immediate Head')            recommendation = rowData[20];
  else if (approverRole === 'Data Protection Officer') recommendation = rowData[23];
  else if (approverRole.startsWith('Data Owner'))   recommendation = rowData[29];
  const splitMarker = '\n\n── Requester Reply ──\n';
  if (recommendation && recommendation.includes(splitMarker)) {
    recommendation = recommendation.split(splitMarker)[0];
  }

  const subject = `[DARF] Re-Review Required: Requester Has Responded - ${requestId}`;

  const htmlBody = getEmailHeader('🔄 Re-Review Required', approverRole + ' — Requester Has Responded') + `
        <div class="section" style="border-left-color: #9C27B0; background-color: #F3E5F5;">
          <h3>Dear ${approverName},</h3>
          <p>The requester for Data Access Request <strong>${requestId}</strong> has <strong>replied to your recommendation</strong>.</p>
        </div>

        <div class="section" style="border-left-color: #FF9800;">
          <h4>📝 Your Original Recommendation:</h4>
          <blockquote style="border-left: 4px solid #FF9800; margin: 10px 0; padding: 10px 20px; background: #FFF3E0; font-style: italic; color: #333;">
            ${recommendation || '(see original email)'}
          </blockquote>
        </div>

        <div class="section" style="border-left-color: #4CAF50; background-color: #F1F8E9;">
          <h4>💬 Requester's Reply:</h4>
          <blockquote style="border-left: 4px solid #4CAF50; margin: 10px 0; padding: 10px 20px; background: #E8F5E9; color: #1B5E20; font-size: 14px; line-height: 1.6;">
            ${requesterReply || '(No reply text provided)'}
          </blockquote>
        </div>

        <div class="section">
          <h4>📋 Request Details:</h4>
          <table>
            <tr><td class="label">Request ID:</td><td><strong>${requestId}</strong></td></tr>
            <tr><td class="label">Requester:</td><td>${rowData[3]} (${rowData[4]})</td></tr>
            <tr><td class="label">Data Source:</td><td>${rowData[7]}</td></tr>
          </table>
        </div>

        ${buildApprovalButtons(approveUrl, rejectUrl, pendingUrl, statusUrl)}
  ` + getEmailFooter();

  MailApp.sendEmail({ to: approverEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// HANDLE APPROVAL/REJECTION ACTIONS
// =============================================
function handleApprovalAction(token, action) {
  try {
    const tokenData = decodeApprovalToken(token);
    if (!tokenData) {
      return createResponsePage('❌ Invalid Token', 'This approval link is invalid or has expired.', 'error');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tokenData.requestId) { rowIndex = i; break; }
    }

    if (rowIndex === -1) {
      return createResponsePage('❌ Request Not Found', `Request ID ${tokenData.requestId} was not found.`, 'error');
    }

    const timestamp = new Date();
    const actionText = action === 'approve' ? 'Endorsed' : 'Disapproved';
    const rowData = data[rowIndex];

    if (tokenData.role === 'Immediate Head') {
      const currentStatus = rowData[18];
      if (currentStatus !== 'Pending' && currentStatus !== 'Pending - Requester Confirmed') {
        return createResponsePage('ℹ️ Already Processed', 'This request has already been acted upon by the Immediate Head.', 'info');
      }
      sheet.getRange(rowIndex + 1, 19).setValue(actionText);
      sheet.getRange(rowIndex + 1, 20).setValue(timestamp);
      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);

      if (action === 'approve') {
        sheet.getRange(rowIndex + 1, 3).setValue('Pending - DPO Review');
        const updatedRow = sheet.getRange(rowIndex + 1, 1, 1, 42).getValues()[0];
        sendDPOApprovalEmail(tokenData.requestId, updatedRow);
        notifyRequesterProgress(tokenData.requestId, rowData, 'Immediate Head', 'endorsed', 'Data Protection Officer');
      } else {
        sheet.getRange(rowIndex + 1, 3).setValue('Disapproved by Immediate Head');
        notifyRequesterOfRejection(tokenData.requestId, rowData, 'Immediate Head');
      }

    } else if (tokenData.role === 'Data Protection Officer') {
      const currentStatus = rowData[21];
      if (currentStatus !== 'Pending' && currentStatus !== 'Pending - Requester Confirmed') {
        return createResponsePage('ℹ️ Already Processed', 'This request has already been evaluated by the DPO.', 'info');
      }
      sheet.getRange(rowIndex + 1, 22).setValue(action === 'approve' ? 'Approved' : 'Not Cleared');
      sheet.getRange(rowIndex + 1, 23).setValue(timestamp);
      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);

      if (action === 'approve') {
        sheet.getRange(rowIndex + 1, 3).setValue('Pending - Data Owner Endorsement');
        const updatedRow = sheet.getRange(rowIndex + 1, 1, 1, 42).getValues()[0];
        sendDataOwnerApprovalEmail(tokenData.requestId, updatedRow);
        notifyRequesterProgress(tokenData.requestId, rowData, 'Data Protection Officer', 'approved', 'Data Owner (' + rowData[24] + ')');
      } else {
        sheet.getRange(rowIndex + 1, 3).setValue('Not Cleared by DPO');
        notifyRequesterOfRejection(tokenData.requestId, rowData, 'Data Protection Officer');
      }

    } else if (tokenData.role.startsWith('Data Owner')) {
      const currentStatus = rowData[27];
      if (currentStatus !== 'Pending' && currentStatus !== 'Pending - Requester Confirmed') {
        return createResponsePage('ℹ️ Already Processed', 'This request has already been acted upon by the Data Owner.', 'info');
      }
      sheet.getRange(rowIndex + 1, 28).setValue(actionText);
      sheet.getRange(rowIndex + 1, 29).setValue(timestamp);
      sheet.getRange(rowIndex + 1, 41).setValue(timestamp);

      if (action === 'approve') {
        sheet.getRange(rowIndex + 1, 3).setValue('Approved - Pending Data Processing');
        sheet.getRange(rowIndex + 1, 33).setValue('Pending'); // Processor Status — required for Grant Access button
        const updatedRow = sheet.getRange(rowIndex + 1, 1, 1, 42).getValues()[0];
        sendDataProcessorNotification(tokenData.requestId, updatedRow);
        notifyRequesterProgress(tokenData.requestId, rowData, 'Data Owner (' + rowData[24] + ')', 'endorsed', 'Data Processor');
      } else {
        sheet.getRange(rowIndex + 1, 3).setValue('Disapproved by Data Owner');
        notifyRequesterOfRejection(tokenData.requestId, rowData, 'Data Owner (' + rowData[24] + ')');
      }
    }

    logApprovalAction(tokenData.requestId, tokenData.approverEmail, action, tokenData.role);

    return createResponsePage(
      `${action === 'approve' ? '✅' : '❌'} Request ${actionText}`,
      `You have successfully <strong>${actionText.toLowerCase()}</strong> the data access request.<br><br>
       <strong>Request ID:</strong> ${tokenData.requestId}<br>
       <strong>Your Role:</strong> ${tokenData.role}<br>
       <strong>Timestamp:</strong> ${timestamp}<br><br>
       ${action === 'approve' ? 'The request has been forwarded to the next approver.' : 'The requester has been notified.'}`,
      action === 'approve' ? 'success' : 'error'
    );

  } catch (error) {
    Logger.log('Error in handleApprovalAction: ' + error.toString());
    return createResponsePage('❌ Error', 'An error occurred: ' + error.toString(), 'error');
  }
}

// =============================================
// HANDLE DATA PROCESSOR GRANTING ACCESS
// =============================================
function handleGrantAccess(token) {
  try {
    const tokenData = decodeApprovalToken(token);
    if (!tokenData || tokenData.role !== 'Data Processor') {
      return createResponsePage('❌ Invalid Token', 'This link is invalid or has expired.', 'error');
    }

    const webAppUrl = _webAppUrl;
    const dashboardUrl = `${webAppUrl}?page=dashboard`;

    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html><html><head>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; text-align: center; padding: 50px 20px; background-color: #f5f5f5; }
          .container { background: white; padding: 40px; border-radius: 15px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); max-width: 600px; margin: 0 auto; }
          h2 { color: ${CONFIG.LASALLE_GREEN}; margin-bottom: 20px; }
          .message { color: #555; line-height: 1.8; margin-bottom: 20px; }
          .button { display: inline-block; padding: 14px 30px; background: ${CONFIG.LASALLE_GREEN}; color: white; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; }
        </style>
      </head><body>
        <div class="container">
          <h2>⚙️ Data Processor Action Required</h2>
          <div class="message">
            <p>Request <strong>${tokenData.requestId}</strong> requires your action.</p>
            <p>Please use the <strong>DARF Dashboard</strong> to set the expected timeline, add comments, and grant access.</p>
          </div>
          <a href="${dashboardUrl}" class="button">📊 Open DARF Dashboard</a>
        </div>
      </body></html>
    `).setTitle('Data Processor - Grant Access');

  } catch (error) {
    Logger.log('Error in handleGrantAccess: ' + error.toString());
    return createResponsePage('❌ Error', 'An error occurred: ' + error.toString(), 'error');
  }
}

// =============================================
// NOTIFICATION EMAILS
// =============================================
function notifyRequesterProgress(requestId, rowData, approvedBy, actionWord, nextStep) {
  const requesterEmail = rowData[4];
  const requesterName = rowData[3];
  const webAppUrl = _webAppUrl;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;

  const subject = `[DARF] Progress Update - ${requestId}`;
  const htmlBody = getEmailHeader('📊 Request Progress Update') + `
        <div class="section">
          <h3>Dear ${requesterName},</h3>
          <p>Your Data Access Request <strong>${requestId}</strong> has been <strong>${actionWord}</strong> by the <strong>${approvedBy}</strong>.</p>
          <p>The request has been forwarded to the <strong>${nextStep}</strong> for the next step.</p>
        </div>
        <div style="text-align: center; margin: 30px 0;">
          <a href="${statusUrl}" class="button info">📊 TRACK REQUEST STATUS</a>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: requesterEmail, subject: subject, htmlBody: htmlBody });
}

function notifyRequesterAccessGranted(requestId, rowData) {
  const requesterEmail = rowData[4];
  const requesterName = rowData[3];
  const webAppUrl = _webAppUrl;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;

  const expectedTimeline = rowData[36] || 'As communicated by the Data Processor';
  const processorComments = rowData[35] || '';
  const accessMethod = rowData[34] || 'Info Hub / As directed by Data Processor';

  const subject = `[DARF] ✅ Data Access Granted - ${requestId}`;
  const htmlBody = getEmailHeader('✅ Data Access Granted', 'Your requested data is now available') + `
        ${getProgressTracker('complete')}
        <div class="section">
          <h3>Dear ${requesterName},</h3>
          <p>Your Data Access Request <strong>${requestId}</strong> has been <strong>fully processed</strong> and access has been granted.</p>
          <table>
            <tr><td class="label">Data Source:</td><td>${rowData[7]}</td></tr>
            <tr><td class="label">Access Period:</td><td>${rowData[13]} to ${rowData[14]}</td></tr>
            <tr><td class="label">Access Method:</td><td><strong>${accessMethod}</strong></td></tr>
            <tr><td class="label">Expected Timeline:</td><td><strong>${expectedTimeline}</strong></td></tr>
          </table>
        </div>
        ${processorComments ? `
        <div class="section" style="border-left-color: #1976D2;">
          <h4>💬 Message from the Data Processor:</h4>
          <p style="font-style: italic; color: #333;">${processorComments}</p>
        </div>` : ''}
        <div class="section" style="border-left-color: #FF9800; background-color: #FFF3E0;">
          <h4>⚠️ Important Reminders:</h4>
          <ul>
            <li>Access is valid only for the stated period</li>
            <li>Adhere to the security measures you specified</li>
            <li>Do not share the data with unauthorized persons</li>
            <li>Ensure proper disposal of data after the access period ends</li>
          </ul>
        </div>
        <div style="text-align: center; margin: 30px 0;">
          <a href="${statusUrl}" class="button info">📊 VIEW REQUEST DETAILS</a>
        </div>
  ` + getEmailFooter();

  MailApp.sendEmail({ to: requesterEmail, subject: subject, htmlBody: htmlBody });
}

function notifyRequesterOfRejection(requestId, rowData, rejectedBy) {
  const requesterEmail = rowData[4];
  const requesterName = rowData[3];
  const webAppUrl = _webAppUrl;
  const statusUrl = `${webAppUrl}?page=status&id=${encodeURIComponent(requestId)}`;

  const subject = `[DARF] Request ${requestId} - Action Required`;
  const htmlBody = `
    <!DOCTYPE html><html><head>
    <style>
      body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
      .header { background-color: #f44336; color: white; padding: 20px; text-align: center; }
      .content { padding: 20px; }
      .section { background-color: #f9f9f9; padding: 15px; margin: 10px 0; border-left: 4px solid #f44336; border-radius: 4px; }
      .button { display: inline-block; padding: 12px 24px; margin: 10px 5px; text-decoration: none; border-radius: 5px; font-weight: bold; background-color: #2196F3; color: white; }
      .footer { text-align: center; padding: 20px; font-size: 12px; color: #666; }
    </style>
    </head><body>
      <div class="header"><h2>❌ Request Disapproved</h2></div>
      <div class="content">
        <div class="section">
          <h3>Dear ${requesterName},</h3>
          <p>Your Data Access Request <strong>${requestId}</strong> has been disapproved by the <strong>${rejectedBy}</strong>.</p>
          <p>Please review your request and resubmit with the necessary corrections or additional information.</p>
        </div>
        <div style="text-align: center; margin: 30px 0;">
          <a href="${statusUrl}" class="button">📊 VIEW REQUEST DETAILS</a>
        </div>
        <div class="section">
          <p>For questions, please contact <a href="mailto:${CONFIG.DPO_EMAIL}">${CONFIG.DPO_EMAIL}</a></p>
        </div>
      </div>
      <div class="footer"><p>This is an automated message from the DLSL Data Privacy Office.</p></div>
    </body></html>
  `;

  MailApp.sendEmail({ to: requesterEmail, subject: subject, htmlBody: htmlBody });
}

// =============================================
// UTILITY FUNCTIONS
// =============================================
function logApprovalAction(requestId, approverEmail, action, role, comments) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.APPROVAL_SHEET);
  if (!sheet) {
    initializeSpreadsheet();
    sheet = ss.getSheetByName(CONFIG.APPROVAL_SHEET);
  }
  sheet.appendRow([requestId, new Date(), approverEmail, action, role, comments || '']);
}

function createResponsePage(title, message, type) {
  const iconColor = type === 'success' ? '#4CAF50' : type === 'error' ? '#f44336' : '#2196F3';
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; text-align: center; padding: 50px 20px; background-color: #f5f5f5; }
        .container { background: white; padding: 40px; border-radius: 15px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); max-width: 600px; margin: 0 auto; }
        h2 { color: #1B5E20; margin-bottom: 20px; }
        .message { color: #555; line-height: 1.8; }
        .icon-bar { width: 60px; height: 4px; background: ${iconColor}; margin: 20px auto; border-radius: 2px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>${title}</h2>
        <div class="icon-bar"></div>
        <div class="message">${message}</div>
      </div>
    </body>
    </html>
  `).setTitle(title.replace(/[^\w\s-]/g, ''));
}

function showStatus(requestId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return HtmlService.createHtmlOutput('<h2>System not initialized</h2>');

  const data = sheet.getDataRange().getValues();
  let request = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) { request = data[i]; break; }
  }

  if (!request) return HtmlService.createHtmlOutput('<h2>Request not found</h2>');

  const formatDate = (date) => {
    if (!date) return '';
    if (date instanceof Date) return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');
    try {
      const dateObj = new Date(date);
      if (!isNaN(dateObj.getTime())) return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');
    } catch (e) {}
    return date.toString();
  };

  const template = HtmlService.createTemplateFromFile('Status');
  template.request = request;
  template.requestId = requestId;
  template.LASALLE_GREEN = CONFIG.LASALLE_GREEN;
  template.formatDate = formatDate;

  return template.evaluate()
    .setTitle('Request Status - ' + requestId)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCurrentUserEmail() {
  try {
    let email = Session.getActiveUser().getEmail();
    if (!email) email = Session.getEffectiveUser().getEmail();
    return email || '';
  } catch(e) {
    Logger.log('Error getting email: ' + e.toString());
    return '';
  }
}

// =============================================
// DASHBOARD - Restricted to DPO & Data Processors
// =============================================
function isDashboardAuthorized() {
  try {
    const email = (getCurrentUserEmail() || '').toLowerCase();
    if (!email) return false;
    if (CONFIG.DPO_EMAIL && email === CONFIG.DPO_EMAIL.toLowerCase()) return true;
    if (CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_EMAIL && email === CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_EMAIL.toLowerCase()) return true;
    if (CONFIG.BULK_REQUEST.DATA_PROCESSOR_EMAIL && email === CONFIG.BULK_REQUEST.DATA_PROCESSOR_EMAIL.toLowerCase()) return true;
    return false;
  } catch(e) {
    Logger.log('Auth check error: ' + e.toString());
    return false;
  }
}

function getDashboardUserRole() {
  const email = getCurrentUserEmail().toLowerCase();
  if (!email) return null;
  const isDPO = CONFIG.DPO_EMAIL && email === CONFIG.DPO_EMAIL.toLowerCase();
  const isSingleProcessor = CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_EMAIL && email === CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_EMAIL.toLowerCase();
  const isBulkProcessor = CONFIG.BULK_REQUEST.DATA_PROCESSOR_EMAIL && email === CONFIG.BULK_REQUEST.DATA_PROCESSOR_EMAIL.toLowerCase();

  let roles = [];
  if (isDPO) roles.push('DPO');
  if (isSingleProcessor) roles.push('Single Request Processor');
  if (isBulkProcessor) roles.push('Bulk Request Processor');

  if (roles.length === 0) return null;
  return roles.join(' & ');
}

function checkDashboardRole() {
  const email = getCurrentUserEmail();
  const role  = getDashboardUserRole();
  Logger.log('Email detected : ' + email);
  Logger.log('Role detected  : ' + role);
  Logger.log('DPO_EMAIL      : ' + CONFIG.DPO_EMAIL);
  Logger.log('Single Proc    : ' + CONFIG.SINGLE_REQUEST.DATA_PROCESSOR_EMAIL);
  Logger.log('Bulk Proc      : ' + CONFIG.BULK_REQUEST.DATA_PROCESSOR_EMAIL);
  return { email, role };
}

function serveDashboard() {
  const email = getCurrentUserEmail().toLowerCase();
  const isAuthorized = isDashboardAuthorized();
  if (!isAuthorized) {
    return createResponsePage(
      '🔒 Access Denied',
      'This dashboard is restricted to authorized personnel only.<br><br>Your email: <code>' + (email || 'unable to detect') + '</code>',
      'error'
    );
  }
  const template = HtmlService.createTemplateFromFile('Dashboard');
  return template.evaluate()
    .setTitle('DARF Dashboard - ' + CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================
// HELPER: Normalize any date value to YYYY-MM-DD
// =============================================
function normalizeDateToYMD(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  } catch(e) {}
  return String(val);
}

// =============================================
// DASHBOARD DATA
// =============================================
function getDashboardRequests() {
  try {
    if (!isDashboardAuthorized()) {
      return { success: false, message: 'Access denied. Email: ' + (getCurrentUserEmail() || 'unknown') };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { success: true, requests: [], role: getDashboardUserRole() };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const requests = [];

    const colIndex = {};
    for (let c = 0; c < headers.length; c++) {
      colIndex[String(headers[c]).trim()] = c;
    }

    function getVal(row, headerName) {
      const idx = colIndex[headerName];
      if (idx === undefined || idx >= row.length) return '';
      const v = row[idx];
      return (v === null || v === undefined) ? '' : v;
    }
    function getDateVal(row, headerName) {
      const val = getVal(row, headerName);
      if (!val) return '';
      try {
        if (val instanceof Date) return val.toISOString();
        const d = new Date(val);
        if (!isNaN(d.getTime())) return d.toISOString();
      } catch(e) {}
      return String(val);
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      // Normalize dates to YYYY-MM-DD for reliable frontend filtering
      var normalizedDateOfRequest = normalizeDateToYMD(getVal(row, 'Date of Request'));
      var normalizedTimestamp = normalizeDateToYMD(getVal(row, 'Timestamp'));
      var normalizedAccessStart = normalizeDateToYMD(getVal(row, 'Access Start Date'));
      var normalizedAccessEnd = normalizeDateToYMD(getVal(row, 'Access End Date'));

      requests.push({
        rowIndex: i,
        requestId: String(getVal(row, 'Request ID')),
        timestamp: getDateVal(row, 'Timestamp'),
        timestampDate: normalizedTimestamp,
        status: String(getVal(row, 'Status')),
        requesterName: String(getVal(row, 'Requester Name')),
        requesterEmail: String(getVal(row, 'Requester Email')),
        requesterOffice: String(getVal(row, 'Requester Office/Course')),
        dateOfRequest: normalizedDateOfRequest || normalizedTimestamp,
        dataSourceCategory: String(getVal(row, 'Data Source Category')),
        personalDataRequested: String(getVal(row, 'Personal Data Requested')),
        internalRecipients: String(getVal(row, 'Internal Recipients')),
        externalRecipients: String(getVal(row, 'External Recipients')),
        accessStartDate: normalizedAccessStart,
        accessEndDate: normalizedAccessEnd,
        securityMeasures: String(getVal(row, 'Security Measures')),
        immediateHeadName: String(getVal(row, 'Immediate Head Name')),
        immediateHeadStatus: String(getVal(row, 'Immediate Head Status')),
        immediateHeadDate: getDateVal(row, 'Immediate Head Date'),
        immediateHeadRemarks: String(getVal(row, 'Immediate Head Remarks')),
        dpoStatus: String(getVal(row, 'DPO Status')),
        dpoDate: getDateVal(row, 'DPO Date'),
        dpoRemarks: String(getVal(row, 'DPO Remarks')),
        dataOwnerTitle: String(getVal(row, 'Data Owner Title')),
        dataOwnerStatus: String(getVal(row, 'Data Owner Status')),
        dataOwnerDate: getDateVal(row, 'Data Owner Date'),
        dataOwnerRemarks: String(getVal(row, 'Data Owner Remarks')),
        dataProcessorName: String(getVal(row, 'Data Processor Name')),
        dataProcessorEmail: String(getVal(row, 'Data Processor Email')),
        dataProcessorStatus: String(getVal(row, 'Data Processor Status')),
        dataProcessorDate: getDateVal(row, 'Data Processor Date'),
        accessMethod: String(getVal(row, 'Access Method')),
        processorComments: String(getVal(row, 'Processor Comments')),
        expectedTimeline: String(getVal(row, 'Expected Timeline')),
        requestType: String(getVal(row, 'Request Type') || 'single')
      });
    }

    return {
      success: true,
      requests: requests,
      role: getDashboardUserRole(),
      userEmail: getCurrentUserEmail()
    };
  } catch (error) {
    Logger.log('Error in getDashboardRequests: ' + error.toString());
    return { success: false, message: 'Error loading data: ' + error.toString() };
  }
}

// =============================================
// GRANT ACCESS FROM DASHBOARD
// =============================================
function grantAccessFromDashboard(requestId, expectedTimeline, processorComments, accessMethod) {
  if (!isDashboardAuthorized()) return { success: false, message: 'Access denied' };

  const role = getDashboardUserRole();
  if (!role || (!role.includes('Processor') && !role.includes('DPO'))) {
    return { success: false, message: 'Only authorized processors can grant access.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === requestId) { rowIndex = i; break; }
    }

    if (rowIndex === -1) return { success: false, message: 'Request not found.' };

    const rowData = data[rowIndex];

    if (rowData[27] !== 'Endorsed') {
      return { success: false, message: 'This request has not yet been endorsed by the Data Owner.' };
    }
    if (rowData[32] !== 'Pending') {
      return { success: false, message: 'This request has already been processed.' };
    }

    const currentEmail = getCurrentUserEmail().toLowerCase();
    const assignedProcessorEmail = (rowData[31] || '').toLowerCase();
    const isDPO = CONFIG.DPO_EMAIL && currentEmail === CONFIG.DPO_EMAIL.toLowerCase();

    if (!isDPO && currentEmail !== assignedProcessorEmail) {
      return { success: false, message: 'You are not the assigned Data Processor for this request. Assigned: ' + rowData[30] };
    }

    const timestamp = new Date();

    sheet.getRange(rowIndex + 1, 33).setValue('Access Granted');
    sheet.getRange(rowIndex + 1, 34).setValue(timestamp);
    sheet.getRange(rowIndex + 1, 35).setValue(accessMethod || 'Info Hub');
    sheet.getRange(rowIndex + 1, 36).setValue(processorComments);
    sheet.getRange(rowIndex + 1, 37).setValue(expectedTimeline);
    sheet.getRange(rowIndex + 1, 3).setValue('Complete - Access Granted');
    sheet.getRange(rowIndex + 1, 41).setValue(timestamp);

    logApprovalAction(requestId, Session.getActiveUser().getEmail(), 'grant-access', 'Data Processor', processorComments);

    const totalCols = sheet.getLastColumn();
    notifyRequesterAccessGranted(requestId, sheet.getRange(rowIndex + 1, 1, 1, totalCols).getValues()[0]);

    return { success: true, message: 'Access granted and requester notified.' };

  } catch (error) {
    Logger.log('Error in grantAccessFromDashboard: ' + error.toString());
    return { success: false, message: 'Error: ' + error.toString() };
  }
}

// =============================================
// EMAIL DASHBOARD REPORT
// =============================================
function emailDashboardReport(to, subject, message, filter) {
  if (!isDashboardAuthorized()) {
    return { success: false, message: 'Access denied' };
  }

  try {
    var validation = validateDLSLEmail(to);
    if (!validation.valid) {
      return { success: false, message: validation.message };
    }

    var result = getDashboardRequests();
    if (!result.success) return { success: false, message: 'Could not load requests' };

    var requests = result.requests || [];

    // Apply same filter logic as frontend
    if (filter === 'complete') {
      requests = requests.filter(function(r) { return r.status.indexOf('Complete') >= 0; });
    } else if (filter === 'rejected') {
      requests = requests.filter(function(r) {
        return r.status.indexOf('Disapproved') >= 0 || r.status.indexOf('Not Cleared') >= 0 || r.status.indexOf('Withdrawn') >= 0;
      });
    } else if (filter === 'pending') {
      requests = requests.filter(function(r) {
        return r.status.indexOf('Pending') >= 0 && r.status.indexOf('Disapproved') < 0 && r.status.indexOf('Awaiting') < 0;
      });
    } else if (filter === 'waiting') {
      requests = requests.filter(function(r) { return r.status.indexOf('Awaiting Requester') >= 0; });
    } else if (filter === 'action') {
      var role = getDashboardUserRole() || '';
      requests = requests.filter(function(r) {
        if ((role.indexOf('Processor') >= 0 || role.indexOf('DPO') >= 0) && r.status === 'Approved - Pending Data Processing') return true;
        if (role.indexOf('DPO') >= 0 && r.status === 'Pending - DPO Review') return true;
        return false;
      });
    }

    var rows = requests.map(function(r) {
      return '<tr>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e0e0e0;">' + r.requestId + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e0e0e0;">' + r.requesterName + '<br><small style="color:#999;">' + r.requesterEmail + '</small></td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e0e0e0;">' + (r.dateOfRequest || '') + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e0e0e0;">' + r.status + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e0e0e0;">' + (r.dataSourceCategory || '—') + '</td>' +
        '</tr>';
    }).join('');

    var filterLabel = filter || 'all';
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM dd, yyyy hh:mm a');

    var htmlBody = getEmailHeader('📊 DARF Dashboard Report') +
      (message ? '<div class="section"><p>' + message + '</p></div>' : '') +
      '<div class="section">' +
      '<p><strong>Filter:</strong> ' + filterLabel +
      ' &nbsp;·&nbsp; <strong>Records:</strong> ' + requests.length +
      ' &nbsp;·&nbsp; <strong>Generated:</strong> ' + now + '</p>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
      '<thead><tr>' +
      '<th style="background:#1B5E20;color:white;padding:8px 10px;text-align:left;">Request ID</th>' +
      '<th style="background:#1B5E20;color:white;padding:8px 10px;text-align:left;">Requester</th>' +
      '<th style="background:#1B5E20;color:white;padding:8px 10px;text-align:left;">Date</th>' +
      '<th style="background:#1B5E20;color:white;padding:8px 10px;text-align:left;">Status</th>' +
      '<th style="background:#1B5E20;color:white;padding:8px 10px;text-align:left;">Data Source</th>' +
      '</tr></thead><tbody>' +
      (rows || '<tr><td colspan="5" style="padding:20px;text-align:center;color:#999;">No records found for this filter.</td></tr>') +
      '</tbody></table>' +
      getEmailFooter();

    MailApp.sendEmail({
      to: to,
      subject: subject || 'DARF Dashboard Report',
      htmlBody: htmlBody
    });

    return { success: true, message: 'Report sent to ' + to };
  } catch (error) {
    Logger.log('Error in emailDashboardReport: ' + error.toString());
    return { success: false, message: 'Error sending report: ' + error.toString() };
  }
}

// =============================================
// UTILITY: Resend processor notification
// Run once from editor to resend to data processor
// =============================================
function checkWebAppUrl() {
  var url = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  Logger.log('WEB_APP_URL = ' + url);
  Logger.log('ScriptApp URL = ' + ScriptApp.getService().getUrl());
}

function resendProcessorNotification() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName(CONFIG.SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'DARF-B-20260228-184120') {
      var row = sheet.getRange(i + 1, 1, 1, 42).getValues()[0];
      sendDataProcessorNotification('DARF-B-20260228-184120', row);
      Logger.log('Notification resent to: ' + row[31]);
      return;
    }
  }
  Logger.log('Request not found.');
}

