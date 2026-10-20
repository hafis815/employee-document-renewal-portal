/**
 * Employee Document Portal — Multi-company Expiry Alerts
 * Sends each company's document-expiry alerts to that company's active HR users.
 * Active Super Admin accounts may also receive global alerts.
 */
const EXPIRY_ALERT_DAYS = [90, 60, 30, 15, 7, 0];

function checkDocumentExpiryAlerts() {
  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const documentSheet = ss.getSheetByName('Documents');
  const rolesSheet = ss.getSheetByName('Roles');
  const employeesSheet = ss.getSheetByName('Employees');
  const companiesSheet = ss.getSheetByName('Companies');

  if (!documentSheet || !rolesSheet || !employeesSheet || !companiesSheet) {
    throw new Error('A required portal sheet was not found.');
  }

  const documents = documentSheet.getDataRange().getValues();
  const roles = rolesSheet.getDataRange().getValues();
  const employeeRows = employeesSheet.getDataRange().getValues();
  const companyRows = companiesSheet.getDataRange().getValues();
  const activeCompanies = {};

  companyRows.slice(1).forEach(function(row) {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const status = String(row[2] || '').trim().toUpperCase();
    if (id && name && status === 'ACTIVE') activeCompanies[id] = name;
  });

  const employees = {};
  employeeRows.slice(1).forEach(function(row) {
    const id = String(row[0] || '').trim();
    if (!id) return;
    employees[id] = {
      companyId: String(row[6] || '').trim(),
      active: row[4] === true || String(row[4] || '').trim().toUpperCase() === 'TRUE'
    };
  });

  const hrByCompany = {};
  const superAdmins = [];
  roles.slice(1).forEach(function(row) {
    const email = String(row[0] || '').trim().toLowerCase();
    const role = String(row[1] || '').trim().toUpperCase();
    const active = row[2] === true || String(row[2] || '').trim().toUpperCase() === 'TRUE';
    const companyId = String(row[3] || '').trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !active) return;

    if (role === 'SUPER_ADMIN') {
      superAdmins.push(email);
    } else if (role === 'HR' && activeCompanies[companyId]) {
      if (!hrByCompany[companyId]) hrByCompany[companyId] = [];
      hrByCompany[companyId].push(email);
    }
  });

  const timezone = Session.getScriptTimeZone();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const properties = PropertiesService.getScriptProperties();
  let sentCount = 0;
  const enabledAlertDays = getEnabledExpiryAlertDays_();

  for (let i = 1; i < documents.length; i++) {
    const row = documents[i];
    const documentId = String(row[0] || '').trim();
    const employeeId = String(row[1] || '').trim();
    const type = String(row[2] || '').trim();
    const name = String(row[3] || '').trim();
    const expiryValue = row[5];
    if (!documentId || !employeeId || !(expiryValue instanceof Date)) continue;

    const employee = employees[employeeId];
    if (!employee || !employee.active || !employee.companyId || !activeCompanies[employee.companyId]) continue;
    const companyId = employee.companyId;
    const hrRecipients = hrByCompany[companyId] || [];
    const recipients = Array.from(new Set(hrRecipients.concat(superAdmins)));
    if (!recipients.length) continue;

    const expiryDate = new Date(expiryValue);
    if (isNaN(expiryDate.getTime())) continue;
    expiryDate.setHours(0, 0, 0, 0);
    const daysRemaining = Math.round((expiryDate.getTime() - today.getTime()) / 86400000);
if (!enabledAlertDays.includes(daysRemaining)) continue;

    const expiryKey = Utilities.formatDate(expiryDate, timezone, 'yyyy-MM-dd');
    // Keep the legacy key format to avoid resending alerts already sent before this update.
    const propertyKey = 'EXPIRY_ALERT_' + documentId + '_' + expiryKey + '_' + daysRemaining;
    if (properties.getProperty(propertyKey)) continue;

    const subject = daysRemaining === 0
      ? 'EXPIRED TODAY: ' + name
      : 'Document expiry in ' + daysRemaining + ' days: ' + name;
    const expiryText = Utilities.formatDate(expiryDate, timezone, 'dd MMM yyyy');
    const message =
      'Employee Document Portal\n\n' +
      'A document requires your attention.\n\n' +
      'Company: ' + activeCompanies[companyId] + '\n' +
      'Employee ID: ' + employeeId + '\n' +
      'Document: ' + name + '\n' +
      'Document type: ' + type + '\n' +
      'Expiry date: ' + expiryText + '\n' +
      'Days remaining: ' + daysRemaining + '\n\n' +
      'Please review this document and arrange renewal if needed.';

    recipients.forEach(function(email) {
      MailApp.sendEmail(email, subject, message);
    });
    properties.setProperty(propertyKey, new Date().toISOString());
    sentCount++;
    console.log('Expiry alert sent for ' + documentId + ' (' + companyId + '), days remaining: ' + daysRemaining);
  }

  console.log('Expiry alert check complete. Alerts sent: ' + sentCount);

try {
  checkFamilyDocumentExpiryAlerts();
} catch (error) {
  console.error('Family expiry alert check failed: ' + error.message);
}

try {
  checkCompanyDocumentExpiryAlerts();
} catch (error) {
  console.error('Company document expiry alert check failed: ' + error.message);
}
}

function createExpiryAlertTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkDocumentExpiryAlerts') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('checkDocumentExpiryAlerts').timeBased().everyDays(1).atHour(8).create();
  return 'Daily expiry alert trigger created.';
}

function testExpiryEmailToSelf() {
  const user = getCurrentUser();

  if (!user || !isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can run this test.');
  }

  const recipient = String(user.email || '').trim();

  if (!recipient || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient)) {
    throw new Error('A valid Super Admin email address is required.');
  }

  MailApp.sendEmail({
    to: recipient,
    subject: 'Employee Document Portal — Expiry Email Test',
    body:
      'This is a test email from the Employee Document Portal.\n\n' +
      'If you received this message, Apps Script email delivery is working.\n\n' +
      'This test does not modify any document records or expiry-alert settings.'
  });

  console.log('Test expiry email sent to ' + recipient);
}

function checkFamilyDocumentExpiryAlerts() {
  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const familyDocsSheet = ss.getSheetByName('FamilyDocuments');
  const familyMembersSheet = ss.getSheetByName('FamilyMembers');
  const employeesSheet = ss.getSheetByName('Employees');
  const companiesSheet = ss.getSheetByName('Companies');
  const rolesSheet = ss.getSheetByName('Roles');

  if (
    !familyDocsSheet || !familyMembersSheet ||
    !employeesSheet || !companiesSheet || !rolesSheet
  ) {
    throw new Error('A required portal sheet was not found.');
  }

  const familyDocs = familyDocsSheet.getDataRange().getValues();
  const familyMembers = familyMembersSheet.getDataRange().getValues();
  const employeeRows = employeesSheet.getDataRange().getValues();
  const companyRows = companiesSheet.getDataRange().getValues();
  const roles = rolesSheet.getDataRange().getValues();

  const activeCompanies = {};
  companyRows.slice(1).forEach(function(row) {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const status = String(row[2] || '').trim().toUpperCase();
    if (id && name && status === 'ACTIVE') activeCompanies[id] = name;
  });

  const employees = {};
  employeeRows.slice(1).forEach(function(row) {
    const id = String(row[0] || '').trim();
    if (!id) return;
    employees[id] = {
      name: String(row[1] || '').trim(),
      companyId: String(row[6] || '').trim(),
      active: row[4] === true ||
        String(row[4] || '').trim().toUpperCase() === 'TRUE'
    };
  });

  // FamilyMembers must link each family member to an employee.
  const memberById = {};
  familyMembers.slice(1).forEach(function(row) {
    const id = String(row[0] || '').trim();
    if (!id) return;
    memberById[id] = {
      employeeId: String(row[1] || '').trim(),
      name: String(row[2] || '').trim(),
     active: row[5] === true ||
  String(row[5] || '').trim().toUpperCase() === 'TRUE'
    };
  });

  const hrByCompany = {};
  const superAdmins = [];

  roles.slice(1).forEach(function(row) {
    const email = String(row[0] || '').trim().toLowerCase();
    const role = String(row[1] || '').trim().toUpperCase();
    const active = row[2] === true ||
      String(row[2] || '').trim().toUpperCase() === 'TRUE';
    const companyId = String(row[3] || '').trim();

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !active) return;

    if (role === 'SUPER_ADMIN') {
      superAdmins.push(email);
    } else if (role === 'HR' && activeCompanies[companyId]) {
      if (!hrByCompany[companyId]) hrByCompany[companyId] = [];
      hrByCompany[companyId].push(email);
    }
  });

  const timezone = Session.getScriptTimeZone();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const properties = PropertiesService.getScriptProperties();
  let sentCount = 0;
  const enabledAlertDays = getEnabledExpiryAlertDays_();

  for (let i = 1; i < familyDocs.length; i++) {
    const row = familyDocs[i];

    const documentId = String(row[0] || '').trim();
    const familyMemberId = String(row[1] || '').trim();
    const type = String(row[2] || '').trim();
    const name = String(row[4] || '').trim();
    const expiryValue = row[6];

    if (!documentId || !familyMemberId || !(expiryValue instanceof Date)) continue;

    const member = memberById[familyMemberId];
    if (!member || !member.active || !member.employeeId) continue;

    const employee = employees[member.employeeId];
    if (!employee || !employee.active || !employee.companyId) continue;

    const companyId = employee.companyId;
    if (!activeCompanies[companyId]) continue;

    const recipients = Array.from(new Set(
      (hrByCompany[companyId] || []).concat(superAdmins)
    ));
    if (!recipients.length) continue;

    const expiryDate = new Date(expiryValue);
    if (isNaN(expiryDate.getTime())) continue;
    expiryDate.setHours(0, 0, 0, 0);

    const daysRemaining = Math.round(
      (expiryDate.getTime() - today.getTime()) / 86400000
    );

    if (!enabledAlertDays.includes(daysRemaining)) continue;

    const expiryKey = Utilities.formatDate(expiryDate, timezone, 'yyyy-MM-dd');
    const propertyKey =
      'FAMILY_EXPIRY_ALERT_' + documentId + '_' + expiryKey + '_' + daysRemaining;

    if (properties.getProperty(propertyKey)) continue;

    const expiryText = Utilities.formatDate(expiryDate, timezone, 'dd MMM yyyy');
    const subject = daysRemaining === 0
      ? 'FAMILY DOCUMENT EXPIRES TODAY: ' + name
      : 'Family document expiry in ' + daysRemaining + ' days: ' + name;

    const message =
      'Employee Document Portal\n\n' +
      'A family-member document requires attention.\n\n' +
      'Company: ' + activeCompanies[companyId] + '\n' +
      'Employee: ' + employee.name + ' (' + member.employeeId + ')\n' +
      'Family member: ' + member.name + '\n' +
      'Document: ' + name + '\n' +
      'Document type: ' + type + '\n' +
      'Expiry date: ' + expiryText + '\n' +
      'Days remaining: ' + daysRemaining + '\n\n' +
      'Please review this document and arrange renewal if needed.';

    // Send to authorized recipients and mark sent only after all sends succeed.
    recipients.forEach(function(email) {
      MailApp.sendEmail(email, subject, message);
    });

    properties.setProperty(propertyKey, new Date().toISOString());
    sentCount++;

    console.log(
      'Family expiry alert sent for ' + documentId +
      ' (' + companyId + '), days remaining: ' + daysRemaining
    );
  }

  console.log('Family expiry alert check complete. Alerts sent: ' + sentCount);
}

function checkCompanyDocumentExpiryAlerts() {
  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const documentSheet = ss.getSheetByName('CompanyDocuments');
  const rolesSheet = ss.getSheetByName('Roles');
  const companiesSheet = ss.getSheetByName('Companies');

  if (!documentSheet || !rolesSheet || !companiesSheet) {
    throw new Error('A required company-document alert sheet was not found.');
  }

  const documentRows = documentSheet.getDataRange().getValues();
  const roleRows = rolesSheet.getDataRange().getValues();
  const companyRows = companiesSheet.getDataRange().getValues();

  // Include only active companies.
  const activeCompanies = {};

  companyRows.slice(1).forEach(function(row) {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const status = String(row[2] || '').trim().toUpperCase();

    if (id && name && status === 'ACTIVE') {
      activeCompanies[id] = name;
    }
  });

  // Build the authorized recipients.
  const hrByCompany = {};
  const superAdmins = [];

  roleRows.slice(1).forEach(function(row) {
    const email = String(row[0] || '').trim().toLowerCase();
    const role = String(row[1] || '').trim().toUpperCase();
    const active =
      row[2] === true ||
      String(row[2] || '').trim().toUpperCase() === 'TRUE';
    const companyId = String(row[3] || '').trim();

    if (
      !email ||
      !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      !active
    ) {
      return;
    }

    if (role === 'SUPER_ADMIN') {
      superAdmins.push(email);
    } else if (role === 'HR' && activeCompanies[companyId]) {
      if (!hrByCompany[companyId]) {
        hrByCompany[companyId] = [];
      }

      hrByCompany[companyId].push(email);
    }
  });

  const timezone = Session.getScriptTimeZone();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const properties = PropertiesService.getScriptProperties();
  let sentCount = 0;
  const enabledAlertDays = getEnabledExpiryAlertDays_();

  for (let i = 1; i < documentRows.length; i++) {
    const row = documentRows[i];

    // CompanyDocuments columns:
    // A ID, B CompanyID, C Category, D Name, E Number,
    // F IssueDate, G ExpiryDate, H DriveFileID,
    // I CreatedAt, J UpdatedAt, K CreatedBy, L UpdatedBy, M Active.
    const documentId = String(row[0] || '').trim();
    const companyId = String(row[1] || '').trim();
    const category = String(row[2] || '').trim();
    const name = String(row[3] || '').trim();
    const documentNumber = String(row[4] || '').trim();
    const expiryValue = row[6];
    const isActive =
      row[12] === true ||
      String(row[12] || '').trim().toUpperCase() === 'TRUE';

    if (
      !documentId ||
      !companyId ||
      !name ||
      !isActive ||
      !activeCompanies[companyId] ||
      !expiryValue
    ) {
      continue;
    }

    // Accept actual spreadsheet dates and valid date strings.
    let expiryDate;

    if (expiryValue instanceof Date) {
      expiryDate = new Date(expiryValue);
    } else {
      const text = String(expiryValue).trim();
      const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);

      if (match) {
        expiryDate = new Date(
          Number(match[1]),
          Number(match[2]) - 1,
          Number(match[3])
        );
      } else {
        expiryDate = new Date(expiryValue);
      }
    }

    if (isNaN(expiryDate.getTime())) continue;

    expiryDate.setHours(0, 0, 0, 0);

    const daysRemaining = Math.round(
      (expiryDate.getTime() - today.getTime()) / 86400000
    );

   if (!enabledAlertDays.includes(daysRemaining)) continue;

    const recipients = Array.from(new Set(
      (hrByCompany[companyId] || []).concat(superAdmins)
    ));

    if (!recipients.length) continue;

    const expiryKey = Utilities.formatDate(
      expiryDate,
      timezone,
      'yyyy-MM-dd'
    );

    // Separate key namespace; existing employee/family keys remain untouched.
    const propertyKey =
      'COMPANY_EXPIRY_ALERT_' +
      documentId + '_' + expiryKey + '_' + daysRemaining;

    if (properties.getProperty(propertyKey)) continue;

    const expiryText = Utilities.formatDate(
      expiryDate,
      timezone,
      'dd MMM yyyy'
    );

    const subject = daysRemaining === 0
      ? 'COMPANY DOCUMENT EXPIRES TODAY: ' + name
      : 'Company document expiry in ' + daysRemaining + ' days: ' + name;

    const message =
      'Employee Document Portal\n\n' +
      'A company document requires your attention.\n\n' +
      'Company: ' + activeCompanies[companyId] + '\n' +
      'Category: ' + category + '\n' +
      'Document: ' + name + '\n' +
      'Document number: ' + (documentNumber || 'Not provided') + '\n' +
      'Expiry date: ' + expiryText + '\n' +
      'Days remaining: ' + daysRemaining + '\n\n' +
      'Please review this document and arrange renewal if needed.';

    // Record the milestone only after every email send succeeds.
    recipients.forEach(function(email) {
      MailApp.sendEmail(email, subject, message);
    });

    properties.setProperty(propertyKey, new Date().toISOString());
    sentCount++;

    console.log(
      'Company expiry alert sent for ' + documentId +
      ' (' + companyId + '), days remaining: ' + daysRemaining
    );
  }

  console.log(
    'Company document expiry alert check complete. Alerts sent: ' +
    sentCount
  );
}



function getEnabledExpiryAlertDays_() {
  const defaults = [90, 60, 30, 15, 7, 0];

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('ExpiryAlertSettings');

  if (!sheet || sheet.getLastRow() < 2) {
    return defaults;
  }

  const rows = sheet.getDataRange().getValues();
  const settings = {};

  rows.slice(1).forEach(function(row) {
    const key = String(row[0] || '').trim();
    const value = row[1];

    if (key) {
      settings[key] =
        value === true ||
        String(value).trim().toUpperCase() === 'TRUE';
    }
  });

  return defaults.filter(function(day) {
    const key = 'ALERT_DAY_' + day;
    return settings[key] === undefined ? true : settings[key];
  });
}

function testEnabledExpiryAlertDays() {
  const days = getEnabledExpiryAlertDays_();
  console.log('Enabled expiry alert days: ' + JSON.stringify(days));
}


function getExpiryAlertSettings() {
  const user = getCurrentUser();

  if (!user || !isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can view expiry settings.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('ExpiryAlertSettings');

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('ExpiryAlertSettings sheet is missing or empty.');
  }

  const values = sheet.getDataRange().getValues();
  const settings = {};

  values.slice(1).forEach(function(row) {
    const key = String(row[0] || '').trim();
    if (!key) return;

    settings[key] =
      row[1] === true ||
      String(row[1]).trim().toUpperCase() === 'TRUE';
  });

  return {
    day90: settings.ALERT_DAY_90 !== false,
    day60: settings.ALERT_DAY_60 !== false,
    day30: settings.ALERT_DAY_30 !== false,
    day15: settings.ALERT_DAY_15 !== false,
    day7: settings.ALERT_DAY_7 !== false,
    day0: settings.ALERT_DAY_0 !== false
  };
}

function saveExpiryAlertSettings(settings) {
  const user = getCurrentUser();

  if (!user || !isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can change expiry settings.');
  }

  if (!settings || typeof settings !== 'object') {
    throw new Error('Invalid expiry settings.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('ExpiryAlertSettings');

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('ExpiryAlertSettings sheet is missing or empty.');
  }

  const allowedDays = [90, 60, 30, 15, 7, 0];
  const values = sheet.getDataRange().getValues();
  const rowByKey = {};

  values.slice(1).forEach(function(row, index) {
    rowByKey[String(row[0] || '').trim()] = index + 2;
  });

  allowedDays.forEach(function(day) {
    const key = 'ALERT_DAY_' + day;

    if (!rowByKey[key]) {
      throw new Error('Missing setting: ' + key);
    }

    const field = 'day' + day;
    const enabled = settings[field];

    if (typeof enabled !== 'boolean') {
      throw new Error('Invalid value for ' + key);
    }

    sheet.getRange(rowByKey[key], 2, 1, 3).setValues([[
      enabled,
      new Date(),
      user.email
    ]]);
  });

  const audit = ss.getSheetByName('AuditLog');
  if (audit) {
    audit.appendRow([
      new Date(),
      user.email,
      'EXPIRY_ALERT_SETTINGS_UPDATED',
      '',
      'Reminder intervals updated.'
    ]);
  }

  return {
    success: true,
    message: 'Expiry reminder settings saved successfully.'
  };
}





