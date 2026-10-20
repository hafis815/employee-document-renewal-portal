

const CONFIG = {

  HR_EMAIL: 'ittrainee@muscatoverseas.com',

  SPREADSHEET_NAME: 'Employee Document Portal Database',

  DRIVE_FOLDER_NAME: 'Employee Document Portal - Private Files'

};



function setupPortal() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SPREADSHEET_ID')) {
    throw new Error('Portal already initialized. Check Script Properties before resetting.');
  }

  const ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);

  const employees = ss.getSheets()[0];
  employees.setName('Employees');
  employees.appendRow([
    'EmployeeID', 'Name', 'GoogleSub', 'GoogleEmail',
    'Active', 'CreatedAt', 'CompanyID'
  ]);

  const documents = ss.insertSheet('Documents');
  documents.appendRow([
    'DocumentID', 'EmployeeID', 'DocumentType', 'DocumentName',
    'IssueDate', 'ExpiryDate', 'DriveFileID', 'CreatedAt', 'UpdatedAt'
  ]);

  const roles = ss.insertSheet('Roles');
  roles.appendRow(['GoogleEmail', 'Role', 'Active', 'CompanyID']);
  // On a fresh installation, the configured initial account is the Super Admin.
  roles.appendRow([CONFIG.HR_EMAIL, 'SUPER_ADMIN', true, '']);

  const companies = ss.insertSheet('Companies');
  companies.appendRow(['CompanyID', 'CompanyName', 'Status', 'CreatedAt']);

  const familyMembers = ss.insertSheet('FamilyMembers');
  familyMembers.appendRow([
    'FamilyMemberID', 'EmployeeID', 'Name', 'Relationship', 'CreatedAt', 'Active'
  ]);

  const familyDocuments = ss.insertSheet('FamilyDocuments');
  familyDocuments.appendRow([
    'FamilyDocumentID', 'FamilyMemberID', 'DocumentType', 'DocumentNumber',
    'DocumentName', 'IssueDate', 'ExpiryDate', 'DriveFileID', 'CreatedAt', 'UpdatedAt'
  ]);

  const audit = ss.insertSheet('AuditLog');
  audit.appendRow([
  'Timestamp',
  'ActorEmail',
  'Action',
  'RecordID',
  'Details',
  'CompanyID'
]);

  const folder = DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
  folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

  props.setProperties({
    SPREADSHEET_ID: ss.getId(),
    DRIVE_FOLDER_ID: folder.getId(),
    HR_EMAIL: CONFIG.HR_EMAIL
  });

  Logger.log('Database URL: ' + ss.getUrl());
  Logger.log('Private Drive folder ID: ' + folder.getId());
}


function doGet() {

  return HtmlService.createHtmlOutputFromFile('Index')

    .setTitle('Employee Document Portal');

}





function getPortalConfig_() {

  const props = PropertiesService.getScriptProperties();



  const spreadsheetId = props.getProperty('SPREADSHEET_ID');

  const folderId = props.getProperty('DRIVE_FOLDER_ID');



  if (!spreadsheetId || !folderId) {

    throw new Error('Portal setup is incomplete.');

  }



  return {

    spreadsheetId: spreadsheetId,

    folderId: folderId,

    hrEmail: props.getProperty('HR_EMAIL')

  };

}

function isActiveCompany_(ss, companyId) {

  const id = String(companyId || '').trim();



  if (!id) return false;



  const sheet = ss.getSheetByName('Companies');



  if (!sheet) {

    throw new Error('Companies sheet was not found.');

  }



  const rows = sheet.getDataRange().getValues();



  return rows.slice(1).some(function(row) {

    return (

      String(row[0] || '').trim() === id &&

      String(row[1] || '').trim() !== '' &&

      String(row[2] || '').trim().toUpperCase() === 'ACTIVE'

    );

  });

}

function isSuperAdmin_(user) {

  return Boolean(user && user.role === 'SUPER_ADMIN');

}

function hasHrAccess_(user) {

  return Boolean(

    user &&

    (user.role === 'HR' || user.role === 'SUPER_ADMIN')

  );

}

function getActiveCompaniesForAdmin() {
  const user = getCurrentUser();
  if (!isSuperAdmin_(user)) throw new Error('Super Admin access is required.');

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('Companies');
  if (!sheet) throw new Error('Companies sheet was not found.');

  return sheet.getDataRange().getValues().slice(1).map(function(row) {
    return {
      companyId: String(row[0] || '').trim(),
      companyName: String(row[1] || '').trim(),
      status: String(row[2] || '').trim().toUpperCase()
    };
  }).filter(function(company) {
    return company.companyId && company.companyName && company.status === 'ACTIVE';
  }).sort(function(a, b) {
    return a.companyName.localeCompare(b.companyName);
  });
}

function assignCompanyHrAccess(email, companyId) {
  const user = getCurrentUser();
  if (!isSuperAdmin_(user)) throw new Error('Only the Super Admin can assign company HR access.');

  email = String(email || '').trim().toLowerCase();
  companyId = String(companyId || '').trim();

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Enter a valid HR email address.');
  }
  if (email === user.email) {
    throw new Error('Your Super Admin account cannot be assigned as a company HR account.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  if (!isActiveCompany_(ss, companyId)) {
    throw new Error('Select an active company.');
  }

  const sheet = ss.getSheetByName('Roles');
  if (!sheet) throw new Error('Roles sheet was not found.');
  const rows = sheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < rows.length; i++) {
    const existingEmail = String(rows[i][0] || '').trim().toLowerCase();
    if (existingEmail !== email) continue;
    const existingRole = String(rows[i][1] || '').trim().toUpperCase();
    if (existingRole === 'SUPER_ADMIN') {
      throw new Error('That email is already a Super Admin account.');
    }
    targetRow = i + 1;
    break;
  }

  const values = [email, 'HR', true, companyId];
  if (targetRow === -1) {
    sheet.appendRow(values);
  } else {
    sheet.getRange(targetRow, 1, 1, 4).setValues([values]);
  }

  

try {
  writeAuditLog_(
    ss,
    user.email,
    'COMPANY_HR_ASSIGNED',
    'HR email: ' + email,
    'HR email: ' + email + '; CompanyID: ' + companyId,
    companyId
  );
} catch (auditError) {
  console.error('Company HR assignment audit logging failed.');
}



  return { success: true, email: email, companyId: companyId,
    message: 'HR access assigned successfully.' };
}


function getCurrentUser() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) {
    throw new Error('Unable to identify your company account. Please sign in with your authorized Google account.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const rolesSheet = ss.getSheetByName('Roles');
  if (!rolesSheet) throw new Error('Roles sheet was not found.');

  const roles = rolesSheet.getDataRange().getValues();
  for (let i = 1; i < roles.length; i++) {
    const roleEmail = String(roles[i][0] || '').trim().toLowerCase();
    const role = String(roles[i][1] || '').trim().toUpperCase();
    const active = roles[i][2] === true || String(roles[i][2] || '').trim().toUpperCase() === 'TRUE';

    if (roleEmail !== email || !active) continue;
    if (role !== 'HR' && role !== 'SUPER_ADMIN') continue;

    const companyId = String(roles[i][3] || '').trim();
    if (role !== 'SUPER_ADMIN' && !isActiveCompany_(ss, companyId)) {
      throw new Error('Your HR account is not assigned to an active company.');
    }

    return {
      email: email,
      name: email,
      role: role,
      companyId: companyId
    };
  }

  const employeesSheet = ss.getSheetByName('Employees');
  if (!employeesSheet) throw new Error('Employees sheet was not found.');
  const employees = employeesSheet.getDataRange().getValues();

  for (let i = 1; i < employees.length; i++) {
    const employeeEmail = String(employees[i][3] || '').trim().toLowerCase();
    const active = employees[i][4] === true || String(employees[i][4] || '').trim().toUpperCase() === 'TRUE';
    if (employeeEmail !== email || !active) continue;

    const companyId = String(employees[i][6] || '').trim();
    if (!isActiveCompany_(ss, companyId)) {
      throw new Error('Your employee account is not assigned to an active company.');
    }

    return {
      email: email,
      name: String(employees[i][1] || email),
      role: 'Employee',
      companyId: companyId
    };
  }

  throw new Error('Your company account is not registered or active. Please contact HR.');
}


function getEmployeeByGoogleSub_(googleSub) {

  if (!googleSub) {

    throw new Error('Verified Google identity is required.');

  }



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const sheet = ss.getSheetByName('Employees');

  const rows = sheet.getDataRange().getValues();



  for (let i = 1; i < rows.length; i++) {

    const employeeId = rows[i][0];

    const storedSub = rows[i][2];

    const active = rows[i][4];



    if (

      String(storedSub) === String(googleSub) &&

      active === true

    ) {

      return {

        employeeId: employeeId,

        name: rows[i][1],

        email: rows[i][3]

      };

    }

  }



  return null;

}



function getEmployeeDocuments_(employeeId) {

  if (!employeeId) {

    throw new Error('Employee ID is required.');

  }



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

 const sheet = ss.getSheetByName('Documents');

if (!sheet) {
  throw new Error('Documents sheet was not found.');
}

const rows = sheet.getDataRange().getValues();

  const results = [];
  

  for (let i = 1; i < rows.length; i++) {

    if (String(rows[i][1]) !== String(employeeId)) {

      continue;

    }



    results.push({

      documentId: rows[i][0],

      type: rows[i][2],

      name: rows[i][3],

      expiryDate: rows[i][5] instanceof Date
  ? Utilities.formatDate(
      rows[i][5],
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    )
  : String(rows[i][5] || '')

    });

  }



  return results;

}




function getAuthorizedDocumentFile(documentId, sessionToken) {
  let user;
  let sessionEmployeeId = '';

  // Password-login employees must have a valid server-side session.
  if (sessionToken) {
    user = getEmployeeSessionUser_(sessionToken);

if (!user || user.role !== 'Employee' || !user.employeeId) {
  throw new Error('You are not authorized to access this document.');
}

if (user.mustChangePassword) {
  throw new Error(
    'Change your temporary password before viewing documents.'
  );
}

sessionEmployeeId = String(user.employeeId).trim();

    if (user.role !== 'Employee' || !sessionEmployeeId) {
      throw new Error('You are not authorized to access this document.');
    }
  } else {
    // Preserve existing Google-account authorization for HR and employees.
    user = getCurrentUser();
  }

  if (!user || !user.role) {
    throw new Error('You are not authorized to access this document.');
  }

  documentId = String(documentId || '').trim();

  if (!documentId) {
    throw new Error('Document ID is required.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const documentsSheet = ss.getSheetByName('Documents');

  if (!documentsSheet) {
    throw new Error('Documents sheet was not found.');
  }

  const rows = documentsSheet.getDataRange().getValues();
  let documentRecord = null;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === documentId) {
      documentRecord = rows[i];
      break;
    }
  }

  if (!documentRecord) {
    throw new Error('Document not found.');
  }

  const documentEmployeeId = String(documentRecord[1] || '').trim();
  const driveFileId = String(documentRecord[6] || '').trim();

  if (!driveFileId) {
    throw new Error('The document file is unavailable.');
  }

  const employeesSheet = ss.getSheetByName('Employees');

  if (!employeesSheet) {
    throw new Error('Employees sheet was not found.');
  }

  const employeeRows = employeesSheet.getDataRange().getValues();
  let owner = null;

  for (let i = 1; i < employeeRows.length; i++) {
    if (
      String(employeeRows[i][0] || '').trim() === documentEmployeeId
    ) {
      owner = employeeRows[i];
      break;
    }
  }

  if (!owner) {
    throw new Error('The document owner was not found.');
  }

  const employeeActive = owner[4] === true ||
    String(owner[4] || '').trim().toUpperCase() === 'TRUE';

  const documentCompanyId = String(owner[6] || '').trim();

  if (
    !employeeActive ||
    !documentCompanyId ||
    !isActiveCompany_(ss, documentCompanyId)
  ) {
    throw new Error('You are not authorized to access this document.');
  }

  // Enforce company-level authorization.
  if (
    !isSuperAdmin_(user) &&
    documentCompanyId !== String(user.companyId || '').trim()
  ) {
    throw new Error('You are not authorized to access this document.');
  }

  // Employees can access only their own documents.
  if (user.role === 'Employee') {
    let authorizedEmployeeId = sessionEmployeeId;

    // Google-login employee: identify the employee by registered email.
    if (!sessionToken) {
      const userEmail = String(user.email || '').trim().toLowerCase();

      for (let i = 1; i < employeeRows.length; i++) {
        const email = String(employeeRows[i][3] || '')
          .trim()
          .toLowerCase();

        const active = employeeRows[i][4] === true ||
          String(employeeRows[i][4] || '').trim().toUpperCase() === 'TRUE';

        if (email === userEmail && active) {
          authorizedEmployeeId = String(employeeRows[i][0] || '').trim();
          break;
        }
      }
    }

    if (
      !authorizedEmployeeId ||
      authorizedEmployeeId !== documentEmployeeId
    ) {
      throw new Error('You are not authorized to access this document.');
    }
  } else if (!hasHrAccess_(user)) {
    throw new Error('You are not authorized to access this document.');
  }

  // Return the file only after all authorization checks pass.
  const file = DriveApp.getFileById(driveFileId);

  return {
    name: file.getName(),
    mimeType: file.getMimeType(),
    base64Data: Utilities.base64Encode(file.getBlob().getBytes())
  };
}




function getAuthorizedFamilyDocumentFile(familyDocumentId) {

  const user = getCurrentUser();



  if (!hasHrAccess_(user)) {

    throw new Error(

      'Only authorized HR users can access family documents.'

    );

  }



  familyDocumentId = String(familyDocumentId || '').trim();



  if (!familyDocumentId) {

    throw new Error('Family document ID is required.');

  }



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const documentsSheet = ss.getSheetByName('FamilyDocuments');

  const membersSheet = ss.getSheetByName('FamilyMembers');



  if (!documentsSheet || !membersSheet) {

    throw new Error('Family document sheets were not found.');

  }



  const documentRows = documentsSheet.getDataRange().getValues();

  const memberRows = membersSheet.getDataRange().getValues();



  let documentRecord = null;



  for (let i = 1; i < documentRows.length; i++) {

    if (String(documentRows[i][0]).trim() === familyDocumentId) {

      documentRecord = documentRows[i];

      break;

    }

  }



  if (!documentRecord) {

    throw new Error('Family document not found.');

  }



  const familyMemberId = String(documentRecord[1]).trim();



  let activeMemberFound = false;



  for (let i = 1; i < memberRows.length; i++) {

    const active =

      memberRows[i][5] === true ||

      String(memberRows[i][5]).trim().toUpperCase() === 'TRUE';



    if (

      String(memberRows[i][0]).trim() === familyMemberId &&

      active

    ) {

      activeMemberFound = true;

      break;

    }

  }



  if (!activeMemberFound) {

    throw new Error('The family member is not active or was not found.');

  }

  // Find the employee linked to this family member.

let familyEmployeeId = '';



for (let i = 1; i < memberRows.length; i++) {

  if (

    String(memberRows[i][0] || '').trim() === familyMemberId

  ) {

    familyEmployeeId = String(memberRows[i][1] || '').trim();

    break;

  }

}



if (!familyEmployeeId) {

  throw new Error('The family member has no employee record.');

}



// Find the employee's company.

const employeesSheet = ss.getSheetByName('Employees');



if (!employeesSheet) {

  throw new Error('Employees sheet was not found.');

}



const employeeRows = employeesSheet.getDataRange().getValues();

let familyDocumentCompanyId = '';



for (let i = 1; i < employeeRows.length; i++) {

  if (

    String(employeeRows[i][0] || '').trim() === familyEmployeeId

  ) {

    familyDocumentCompanyId =

      String(employeeRows[i][6] || '').trim();

    break;

  }

}



// Only HR belonging to the same company may access the file.

if (

  !familyDocumentCompanyId ||

  (

    !isSuperAdmin_(user) &&

    familyDocumentCompanyId !== String(user.companyId || '').trim()

  )

) {

  throw new Error(

    'You are not authorized to access this family document.'

  );

}



  const driveFileId = String(documentRecord[7] || '').trim();



  if (!driveFileId) {

    throw new Error('The family document file is unavailable.');

  }



  const file = DriveApp.getFileById(driveFileId);



  return {

    name: file.getName(),

    mimeType: file.getMimeType(),

    base64Data: Utilities.base64Encode(file.getBlob().getBytes())

  };

}



const GOOGLE_CLIENT_ID = '471532447554-9hnet3213phfkgu1c0r9t9m657877s23.apps.googleusercontent.com';



function authenticateWithGoogle(idToken) {
  if (!idToken || typeof idToken !== 'string') throw new Error('Please sign in with Google.');

  const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('Google sign-in verification failed.');

  const token = JSON.parse(response.getContentText());
  if (token.aud !== GOOGLE_CLIENT_ID) throw new Error('Invalid Google sign-in audience.');
  if (token.iss !== 'https://accounts.google.com' && token.iss !== 'accounts.google.com') {
    throw new Error('Invalid Google token issuer.');
  }
  if (Number(token.exp) <= Math.floor(Date.now() / 1000)) {
    throw new Error('Google sign-in has expired. Please try again.');
  }
  if (token.email_verified !== 'true' || !token.email) throw new Error('A verified Google email is required.');

  const email = String(token.email).trim().toLowerCase();
  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const rolesSheet = ss.getSheetByName('Roles');
  const employeesSheet = ss.getSheetByName('Employees');
  if (!rolesSheet || !employeesSheet) throw new Error('Roles or Employees sheet was not found.');

  let role = null;
  let companyId = '';
  const roles = rolesSheet.getDataRange().getValues();
  for (let i = 1; i < roles.length; i++) {
    const active = roles[i][2] === true || String(roles[i][2] || '').trim().toUpperCase() === 'TRUE';
    if (String(roles[i][0] || '').trim().toLowerCase() !== email || !active) continue;
    const matchedRole = String(roles[i][1] || '').trim().toUpperCase();
    if (matchedRole !== 'HR' && matchedRole !== 'SUPER_ADMIN') continue;
    role = matchedRole;
    companyId = String(roles[i][3] || '').trim();
    break;
  }

  if (!role) {
    const employees = employeesSheet.getDataRange().getValues();
    for (let i = 1; i < employees.length; i++) {
      const active = employees[i][4] === true || String(employees[i][4] || '').trim().toUpperCase() === 'TRUE';
      if (String(employees[i][3] || '').trim().toLowerCase() !== email || !active) continue;
      role = 'Employee';
      companyId = String(employees[i][6] || '').trim();
      break;
    }
  }

  if (!role) throw new Error('Your Google account has not been registered by HR.');
  if (role !== 'SUPER_ADMIN' && !isActiveCompany_(ss, companyId)) {
    throw new Error('Your account is not assigned to an active company.');
  }

  return { email: email, name: token.name || '', role: role, companyId: companyId };
}



function getMyDocuments(sessionToken) {
  let user;
  let employeeId = null;

  // Password login: validate the server-side session.
  if (sessionToken) {
    user = getEmployeeSessionUser_(sessionToken);

if (!user || user.role !== 'Employee' || !user.employeeId) {
  throw new Error('Employee access is required.');
}

if (user.mustChangePassword) {
  throw new Error(
    'Change your temporary password before accessing documents.'
  );
}

    employeeId = String(user.employeeId).trim();

  } else {
    // Preserve the existing Google-account login.
    user = getCurrentUser();

    if (user.role !== 'Employee') {
      throw new Error(
        'This page is available to registered employees only.'
      );
    }

    const config = getPortalConfig_();
    const ss = SpreadsheetApp.openById(config.spreadsheetId);
    const employeeSheet = ss.getSheetByName('Employees');

    if (!employeeSheet) {
      throw new Error('Employees sheet was not found.');
    }

    const employeeRows = employeeSheet.getDataRange().getValues();

    for (let i = 1; i < employeeRows.length; i++) {
      const email = String(employeeRows[i][3] || '')
        .trim()
        .toLowerCase();

      const active = employeeRows[i][4] === true ||
        String(employeeRows[i][4] || '').trim().toUpperCase() === 'TRUE';

      if (email === String(user.email || '').trim().toLowerCase() && active) {
        employeeId = String(employeeRows[i][0] || '').trim();
        break;
      }
    }
  }

  if (!employeeId) {
    throw new Error(
      'No active employee record was found for your account.'
    );
  }

  // Return only this employee's document metadata.
  const documents = getEmployeeDocuments_(employeeId);
Logger.log('Returning documents: ' + JSON.stringify(documents));
return documents;
}






function getHrDashboardData() {

  const user = getCurrentUser();



  if (!hasHrAccess_(user)) {

  throw new Error('HR access is required.');

}



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);



  const employeeSheet = ss.getSheetByName('Employees');

  const documentSheet = ss.getSheetByName('Documents');

  const familySheet = ss.getSheetByName('FamilyMembers');

  const familyDocumentSheet = ss.getSheetByName('FamilyDocuments');



  if (!employeeSheet || !documentSheet ||

      !familySheet || !familyDocumentSheet) {

    throw new Error('Required portal sheets were not found.');

  }



  const employeeRows = employeeSheet.getDataRange().getValues();

  const documentRows = documentSheet.getDataRange().getValues();

  const familyRows = familySheet.getDataRange().getValues();

  const familyDocumentRows =

    familyDocumentSheet.getDataRange().getValues();



  const employees = [];

  const employeeIds = {};

  const employeeNames = {};



  for (let i = 1; i < employeeRows.length; i++) {

    const row = employeeRows[i];

    const id = String(row[0] || '').trim();



    if (!id) continue;

    const isActive = row[4] === true ||
      String(row[4] || '').trim().toUpperCase() === 'TRUE';
    if (!isActive) continue;

    const employeeCompanyId = String(row[6] || '').trim();



if (

  !isSuperAdmin_(user) &&

  employeeCompanyId !== String(user.companyId || '').trim()

) continue;



    employeeIds[id] = true;

    employeeNames[id] = String(row[1] || '');



    employees.push({

      employeeId: id,

      name: String(row[1] || ''),

      email: String(row[3] || ''),

      active: row[4] === true

    });

  }



  const today = new Date();

  today.setHours(0, 0, 0, 0);



  function getExpiryStatus(value) {

    if (!value) return 'No expiry date';



    let expiry;



    if (value instanceof Date && !isNaN(value.getTime())) {

      expiry = new Date(value);

    } else {

      const text = String(value).trim();

      const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);



      if (match) {

        expiry = new Date(

          Number(match[1]),

          Number(match[2]) - 1,

          Number(match[3])

        );

      } else {

        expiry = new Date(value);

      }

    }



    if (isNaN(expiry.getTime())) return 'Unknown';



    expiry.setHours(0, 0, 0, 0);



    if (expiry < today) return 'Expired';



    const days = Math.round(

      (expiry.getTime() - today.getTime()) / 86400000

    );



    if (days <= 15) return 'Expires in 15 days';

    if (days <= 30) return 'Expires in 30 days';

    if (days <= 60) return 'Expires in 60 days';

    if (days <= 90) return 'Expires in 90 days';



    return 'Valid';

  }



  const documents = [];



  for (let i = 1; i < documentRows.length; i++) {

    const row = documentRows[i];

    const employeeId = String(row[1] || '').trim();



    if (!employeeIds[employeeId]) continue;



    const expiryDate = formatSheetDate_(row[5]);



    documents.push({

      documentId: String(row[0] || ''),

      employeeId: employeeId,

      employeeName: employeeNames[employeeId] || employeeId,

      type: String(row[2] || ''),

      name: String(row[3] || ''),

      issueDate: formatSheetDate_(row[4]),

      expiryDate: expiryDate,

      status: getExpiryStatus(row[5])

    });

  }



  const familyMembers = [];

  const familyMemberInfo = {};



  for (let i = 1; i < familyRows.length; i++) {

    const row = familyRows[i];

    const id = String(row[0] || '').trim();

    const employeeId = String(row[1] || '').trim();



    if (!id || !employeeIds[employeeId]) continue;



    const active =

      row[5] === true ||

      String(row[5]).trim().toUpperCase() === 'TRUE';



    const member = {

      familyMemberId: id,

      employeeId: employeeId,

      employeeName: employeeNames[employeeId] || employeeId,

      name: String(row[2] || ''),

      relationship: String(row[3] || ''),

      active: active

    };



    familyMemberInfo[id] = member;

    familyMembers.push(member);

  }



  const familyDocuments = [];



  for (let i = 1; i < familyDocumentRows.length; i++) {

    const row = familyDocumentRows[i];

    const memberId = String(row[1] || '').trim();

    const member = familyMemberInfo[memberId];



    if (!member) continue;



    familyDocuments.push({

      familyDocumentId: String(row[0] || ''),

      employeeId: member.employeeId,

      employeeName: member.employeeName,

      familyMemberId: memberId,

      familyMemberName: member.name,

      relationship: member.relationship,

      type: String(row[2] || ''),

      name: String(row[4] || ''),

      issueDate: formatSheetDate_(row[5]),

      expiryDate: formatSheetDate_(row[6]),

      status: getExpiryStatus(row[6])

    });

  }



  const allDocuments = documents.concat(

    familyDocuments.map(function(doc) {

      return {

        documentId: doc.familyDocumentId,

        employeeId: doc.employeeId,

        employeeName: doc.employeeName,

        type: doc.type,

        name: doc.name,

        expiryDate: doc.expiryDate,

        status: doc.status,

        familyMemberName: doc.familyMemberName

      };

    })

  );



  const counts = {

    totalEmployees: employees.length,

    activeEmployees: employees.filter(function(e) {

      return e.active;

    }).length,

    totalFamilyMembers: familyMembers.length,

    totalEmployeeDocuments: documents.length,

    totalFamilyDocuments: familyDocuments.length,

    totalDocuments: allDocuments.length,

    expired: 0,

    expiringIn15Days: 0,

    expiringIn30Days: 0,

    expiringIn60Days: 0,

    expiringIn90Days: 0,

    noExpiryDate: 0

  };



  allDocuments.forEach(function(doc) {

    if (doc.status === 'Expired') {

      counts.expired++;

    } else if (doc.status === 'Expires in 15 days') {

      counts.expiringIn15Days++;

    } else if (doc.status === 'Expires in 30 days') {

      counts.expiringIn30Days++;

    } else if (doc.status === 'Expires in 60 days') {

      counts.expiringIn60Days++;

    } else if (doc.status === 'Expires in 90 days') {

      counts.expiringIn90Days++;

    } else if (doc.status === 'No expiry date') {

      counts.noExpiryDate++;

    }

  });



  return {

    employees: employees,

    documents: documents,

    familyMembers: familyMembers,

    familyDocuments: familyDocuments,

    counts: counts

  };

}





function addEmployee(employeeName, employeeEmail, targetCompanyId) {
  const user = getCurrentUser();
  if (!hasHrAccess_(user)) throw new Error('Only authorized HR users can add employees.');

  const companyId = isSuperAdmin_(user)
    ? String(targetCompanyId || '').trim()
    : String(user.companyId || '').trim();
  if (!companyId) {
    throw new Error(isSuperAdmin_(user)
      ? 'Please select a company for this employee.'
      : 'Your HR account is not assigned to a company.');
  }

  const name = String(employeeName || '').trim();
  const email = String(employeeEmail || '').trim().toLowerCase();
  if (!name) throw new Error('Employee name is required.');
  if (name.length > 150) throw new Error('Employee name is too long.');
  if (email.length > 254 || !/^[^\s@]+@muscatoverseas\.com$/i.test(email)) {
    throw new Error('Enter a valid Muscat Overseas company email address.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  if (!isActiveCompany_(ss, companyId)) throw new Error('The selected company is not active or does not exist.');
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees sheet was not found.');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][3] || '').trim().toLowerCase() === email) {
      throw new Error('This email address is already registered.');
    }
  }

  
  const employeeId = 'EMP-' + Utilities.getUuid().substring(0, 8).toUpperCase();
  sheet.appendRow([employeeId, name, '', email, true, new Date(), companyId]);

  try {
    writeAuditLog_(
      ss,
      user.email,
      'EMPLOYEE_ADDED',
      employeeId,
      'Employee registered: ' + name + '; Email: ' + email,
      companyId
    );
  } catch (auditError) {
    console.error('Employee registration audit logging failed.');
  }

  return {
    employeeId: employeeId,
    name: name,
    email: email,
    active: true,
    companyId: companyId
  };

}


function uploadEmployeeDocument(

  employeeId,

  documentType,

  issueDate,

  expiryDate,

  fileName,

  mimeType,

  base64Data

) {

  // Only authorized HR users may upload employee documents.

  const user = getCurrentUser();



  if (!hasHrAccess_(user)) {

  throw new Error('Only authorized HR users can upload documents.');

}



  employeeId = String(employeeId || '').trim();

  documentType = String(documentType || '').trim();

  fileName = String(fileName || '').trim();

  mimeType = String(mimeType || '').toLowerCase();



  if (!employeeId || !documentType || !fileName || !base64Data) {

    throw new Error('Please complete all required fields and select a file.');

  }



  const allowedTypes = {

    'application/pdf': '.pdf',

    'image/jpeg': '.jpg',

    'image/png': '.png'

  };



  if (!allowedTypes[mimeType]) {

    throw new Error('Only PDF, JPEG, and PNG files are allowed.');

  }

  const expectedExtension = allowedTypes[mimeType];



const lowerFileName = fileName.toLowerCase();



if (

  mimeType === 'image/jpeg'

    ? !(/\.(jpg|jpeg)$/.test(lowerFileName))

    : !lowerFileName.endsWith(expectedExtension)

) { 

  throw new Error('The file extension does not match its file type.');

}



  if (fileName.length > 200) {

    throw new Error('The file name is too long.');

  }



  // Decode and enforce a 5 MB limit on the server.

  const bytes = Utilities.base64Decode(base64Data);



  if (bytes.length === 0 || bytes.length > 5 * 1024 * 1024) {

    throw new Error('The file must be larger than 0 bytes and no more than 5 MB.');

  }



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const employeesSheet = ss.getSheetByName('Employees');

  const documentsSheet = ss.getSheetByName('Documents');



  if (!employeesSheet || !documentsSheet) {

    throw new Error('Required portal sheets were not found.');

  }



  // Confirm that the target employee exists and is active.

  const employeeRows = employeesSheet.getDataRange().getValues();

  let activeEmployeeFound = false;
let employeeCompanyId = '';



  for (let i = 1; i < employeeRows.length; i++) {
    const rowEmployeeId = String(employeeRows[i][0] || '').trim();
    const rowCompanyId = String(employeeRows[i][6] || '').trim();
    const rowActive = employeeRows[i][4] === true ||
      String(employeeRows[i][4] || '').trim().toUpperCase() === 'TRUE';
    const companyIsActive = isActiveCompany_(ss, rowCompanyId);
    const sameCompany = isSuperAdmin_(user) ||
      rowCompanyId === String(user.companyId || '').trim();

    if (rowEmployeeId === employeeId && rowActive && companyIsActive && sameCompany) {
  activeEmployeeFound = true;
  employeeCompanyId = rowCompanyId;
  break;
}
  }



  if (!activeEmployeeFound) {

    throw new Error('The selected employee does not exist or is inactive.');

  }



  // Validate optional dates before creating the file.

  function parseOptionalDate_(value, fieldName) {

    if (!value) return '';



    const text = String(value);

    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {

      throw new Error('Invalid ' + fieldName + '.');

    }



    const parts = text.split('-').map(Number);

    const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));



    if (

      date.getUTCFullYear() !== parts[0] ||

      date.getUTCMonth() !== parts[1] - 1 ||

      date.getUTCDate() !== parts[2]

    ) {

      throw new Error('Invalid ' + fieldName + '.');

    }



    return date;

  }



  const parsedIssueDate = parseOptionalDate_(issueDate, 'issue date');

  const parsedExpiryDate = parseOptionalDate_(expiryDate, 'expiry date');



  if (

    parsedIssueDate &&

    parsedExpiryDate &&

    parsedExpiryDate < parsedIssueDate

  ) {

    throw new Error('Expiry date cannot be earlier than issue date.');

  }



  // Save the file into the existing private folder.

  const safeName = fileName.replace(/[\\/:*?"<>|]/g, '_');

  const blob = Utilities.newBlob(bytes, mimeType, safeName);

  const folder = DriveApp.getFolderById(config.folderId);

  const file = folder.createFile(blob);



  try {

    const documentId = 'DOC-' +

      Utilities.getUuid().substring(0, 8).toUpperCase();

    const now = new Date();



    documentsSheet.appendRow([

      documentId,

      employeeId,

      documentType,

      safeName,

      parsedIssueDate,

      parsedExpiryDate,

      file.getId(),

      now,

      now

    ]);

try {
  writeAuditLog_(
    ss,
    user.email,
    'EMPLOYEE_DOCUMENT_UPLOADED',
    documentId,
    'EmployeeID: ' + employeeId +
      '; Document: ' + safeName,
    employeeCompanyId
  );
} catch (auditError) {
  console.error('Employee upload audit logging failed.');
}




    return {

      documentId: documentId,

      employeeId: employeeId,

      name: safeName,

      type: documentType,

      expiryDate: formatSheetDate_(parsedExpiryDate)

    };

  } catch (error) {

    // Avoid leaving an untracked file if the sheet write fails.

    file.setTrashed(true);

    throw new Error('The document record could not be saved.');

  }

}







function formatSheetDate_(value) {

  if (!value) return '';



  if (value instanceof Date && !isNaN(value.getTime())) {

    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      'yyyy-MM-dd'

    );

  }



  return String(value);

}



function addFamilyMember(employeeId, memberName, relationship) {
  const user = getCurrentUser();
  if (!hasHrAccess_(user)) throw new Error('HR access is required.');

  employeeId = String(employeeId || '').trim();
  memberName = String(memberName || '').trim();
  relationship = String(relationship || '').trim();
  if (!employeeId || !memberName || !relationship) {
    throw new Error('Employee, family member name, and relationship are required.');
  }
  const allowedRelationships = ['Spouse', 'Son', 'Daughter', 'Father', 'Mother', 'Other'];
  if (!allowedRelationships.includes(relationship)) throw new Error('Please select a valid relationship.');

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const employeeSheet = ss.getSheetByName('Employees');
  const familySheet = ss.getSheetByName('FamilyMembers');
  if (!employeeSheet || !familySheet) throw new Error('Employees or FamilyMembers sheet was not found.');


const employees = employeeSheet.getDataRange().getValues();
let employeeCompanyId = '';

const employeeExists = employees.slice(1).some(function(row) {
  const active = row[4] === true ||
    String(row[4] || '').trim().toUpperCase() === 'TRUE';
  const rowCompanyId = String(row[6] || '').trim();
  const companyIsActive = isActiveCompany_(ss, rowCompanyId);
  const matchingCompany = isSuperAdmin_(user) ||
    rowCompanyId === String(user.companyId || '').trim();

  const matches = String(row[0] || '').trim() === employeeId &&
    active && companyIsActive && matchingCompany;

  if (matches) {
    employeeCompanyId = rowCompanyId;
  }

  return matches;
});

if (!employeeExists) {
  throw new Error(
    'The selected employee does not exist, is inactive, or belongs to another company.'
  );
}
const memberId = 'FAM-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();

familySheet.appendRow([
  memberId,
  employeeId,
  memberName,
  relationship,
  new Date(),
  true
]);

try {
  writeAuditLog_(
    ss,
    user.email,
    'FAMILY_MEMBER_ADDED',
    memberId,
    'EmployeeID: ' + employeeId,
    employeeCompanyId
  );
} catch (auditError) {
  console.error('Family member registration audit logging failed.');
}
}


function uploadFamilyDocument(

  familyMemberId,

  documentType,

  documentNumber,

  issueDate,

  expiryDate,

  fileName,

  mimeType,

  base64Data

) {

  const user = getCurrentUser();



  if (!hasHrAccess_(user)) {

  throw new Error('Only authorized HR users can upload family documents.');

}



  familyMemberId = String(familyMemberId || '').trim();

  documentType = String(documentType || '').trim();

  documentNumber = String(documentNumber || '').trim();

  fileName = String(fileName || '').trim();

  mimeType = String(mimeType || '').toLowerCase();



  if (!familyMemberId || !documentType || !fileName || !base64Data) {

    throw new Error('Please complete the required fields and select a file.');

  }



  const allowedTypes = {

    'application/pdf': '.pdf',

    'image/jpeg': '.jpg',

    'image/png': '.png'

  };



  if (!allowedTypes[mimeType]) {

    throw new Error('Only PDF, JPEG, and PNG files are allowed.');

  }



  const expectedExtension = allowedTypes[mimeType];

  const lowerFileName = fileName.toLowerCase();



  if (

    mimeType === 'image/jpeg'

      ? !(/\.(jpg|jpeg)$/.test(lowerFileName))

      : !lowerFileName.endsWith(expectedExtension)

  ) {

    throw new Error('The file extension does not match its file type.');

  }



  if (fileName.length > 200) {

    throw new Error('The file name is too long.');

  }



  const bytes = Utilities.base64Decode(base64Data);



  if (bytes.length === 0 || bytes.length > 5 * 1024 * 1024) {

    throw new Error('The file must be larger than 0 bytes and no more than 5 MB.');

  }



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const familySheet = ss.getSheetByName('FamilyMembers');

  const documentsSheet = ss.getSheetByName('FamilyDocuments');



  if (!familySheet || !documentsSheet) {

    throw new Error('FamilyMembers or FamilyDocuments sheet was not found.');

  }



  const familyRows = familySheet.getDataRange().getValues();

  let activeMember = false;



  for (let i = 1; i < familyRows.length; i++) {

    if (

      String(familyRows[i][0] || '').trim() === familyMemberId &&

      (familyRows[i][5] === true || String(familyRows[i][5] || '').trim().toUpperCase() === 'TRUE')

    ) {

      activeMember = true;

      break;

    }

  }



  if (!activeMember) {

    throw new Error('The family member does not exist or is inactive.');

  }

  // Find the employee linked to this family member.

let familyEmployeeId = '';



for (let i = 1; i < familyRows.length; i++) {

  if (

    String(familyRows[i][0] || '').trim() === familyMemberId

  ) {

    familyEmployeeId = String(familyRows[i][1] || '').trim();

    break;

  }

}



if (!familyEmployeeId) {

  throw new Error('The family member has no employee record.');

}



// Verify the employee belongs to the signed-in HR user's company.

const employeesSheet = ss.getSheetByName('Employees');



if (!employeesSheet) {

  throw new Error('Employees sheet was not found.');

}



const employeeRows = employeesSheet.getDataRange().getValues();
let authorizedFamilyCompany = false;
let familyDocumentCompanyId = '';

for (let i = 1; i < employeeRows.length; i++) {
  const employeeActive =
    employeeRows[i][4] === true ||
    String(employeeRows[i][4] || '').trim().toUpperCase() === 'TRUE';

  const employeeCompanyId = String(employeeRows[i][6] || '').trim();
  const companyIsActive = isActiveCompany_(ss, employeeCompanyId);

  if (
    String(employeeRows[i][0] || '').trim() === familyEmployeeId &&
    employeeActive &&
    companyIsActive &&
    (isSuperAdmin_(user) ||
      employeeCompanyId === String(user.companyId || '').trim())
  ) {
    authorizedFamilyCompany = true;
    familyDocumentCompanyId = employeeCompanyId;
    break;
  }
}

if (!authorizedFamilyCompany) {
  throw new Error(
    'You can upload family documents only for your own company.'
  );
}




if (!authorizedFamilyCompany) {

  throw new Error(

    'You can upload family documents only for your own company.'

  );

}



  function parseFamilyDate_(value, fieldName) {

    if (!value) return '';



    const text = String(value);



    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {

      throw new Error('Invalid ' + fieldName + '.');

    }



    const parts = text.split('-').map(Number);

    const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));



    if (

      date.getUTCFullYear() !== parts[0] ||

      date.getUTCMonth() !== parts[1] - 1 ||

      date.getUTCDate() !== parts[2]

    ) {

      throw new Error('Invalid ' + fieldName + '.');

    }



    return date;

  }



  const parsedIssueDate = parseFamilyDate_(issueDate, 'issue date');

  const parsedExpiryDate = parseFamilyDate_(expiryDate, 'expiry date');



  if (

    parsedIssueDate &&

    parsedExpiryDate &&

    parsedExpiryDate < parsedIssueDate

  ) {

    throw new Error('Expiry date cannot be earlier than issue date.');

  }



  const safeName = fileName.replace(/[\\/:*?"<>|]/g, '_');

  const blob = Utilities.newBlob(bytes, mimeType, safeName);

  const folder = DriveApp.getFolderById(config.folderId);

  const file = folder.createFile(blob);



  try {

    const familyDocumentId = 'FDOC-' +

      Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();

    const now = new Date();



    documentsSheet.appendRow([

      familyDocumentId,

      familyMemberId,

      documentType,

      documentNumber,

      safeName,

      parsedIssueDate,

      parsedExpiryDate,

      file.getId(),

      now,

      now

    ]);


try {
  writeAuditLog_(
    ss,
    user.email,
    'FAMILY_DOCUMENT_UPLOADED',
    familyDocumentId,
    'FamilyMemberID: ' + familyMemberId +
      '; Document: ' + safeName,
    familyDocumentCompanyId
  );
} catch (auditError) {
  console.error('Family upload audit logging failed.');
}



    return {

      familyDocumentId: familyDocumentId,

      familyMemberId: familyMemberId,

      name: safeName,

      type: documentType,

      expiryDate: formatSheetDate_(parsedExpiryDate)

    };

  } catch (error) {

    file.setTrashed(true);

    throw new Error('The family document record could not be saved.');

  }

}



function getFamilyMembersForHr() {

  const user = getCurrentUser();



  if (!hasHrAccess_(user)) {

  throw new Error('Access denied. HR access is required.');

}



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);



  const familySheet = ss.getSheetByName('FamilyMembers');

  const employeeSheet = ss.getSheetByName('Employees');



  if (!familySheet || !employeeSheet) {

    throw new Error('FamilyMembers or Employees sheet was not found.');

  }



  const familyRows = familySheet.getDataRange().getValues();

  const employeeRows = employeeSheet.getDataRange().getValues();



  const employeeNames = {};

const employeeCompanyIds = {};



  for (let i = 1; i < employeeRows.length; i++) {

  const employeeId = String(employeeRows[i][0] || '').trim();

  const employeeName = String(employeeRows[i][1] || '').trim();



  if (employeeId) {

    employeeNames[employeeId] = employeeName;

    employeeCompanyIds[employeeId] =

      String(employeeRows[i][6] || '').trim();

  }

}



  const members = [];



  for (let i = 1; i < familyRows.length; i++) {

    const row = familyRows[i];



    const familyMemberId = String(row[0]).trim();

    const employeeId = String(row[1]).trim();

    const name = String(row[2]).trim();

    const relationship = String(row[3]).trim();

    const active = row[5] === true ||

      String(row[5]).trim().toUpperCase() === 'TRUE';



    if (

  familyMemberId &&

  employeeId &&

  name &&

  active &&

  (

    isSuperAdmin_(user) ||

    employeeCompanyIds[employeeId] ===

      String(user.companyId || '').trim()

  )

) {

  members.push({

        familyMemberId: familyMemberId,

        employeeId: employeeId,

        employeeName: employeeNames[employeeId] || employeeId,

        name: name,

        relationship: relationship

      });

    }

  }



  return { members: members };

}



function getHrFamilyDocumentData() {

  const user = getCurrentUser();



  if (!hasHrAccess_(user)) {

  throw new Error('Access denied. HR access is required.');

}



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);



  const familySheet = ss.getSheetByName('FamilyMembers');

  const documentSheet = ss.getSheetByName('FamilyDocuments');

  const employeeSheet = ss.getSheetByName('Employees');



  if (!familySheet || !documentSheet || !employeeSheet) {

    throw new Error('A required spreadsheet sheet was not found.');

  }



  const familyRows = familySheet.getDataRange().getValues();

  const documentRows = documentSheet.getDataRange().getValues();

  const employeeRows = employeeSheet.getDataRange().getValues();



  const employees = {};

const employeeCompanyIds = {};



employeeRows.slice(1).forEach(function(row) {

  const employeeId = String(row[0] || '').trim();



  if (!employeeId) return;



  employees[employeeId] = String(row[1] || '').trim();

  employeeCompanyIds[employeeId] =

    String(row[6] || '').trim();

});



  const members = {};

  familyRows.slice(1).forEach(function(row) {

    const id = String(row[0]).trim();

    const employeeId = String(row[1] || '').trim();



if (

  !isSuperAdmin_(user) &&

  employeeCompanyIds[employeeId] !==

    String(user.companyId || '').trim()

) return;



    if (id) {

      members[id] = {

        name: String(row[2]).trim(),

        relationship: String(row[3]).trim(),

        employeeId: String(row[1]).trim(),

        active: row[5] === true ||

          String(row[5]).trim().toUpperCase() === 'TRUE'

      };

    }

  });



  const documents = documentRows.slice(1).map(function(row) {

    const memberId = String(row[1]).trim();

    const member = members[memberId];



    if (!member || !member.active) return null;



    const expiry = row[6];

    let expiryDate = '';



    if (expiry instanceof Date && !isNaN(expiry.getTime())) {

      expiryDate = Utilities.formatDate(

        expiry,

        Session.getScriptTimeZone(),

        'yyyy-MM-dd'

      );

    } else if (expiry) {

      expiryDate = String(expiry);

    }



    let status = 'No expiry date';



    if (expiryDate) {

      const today = new Date();

      today.setHours(0, 0, 0, 0);



      const expiryDay = new Date(expiryDate + 'T00:00:00');

      expiryDay.setHours(0, 0, 0, 0);



      if (expiryDay < today) {

        status = 'Expired';

      } else if (

        Math.round((expiryDay.getTime() - today.getTime()) / 86400000) <= 90

      ) {

        status = 'Expiring soon';

      } else {

        status = 'Valid';

      }

    }



    return {

      familyDocumentId: String(row[0]).trim(),

      employeeId: member.employeeId,

      employeeName: employees[member.employeeId] || member.employeeId,

      familyMemberId: memberId,

      familyMemberName: member.name,

      relationship: member.relationship,

      documentType: String(row[2]).trim(),

      documentName: String(row[4]).trim(),

      expiryDate: expiryDate,

      status: status

    };

  }).filter(function(doc) {

    return doc !== null;

  });



  return { documents: documents };

}

// Validate a replacement document before saving it.

function prepareRenewalUpload_(fileName, mimeType, base64Data, issueDate, expiryDate) {

  fileName = String(fileName || '').trim().split(/[\\/]/).pop();

  fileName = fileName.replace(/[:*?"<>|]/g, '_');

  mimeType = String(mimeType || '').trim().toLowerCase();



  const allowedTypes = {

    'application/pdf': ['.pdf'],

    'image/jpeg': ['.jpg', '.jpeg'],

    'image/png': ['.png']

  };



  if (!fileName || !base64Data || !allowedTypes[mimeType]) {

    throw new Error('Choose a valid PDF, JPG, or PNG replacement file.');

  }



  const extension = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();



  if (!allowedTypes[mimeType].includes(extension)) {

    throw new Error('The file extension does not match the selected file type.');

  }



  if (fileName.length > 200) {

    throw new Error('The file name must be 200 characters or fewer.');

  }



  const bytes = Utilities.base64Decode(base64Data);



  if (bytes.length === 0 || bytes.length > 5 * 1024 * 1024) {

    throw new Error('The replacement file must be no larger than 5 MB.');

  }



  function parseRenewalDate_(value, label, required) {

    const text = String(value || '').trim();



    if (!text) {

      if (required) throw new Error(label + ' is required.');

      return '';

    }



    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {

      throw new Error('Enter a valid ' + label + '.');

    }



    const parts = text.split('-').map(Number);

    const date = new Date(parts[0], parts[1] - 1, parts[2]);



    if (

      date.getFullYear() !== parts[0] ||

      date.getMonth() !== parts[1] - 1 ||

      date.getDate() !== parts[2]

    ) {

      throw new Error('Enter a valid ' + label + '.');

    }



    return date;

  }



  const parsedIssueDate = parseRenewalDate_(

    issueDate, 'issue date', false

  );



  const parsedExpiryDate = parseRenewalDate_(

    expiryDate, 'new expiry date', true

  );



  if (parsedIssueDate && parsedExpiryDate < parsedIssueDate) {

    throw new Error('The new expiry date cannot be earlier than the issue date.');

  }



  return {

    fileName: fileName,

    mimeType: mimeType,

    bytes: bytes,

    issueDate: parsedIssueDate,

    expiryDate: parsedExpiryDate

  };

}





// Renew an employee document without creating a second document record.

// HR can renew any employee document. Employees can renew only their own.

function renewEmployeeDocument(

  documentId, issueDate, expiryDate, fileName, mimeType, base64Data

) {

  const user = getCurrentUser();



 if (!user || (!hasHrAccess_(user) && user.role !== 'Employee')) {

  throw new Error('You are not authorized to renew employee documents.');

}



  documentId = String(documentId || '').trim();



  if (!documentId) {

    throw new Error('Document ID is required.');

  }



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const documentsSheet = ss.getSheetByName('Documents');

  const employeesSheet = ss.getSheetByName('Employees');



  if (!documentsSheet || !employeesSheet) {

    throw new Error('The required document sheets could not be found.');

  }



  const documentRows = documentsSheet.getDataRange().getValues();

  let documentRow = null;

  let sheetRowNumber = -1;



  for (let i = 1; i < documentRows.length; i++) {

    if (String(documentRows[i][0] || '').trim() === documentId) {

      documentRow = documentRows[i].slice();

      sheetRowNumber = i + 1;

      break;

    }

  }



  if (!documentRow) {

    throw new Error('The document record could not be found.');

  }



  const documentEmployeeId = String(documentRow[1] || '').trim();

    // Verify that the document belongs to the signed-in user's company.

  const companyRows = employeesSheet.getDataRange().getValues();

  let documentCompanyId = '';



  for (let i = 1; i < companyRows.length; i++) {

    if (

      String(companyRows[i][0] || '').trim() === documentEmployeeId

    ) {

      documentCompanyId = String(companyRows[i][6] || '').trim();

      break;

    }

  }



  if (

  !documentCompanyId ||

  (

    !isSuperAdmin_(user) &&

    documentCompanyId !== String(user.companyId || '').trim()

  )

) {

  throw new Error(

    'You are not authorized to renew this document.'

  );

}



  if (user.role === 'Employee') {

    const employeeRows = employeesSheet.getDataRange().getValues();



    const authorizedEmployee = employeeRows.slice(1).some(function(row) {

      return (

        String(row[0] || '').trim() === documentEmployeeId &&

        String(row[3] || '').trim().toLowerCase() ===

          String(user.email || '').trim().toLowerCase() &&

        (row[4] === true ||

          String(row[4]).trim().toUpperCase() === 'TRUE')

      );

    });



    if (!authorizedEmployee) {

      throw new Error('You can renew only your own employee documents.');

    }

  }



  const upload = prepareRenewalUpload_(

    fileName, mimeType, base64Data, issueDate, expiryDate

  );



  const oldDriveFileId = String(documentRow[6] || '').trim();

  const newFile = DriveApp.getFolderById(config.folderId).createFile(

    Utilities.newBlob(upload.bytes, upload.mimeType, upload.fileName)

  );



  // Preserve the record ID and original creation date.

  documentRow[3] = upload.fileName;

  if (upload.issueDate) documentRow[4] = upload.issueDate;

  documentRow[5] = upload.expiryDate;

  documentRow[6] = newFile.getId();

  documentRow[8] = new Date();



  try {

    documentsSheet.getRange(sheetRowNumber, 1, 1, documentRow.length)

      .setValues([documentRow]);

  } catch (error) {

    newFile.setTrashed(true);

    throw new Error('The renewal could not be saved. The old document was kept.');

  }



  // The sheet now points to the replacement. Trash the old Drive file.

  if (oldDriveFileId && oldDriveFileId !== newFile.getId()) {

    try {

      DriveApp.getFileById(oldDriveFileId).setTrashed(true);

    } catch (error) {

      console.error('Renewal saved, but the previous file could not be trashed.');

    }

  }




  try {
    writeAuditLog_(
      ss,
      user.email,
      'EMPLOYEE_DOCUMENT_RENEWED',
      documentId,
      'New expiry date: ' +
        Utilities.formatDate(
          upload.expiryDate,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        ),
      documentCompanyId
    );
  } catch (auditError) {
    console.error('Could not write the document renewal audit entry.');
  }




  return {

    success: true,

    documentId: documentId,

    message: 'Document renewed successfully.'

  };

}

// Renew a family document. HR access only.

function renewFamilyDocument(

  familyDocumentId,

  issueDate,

  expiryDate,

  fileName,

  mimeType,

  base64Data

) {

  const user = getCurrentUser();



  if (!hasHrAccess_(user)) {
    throw new Error('Only authorized HR users can renew family documents.');
  }



  familyDocumentId = String(familyDocumentId || '').trim();



  if (!familyDocumentId) {

    throw new Error('Family document ID is required.');

  }



  const config = getPortalConfig_();

  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const documentsSheet = ss.getSheetByName('FamilyDocuments');

  const membersSheet = ss.getSheetByName('FamilyMembers');



  if (!documentsSheet || !membersSheet) {

    throw new Error('The family document sheets could not be found.');

  }



  const documentRows = documentsSheet.getDataRange().getValues();

  let documentRow = null;

  let sheetRowNumber = -1;



  for (let i = 1; i < documentRows.length; i++) {

    if (

      String(documentRows[i][0] || '').trim() === familyDocumentId

    ) {

      documentRow = documentRows[i].slice();

      sheetRowNumber = i + 1;

      break;

    }

  }



  if (!documentRow) {

    throw new Error('The family document record could not be found.');

  }



  // Confirm that the linked family member is active.

  const familyMemberId = String(documentRow[1] || '').trim();

  // Find the employee linked to this family member.

const familySheetForCompany =

  ss.getSheetByName('FamilyMembers');



const employeesSheetForCompany =

  ss.getSheetByName('Employees');



if (!familySheetForCompany || !employeesSheetForCompany) {

  throw new Error('Required company records could not be found.');

}



const familyRowsForCompany =

  familySheetForCompany.getDataRange().getValues();



let linkedEmployeeId = '';



for (let i = 1; i < familyRowsForCompany.length; i++) {

  if (

    String(familyRowsForCompany[i][0] || '').trim() === familyMemberId

  ) {

    linkedEmployeeId =

      String(familyRowsForCompany[i][1] || '').trim();

    break;

  }

}



if (!linkedEmployeeId) {

  throw new Error('The family member has no employee record.');

}



// Verify the employee's company.

const employeeRowsForCompany =

  employeesSheetForCompany.getDataRange().getValues();



let familyDocumentCompanyId = '';



for (let i = 1; i < employeeRowsForCompany.length; i++) {

  if (

    String(employeeRowsForCompany[i][0] || '').trim() === linkedEmployeeId

  ) {

    familyDocumentCompanyId =

      String(employeeRowsForCompany[i][6] || '').trim();

    break;

  }

}



if (

  !familyDocumentCompanyId ||

  (

    !isSuperAdmin_(user) &&

    familyDocumentCompanyId !== String(user.companyId || '').trim()

  )

) {

  throw new Error(

    'You are not authorized to renew this family document.'

  );

}

  const memberRows = membersSheet.getDataRange().getValues();



  const activeMemberFound = memberRows.slice(1).some(function(row) {

    const isActive =

      row[5] === true ||

      String(row[5] || '').trim().toUpperCase() === 'TRUE';



    return (

      String(row[0] || '').trim() === familyMemberId &&

      isActive

    );

  });



  if (!activeMemberFound) {

    throw new Error('The family member is inactive or could not be found.');

  }



  // Validate the replacement file and dates using Step 1's helper.

  const upload = prepareRenewalUpload_(

    fileName,

    mimeType,

    base64Data,

    issueDate,

    expiryDate

  );



  const oldDriveFileId = String(documentRow[7] || '').trim();



  // Create the replacement in the existing private Drive folder.

  const newFile = DriveApp.getFolderById(config.folderId).createFile(

    Utilities.newBlob(

      upload.bytes,

      upload.mimeType,

      upload.fileName

    )

  );



  // Preserve document ID, family member ID, and document number.

  documentRow[4] = upload.fileName;



  if (upload.issueDate) {

    documentRow[5] = upload.issueDate;

  }



  documentRow[6] = upload.expiryDate;

  documentRow[7] = newFile.getId();

  documentRow[9] = new Date();



  try {

    documentsSheet

      .getRange(sheetRowNumber, 1, 1, documentRow.length)

      .setValues([documentRow]);

  } catch (error) {

    newFile.setTrashed(true);

    throw new Error(

      'The renewal could not be saved. The previous document was kept.'

    );

  }



  // Remove the previous file only after the sheet update succeeds.

  if (oldDriveFileId && oldDriveFileId !== newFile.getId()) {

    try {

      DriveApp.getFileById(oldDriveFileId).setTrashed(true);

    } catch (error) {

      console.error(

        'Renewal saved, but the previous Drive file could not be trashed.'

      );

    }

  }


  // Record the family document renewal in the audit log.
  try {
    writeAuditLog_(
      ss,
      user.email,
      'FAMILY_DOCUMENT_RENEWED',
      familyDocumentId,
      'New expiry date: ' +
        Utilities.formatDate(
          upload.expiryDate,
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        ),
      familyDocumentCompanyId
    );
  } catch (auditError) {
    console.error('Could not write the family renewal audit entry.');
  }




  return {

    success: true,

    familyDocumentId: familyDocumentId,

    message: 'Family document renewed successfully.'

  };

}


function cleanupExpiryTestDocument() {
  const documentId = 'DOC-7093CC22';
  const expectedDriveFileId = '1Y2uxQijrufLw_MwgYSQ1bqNe2aIiOMO1';
  const expectedEmployeeId = 'EMP-12309798';
  const expectedFileName = 'Free_Test_Data_100KB_PDF.pdf';

  const user = getCurrentUser();
  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can run this cleanup.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('Documents');

  if (!sheet) {
    throw new Error('Documents sheet not found.');
  }

  const rows = sheet.getDataRange().getValues();
  let targetRow = -1;

  // Locate the exact test document record.
  for (let i = 1; i < rows.length; i++) {
    if (
      String(rows[i][0] || '').trim() === documentId &&
      String(rows[i][1] || '').trim() === expectedEmployeeId &&
      String(rows[i][3] || '').trim() === expectedFileName &&
      String(rows[i][6] || '').trim() === expectedDriveFileId
    ) {
      if (targetRow !== -1) {
        throw new Error(
          'Duplicate matching records found. Nothing was deleted.'
        );
      }

      targetRow = i + 1;
    }
  }

  if (targetRow === -1) {
    throw new Error(
      'Exact test record not found. Nothing was deleted.'
    );
  }

  // Find the test employee's company for the audit entry.
  const employeesSheet = ss.getSheetByName('Employees');
  if (!employeesSheet) {
    throw new Error('Employees sheet not found.');
  }

  const employeeRows = employeesSheet.getDataRange().getValues();
  let cleanupCompanyId = '';

  for (let i = 1; i < employeeRows.length; i++) {
    if (
      String(employeeRows[i][0] || '').trim() === expectedEmployeeId
    ) {
      cleanupCompanyId = String(employeeRows[i][6] || '').trim();
      break;
    }
  }

  if (!cleanupCompanyId) {
    throw new Error(
      'CompanyID for the test employee was not found. Nothing was deleted.'
    );
  }

  const file = DriveApp.getFileById(expectedDriveFileId);

  if (file.isTrashed()) {
    throw new Error(
      'The test file is already in the trash. Nothing was changed.'
    );
  }

  // Trash the exact test file first.
  file.setTrashed(true);

  try {
    sheet.deleteRow(targetRow);
  } catch (error) {
    // Restore the file if the sheet row could not be removed.
    file.setTrashed(false);
    throw new Error(
      'Could not remove the test row; the file was restored.'
    );
  }

  try {
    writeAuditLog_(
      ss,
      user.email,
      'TEST_DOCUMENT_CLEANUP',
      documentId,
      'Removed temporary expiry test document and trashed its Drive file.',
      cleanupCompanyId
    );
  } catch (auditError) {
    console.error('Test document cleanup audit logging failed.');
  }

  return 'Temporary expiry test document removed successfully.';
}


function initializeCompanyDocumentSheets() {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can initialize these sheets.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const definitions = {
    CompanyDocuments: [
      'CompanyDocumentID',
      'CompanyID',
      'Category',
      'DocumentName',
      'DocumentNumber',
      'IssueDate',
      'ExpiryDate',
      'DriveFileID',
      'CreatedAt',
      'UpdatedAt',
      'CreatedBy',
      'UpdatedBy',
      'Active'
    ],
    CompanyDocumentHistory: [
      'HistoryID',
      'CompanyDocumentID',
      'CompanyID',
      'Category',
      'DocumentName',
      'DocumentNumber',
      'IssueDate',
      'ExpiryDate',
      'DriveFileID',
      'RecordedAt',
      'RecordedBy',
      'ChangeType'
    ]
  };

  const createdSheets = [];

  Object.keys(definitions).forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, definitions[sheetName].length)
        .setValues([definitions[sheetName]]);
      sheet.setFrozenRows(1);
      createdSheets.push(sheetName);
      return;
    }

    // Initialize an entirely empty sheet, but never overwrite existing data.
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, definitions[sheetName].length)
        .setValues([definitions[sheetName]]);
      sheet.setFrozenRows(1);
      createdSheets.push(sheetName);
    }
  });

 
if (createdSheets.length > 0) {
  try {
    writeAuditLog_(
      ss,
      user.email,
      'COMPANY_DOCUMENT_SHEETS_INITIALIZED',
      '',
      'Created or initialized: ' + createdSheets.join(', ')
    );
  } catch (auditError) {
    console.error('Company document sheet initialization audit logging failed.');
  }
}


  return {
    success: true,
    message: createdSheets.length
      ? 'Initialized: ' + createdSheets.join(', ')
      : 'Both company-document sheets already exist. No changes were made.'
  };
}

const COMPANY_DOCUMENT_CATEGORIES = [
  'Visa',
  'Insurance',
  'Rent Agreement',
  'ID Card',
  'Fire & Safety Certificate',
  'Chamber of Commerce Certificate',
  'Vehicle Mulkiya',
  'Municipality Certificate',
  'Employee Contract',
  'Tourism Certificate',
  'Logo Renewal',
  'Tax Certificate',
  'Passport',
  'CR',
  'Others'
];

function parseCompanyDocumentDate_(value, fieldName) {
  if (!value) return '';

  const text = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    throw new Error('Invalid ' + fieldName + '.');
  }

  const parts = text.split('-').map(Number);
  const date = new Date(parts[0], parts[1] - 1, parts[2]);

  if (
    date.getFullYear() !== parts[0] ||
    date.getMonth() !== parts[1] - 1 ||
    date.getDate() !== parts[2]
  ) {
    throw new Error('Invalid ' + fieldName + '.');
  }

  return date;
}

function getCompanyDocumentExpiryStatus_(value) {
  if (!value) return 'No expiry date';

  let expiry;
  if (value instanceof Date) {
    expiry = new Date(value);
  } else {
    const text = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
      expiry = parseCompanyDocumentDate_(text, 'expiry date');
    } else {
      expiry = new Date(value);
    }
  }

  if (!expiry || isNaN(expiry.getTime())) return 'Unknown';

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  expiry.setHours(0, 0, 0, 0);

  const days = Math.round(
    (expiry.getTime() - today.getTime()) / 86400000
  );

  if (days < 0) return 'Expired';
  if (days <= 15) return 'Expires in 15 days';
  if (days <= 30) return 'Expires in 30 days';
  if (days <= 60) return 'Expires in 60 days';
  if (days <= 90) return 'Expires in 90 days';

  return 'Valid';
}

function getCompanyDocumentsData() {
  const user = getCurrentUser();
  if (!hasHrAccess_(user)) {
    throw new Error('Authorized HR access is required.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('CompanyDocuments');
  const companiesSheet = ss.getSheetByName('Companies');

  if (!sheet || !companiesSheet) {
    throw new Error('Company document sheets were not found.');
  }

  const companies = {};
  companiesSheet.getDataRange().getValues().slice(1).forEach(function(row) {
    const id = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const status = String(row[2] || '').trim().toUpperCase();
    if (id && name && status === 'ACTIVE') companies[id] = name;
  });

  return sheet.getDataRange().getValues().slice(1).reduce(function(result, row) {
    const id = String(row[0] || '').trim();
    const companyId = String(row[1] || '').trim();
    const active = row[12] === true ||
      String(row[12] || '').trim().toUpperCase() === 'TRUE';

    if (!id || !active || !companies[companyId]) return result;

    if (
      !isSuperAdmin_(user) &&
      companyId !== String(user.companyId || '').trim()
    ) return result;

    result.push({
      companyDocumentId: id,
      companyId: companyId,
      companyName: companies[companyId],
      category: String(row[2] || ''),
      name: String(row[3] || ''),
      documentNumber: String(row[4] || ''),
      issueDate: formatSheetDate_(row[5]),
      expiryDate: formatSheetDate_(row[6]),
      status: getCompanyDocumentExpiryStatus_(row[6])
    });

    return result;
  }, []);
}

function uploadCompanyDocument(
  targetCompanyId,
  category,
  documentName,
  documentNumber,
  issueDate,
  expiryDate,
  fileName,
  mimeType,
  base64Data
) {
  const user = getCurrentUser();
  if (!hasHrAccess_(user)) {
    throw new Error('Only authorized HR users can upload company documents.');
  }

  const companyId = isSuperAdmin_(user)
    ? String(targetCompanyId || '').trim()
    : String(user.companyId || '').trim();

  if (!companyId) throw new Error('A company must be selected.');
  category = String(category || '').trim();
  documentName = String(documentName || '').trim();
  documentNumber = String(documentNumber || '').trim();
  fileName = String(fileName || '').trim().split(/[\\/]/).pop();
  mimeType = String(mimeType || '').trim().toLowerCase();

  if (!isActiveCompany_(
    SpreadsheetApp.openById(getPortalConfig_().spreadsheetId), companyId
  )) {
    throw new Error('The selected company is not active.');
  }

  if (!COMPANY_DOCUMENT_CATEGORIES.includes(category)) {
    throw new Error('Select a valid company document category.');
  }
  if (!documentName || documentName.length > 200) {
    throw new Error('Enter a document name of 1–200 characters.');
  }
  if (documentNumber.length > 100) {
    throw new Error('Document number must be 100 characters or fewer.');
  }

  const allowedTypes = {
    'application/pdf': ['.pdf'],
    'image/jpeg': ['.jpg', '.jpeg'],
    'image/png': ['.png']
  };

  if (!allowedTypes[mimeType]) {
    throw new Error('Only PDF, JPG, and PNG files are allowed.');
  }

  const extension = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();
  if (!fileName || !allowedTypes[mimeType].includes(extension)) {
    throw new Error('The file extension does not match its type.');
  }

  const bytes = Utilities.base64Decode(base64Data || '');
  if (!bytes.length || bytes.length > 5 * 1024 * 1024) {
    throw new Error('The file must be larger than 0 bytes and no more than 5 MB.');
  }

  const parsedIssueDate = parseCompanyDocumentDate_(issueDate, 'issue date');
  const parsedExpiryDate = parseCompanyDocumentDate_(expiryDate, 'expiry date');

  if (parsedIssueDate && parsedExpiryDate &&
      parsedExpiryDate < parsedIssueDate) {
    throw new Error('Expiry date cannot be earlier than issue date.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('CompanyDocuments');
  const history = ss.getSheetByName('CompanyDocumentHistory');

  if (!sheet || !history) {
    throw new Error('Company document sheets were not found.');
  }

  const safeName = fileName.replace(/[:*?"<>|]/g, '_');
  const file = DriveApp.getFolderById(config.folderId).createFile(
    Utilities.newBlob(bytes, mimeType, safeName)
  );

  const now = new Date();
  const id = 'CDOC-' + Utilities.getUuid().substring(0, 8).toUpperCase();

  try {
    sheet.appendRow([
      id, companyId, category, documentName, documentNumber,
      parsedIssueDate, parsedExpiryDate, file.getId(),
      now, now, user.email, user.email, true
    ]);
  } catch (error) {
    file.setTrashed(true);
    throw new Error('Could not save the company document record.');
  }

  try {
    history.appendRow([
      'CDH-' + Utilities.getUuid().substring(0, 8).toUpperCase(),
      id, companyId, category, documentName, documentNumber,
      parsedIssueDate, parsedExpiryDate, file.getId(),
      now, user.email, 'CREATED'
    ]);
  } catch (error) {
    console.error('Company document saved, but history could not be recorded.');
  }

  const audit = ss.getSheetByName('AuditLog');
  if (audit) {
    try {
      audit.appendRow([
        now, user.email, 'COMPANY_DOCUMENT_ADDED', id,
        'CompanyID: ' + companyId + '; Category: ' + category
      ]);
    } catch (error) {
      console.error('Could not write company-document audit entry.');
    }
  }

  return { success: true, companyDocumentId: id, message: 'Company document uploaded.' };
}

function getAuthorizedCompanyDocumentFile(companyDocumentId) {
  const user = getCurrentUser();
  if (!hasHrAccess_(user)) {
    throw new Error('Authorized HR access is required.');
  }

  companyDocumentId = String(companyDocumentId || '').trim();
  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('CompanyDocuments');

  if (!sheet) throw new Error('CompanyDocuments sheet was not found.');

  const rows = sheet.getDataRange().getValues();
  const row = rows.slice(1).find(function(item) {
    return String(item[0] || '').trim() === companyDocumentId;
  });

  if (!row) throw new Error('Company document was not found.');

  const companyId = String(row[1] || '').trim();
  const active = row[12] === true ||
    String(row[12] || '').trim().toUpperCase() === 'TRUE';

  if (!active || !isActiveCompany_(ss, companyId)) {
    throw new Error('The company document is not active.');
  }
  if (
    !isSuperAdmin_(user) &&
    companyId !== String(user.companyId || '').trim()
  ) {
    throw new Error('You are not authorized to view this company document.');
  }

  const fileId = String(row[7] || '').trim();
  if (!fileId) throw new Error('The document file is unavailable.');

  const file = DriveApp.getFileById(fileId);
  if (file.isTrashed()) throw new Error('The document file is unavailable.');

  return {
    name: file.getName(),
    mimeType: file.getMimeType(),
    base64Data: Utilities.base64Encode(file.getBlob().getBytes())
  };
}

function renewCompanyDocument(
  companyDocumentId,
  issueDate,
  expiryDate,
  fileName,
  mimeType,
  base64Data
) {
  const user = getCurrentUser();

  if (!hasHrAccess_(user)) {
    throw new Error('Authorized HR access is required.');
  }

  companyDocumentId = String(companyDocumentId || '').trim();
  if (!companyDocumentId) {
    throw new Error('Company document ID is required.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('CompanyDocuments');
  const history = ss.getSheetByName('CompanyDocumentHistory');

  if (!sheet || !history) {
    throw new Error('Company document sheets were not found.');
  }

  const rows = sheet.getDataRange().getValues();
  let oldRow = null;
  let rowNumber = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === companyDocumentId) {
      oldRow = rows[i].slice();
      rowNumber = i + 1;
      break;
    }
  }

  if (!oldRow) {
    throw new Error('Company document was not found.');
  }

  const companyId = String(oldRow[1] || '').trim();
  const active = oldRow[12] === true ||
    String(oldRow[12] || '').trim().toUpperCase() === 'TRUE';

  if (!active) {
    throw new Error('This company document is inactive.');
  }

  if (!isActiveCompany_(ss, companyId)) {
    throw new Error('The company is inactive or was not found.');
  }

  if (
    !isSuperAdmin_(user) &&
    companyId !== String(user.companyId || '').trim()
  ) {
    throw new Error('You are not authorized to renew this company document.');
  }

  const upload = prepareRenewalUpload_(
    fileName, mimeType, base64Data, issueDate, expiryDate
  );

  const oldDriveFileId = String(oldRow[7] || '').trim();

  const newFile = DriveApp.getFolderById(config.folderId).createFile(
    Utilities.newBlob(upload.bytes, upload.mimeType, upload.fileName)
  );

  const now = new Date();
  const updatedRow = oldRow.slice();

  if (upload.issueDate) updatedRow[5] = upload.issueDate;
  updatedRow[6] = upload.expiryDate;
  updatedRow[7] = newFile.getId();
  updatedRow[9] = now;
  updatedRow[11] = user.email;

  // Save the replacement record first. Keep the old file if this fails.
  try {
    sheet.getRange(rowNumber, 1, 1, updatedRow.length)
      .setValues([updatedRow]);
  } catch (error) {
    newFile.setTrashed(true);
    throw new Error('Renewal could not be saved. The old file was kept.');
  }

  // Preserve the previous metadata and Drive file ID in renewal history.
  try {
    history.appendRow([
      'CDH-' + Utilities.getUuid().substring(0, 8).toUpperCase(),
      oldRow[0], oldRow[1], oldRow[2], oldRow[3], oldRow[4],
      oldRow[5], oldRow[6], oldRow[7], now, user.email, 'SUPERSEDED'
    ]);

    history.appendRow([
      'CDH-' + Utilities.getUuid().substring(0, 8).toUpperCase(),
      updatedRow[0], updatedRow[1], updatedRow[2], updatedRow[3],
      updatedRow[4], updatedRow[5], updatedRow[6], updatedRow[7],
      now, user.email, 'RENEWED'
    ]);
  } catch (error) {
    console.error(
      'Renewal saved, but the full history could not be recorded.'
    );
  }

  // The sheet now points to the replacement. Trash the old file afterward.
  if (oldDriveFileId && oldDriveFileId !== newFile.getId()) {
    try {
      DriveApp.getFileById(oldDriveFileId).setTrashed(true);
    } catch (error) {
      console.error(
        'Renewal saved, but the previous Drive file could not be trashed.'
      );
    }
  }

  const audit = ss.getSheetByName('AuditLog');
  if (audit) {
    try {
      audit.appendRow([
        now,
        user.email,
        'COMPANY_DOCUMENT_RENEWED',
        companyDocumentId,
        'CompanyID: ' + companyId +
          '; New expiry date: ' +
          Utilities.formatDate(
            upload.expiryDate,
            Session.getScriptTimeZone(),
            'yyyy-MM-dd'
          )
      ]);
    } catch (error) {
      console.error('Could not write the company renewal audit entry.');
    }
  }

  return {
    success: true,
    companyDocumentId: companyDocumentId,
    message: 'Company document renewed successfully.'
  };
}


function initializeExpiryAlertSettingsSheet() {
  const user = getCurrentUser();

  if (!user || !isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can initialize alert settings.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  let sheet = ss.getSheetByName('ExpiryAlertSettings');

  if (sheet) {
    console.log(
      'ExpiryAlertSettings already exists. No changes were made.'
    );
    return 'Settings sheet already exists; existing data was preserved.';
  }

  sheet = ss.insertSheet('ExpiryAlertSettings');

  sheet.appendRow([
    'Setting',
    'Value',
    'UpdatedAt',
    'UpdatedBy'
  ]);

  const now = new Date();
  const defaults = [
    ['ALERT_DAY_90', true, now, user.email],
    ['ALERT_DAY_60', true, now, user.email],
    ['ALERT_DAY_30', true, now, user.email],
    ['ALERT_DAY_15', true, now, user.email],
    ['ALERT_DAY_7', true, now, user.email],
    ['ALERT_DAY_0', true, now, user.email]
  ];

  sheet.getRange(2, 1, defaults.length, 4).setValues(defaults);
  sheet.setFrozenRows(1);

  console.log('ExpiryAlertSettings sheet created with default settings.');

  return 'Expiry alert settings initialized successfully.';
}

function getExpiryAlertSettings() {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can view expiry alert settings.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('ExpiryAlertSettings');

  if (!sheet) {
    throw new Error('ExpiryAlertSettings sheet was not found.');
  }

  const rows = sheet.getDataRange().getValues();
  const settings = {};

  [90, 60, 30, 15, 7, 0].forEach(function(day) {
    settings[day] = true;
  });

  rows.slice(1).forEach(function(row) {
    const key = String(row[0] || '').trim();
    const match = key.match(/^ALERT_DAY_(90|60|30|15|7|0)$/);

    if (match) {
      const day = Number(match[1]);
      const value = row[1];

      settings[day] =
        value === true ||
        String(value).trim().toUpperCase() === 'TRUE';
    }
  });

  return settings;
}


function saveExpiryAlertSettings(settings) {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can change expiry alert settings.');
  }

  if (!settings || typeof settings !== 'object' || Array.isArray(settings)) {
    throw new Error('Invalid expiry alert settings.');
  }

  const allowedDays = [90, 60, 30, 15, 7, 0];

  allowedDays.forEach(function(day) {
    if (typeof settings[day] !== 'boolean') {
      throw new Error('Invalid setting for the ' + day + '-day reminder.');
    }
  });

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('ExpiryAlertSettings');

  if (!sheet) {
    throw new Error('ExpiryAlertSettings sheet was not found.');
  }

  const rows = sheet.getDataRange().getValues();
  const rowByKey = {};

  rows.slice(1).forEach(function(row, index) {
    const key = String(row[0] || '').trim();
    if (key) rowByKey[key] = index + 2;
  });

  const now = new Date();

  allowedDays.forEach(function(day) {
    const key = 'ALERT_DAY_' + day;
    let rowNumber = rowByKey[key];

    if (!rowNumber) {
      rowNumber = sheet.getLastRow() + 1;
      sheet.getRange(rowNumber, 1).setValue(key);
    }

    sheet.getRange(rowNumber, 2, 1, 3).setValues([[
      settings[day],
      now,
      user.email
    ]]);
  });


try {
  writeAuditLog_(
    ss,
    user.email,
    'EXPIRY_ALERT_SETTINGS_UPDATED',
    '',
    'Enabled reminder days: ' +
      allowedDays.filter(function(day) {
        return settings[day];
      }).join(', ')
  );
} catch (auditError) {
  console.error('Expiry alert settings audit logging failed.');
}


  return {
    success: true,
    message: 'Expiry alert settings saved successfully.'
  };
}


function getUserRoleManagementData() {
  const currentUser = getCurrentUser();

  if (!isSuperAdmin_(currentUser)) {
    throw new Error('Only the Super Admin can view user role management.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const rolesSheet = ss.getSheetByName('Roles');
  const employeesSheet = ss.getSheetByName('Employees');
  const companiesSheet = ss.getSheetByName('Companies');

  if (!rolesSheet || !employeesSheet || !companiesSheet) {
    throw new Error('A required user-management sheet was not found.');
  }

  const isActive = function(value) {
    return value === true ||
      String(value || '').trim().toUpperCase() === 'TRUE';
  };

  const companyNames = {};

  companiesSheet.getDataRange().getValues().slice(1)
    .forEach(function(row) {
      const id = String(row[0] || '').trim();
      const name = String(row[1] || '').trim();
      const status = String(row[2] || '').trim().toUpperCase();

      if (id && name && status === 'ACTIVE') {
        companyNames[id] = name;
      }
    });

  const usersByEmail = {};

  // Start with registered employees.
  employeesSheet.getDataRange().getValues().slice(1)
    .forEach(function(row) {
      const email = String(row[3] || '').trim().toLowerCase();
      const employeeId = String(row[0] || '').trim();
      const companyId = String(row[6] || '').trim();

      if (!email) return;

      usersByEmail[email] = {
        email: email,
        name: String(row[1] || email).trim(),
        employeeId: employeeId,
        role: 'Employee',
        companyId: companyId,
        companyName: companyNames[companyId] || 'Unassigned',
        active: isActive(row[4])
      };
    });

  // Active Admin/HR roles take precedence, matching login checks.
  rolesSheet.getDataRange().getValues().slice(1)
    .forEach(function(row) {
      const email = String(row[0] || '').trim().toLowerCase();
      const role = String(row[1] || '').trim().toUpperCase();
      const active = isActive(row[2]);
      const companyId = String(row[3] || '').trim();

      if (!email ||
          (role !== 'SUPER_ADMIN' && role !== 'HR')) {
        return;
      }

      const existing = usersByEmail[email];

      // Preserve an employee's effective role if an inactive
      // Admin/HR assignment no longer grants elevated access.
      if (!active && existing) return;

      usersByEmail[email] = {
        email: email,
        name: existing ? existing.name : email,
        employeeId: existing ? existing.employeeId : '',
        role: role === 'SUPER_ADMIN' ? 'Admin' : 'HR',
        companyId: companyId,
        companyName: role === 'SUPER_ADMIN'
          ? 'All companies'
          : (companyNames[companyId] || 'Unassigned'),
        active: active
      };
    });

  const roleOrder = {
    Admin: 1,
    HR: 2,
    Employee: 3
  };

 
return Object.values(usersByEmail)
    .filter(function(user) {
      return user.active;
    })
    .sort(function(a, b) {
      if (roleOrder[a.role] !== roleOrder[b.role]) {
        return roleOrder[a.role] - roleOrder[b.role];
      }

      return a.name.localeCompare(b.name);
    });

}

function saveUserRoleManagementRole(email, role, companyId) {
  const currentUser = getCurrentUser();

  if (!isSuperAdmin_(currentUser)) {
    throw new Error('Only the Super Admin can change user roles.');
  }

  email = String(email || '').trim().toLowerCase();
  role = String(role || '').trim().toUpperCase();
  companyId = String(companyId || '').trim();

  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Enter a valid email address.');
  }

  if (email === currentUser.email) {
    throw new Error('You cannot change your own Super Admin role.');
  }

  if (role !== 'HR' && role !== 'EMPLOYEE') {
    throw new Error('Only HR or Employee roles can be assigned here.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const rolesSheet = ss.getSheetByName('Roles');
  const employeesSheet = ss.getSheetByName('Employees');

  if (!rolesSheet || !employeesSheet) {
    throw new Error('Required user-management sheets were not found.');
  }

  const roleRows = rolesSheet.getDataRange().getValues();
  const employeeRows = employeesSheet.getDataRange().getValues();

  const employee = employeeRows.slice(1).find(function(row) {
    return String(row[3] || '').trim().toLowerCase() === email &&
      (row[4] === true ||
       String(row[4] || '').trim().toUpperCase() === 'TRUE');
  });

  const existingRole = roleRows.slice(1).find(function(row) {
    return String(row[0] || '').trim().toLowerCase() === email &&
      String(row[1] || '').trim().toUpperCase() === 'SUPER_ADMIN';
  });

  if (existingRole) {
    throw new Error('Super Admin accounts cannot be changed here.');
  }

  if (role === 'HR') {
    if (!isActiveCompany_(ss, companyId)) {
      throw new Error('Select an active company for the HR account.');
    }

    const result = assignCompanyHrAccess(email, companyId);
    return result;
  }

  if (!employee) {
    throw new Error(
      'This email does not belong to an active registered employee.'
    );
  }

  let changed = false;

  for (let i = 1; i < roleRows.length; i++) {
    const rowEmail = String(roleRows[i][0] || '').trim().toLowerCase();
    const rowRole = String(roleRows[i][1] || '').trim().toUpperCase();
    const active = roleRows[i][2] === true ||
      String(roleRows[i][2] || '').trim().toUpperCase() === 'TRUE';

    if (rowEmail === email && rowRole === 'HR' && active) {
      rolesSheet.getRange(i + 1, 3).setValue(false);
      changed = true;
    }
  }


try {
 
writeAuditLog_(
  ss,
  currentUser.email,
  'USER_ROLE_CHANGED',
  String(employee[0] || ''),
  'Email: ' + email + '; Role: Employee' +
    (changed ? '; Previous HR access deactivated' : ''),
  String(employee[6] || '').trim()
);

} catch (auditError) {
  console.error('User role change audit logging failed.');
}


  return {
    success: true,
    email: email,
    role: 'Employee',
    message: 'Employee role saved successfully.'
  };
}


function initializeEmployeeGroupSheets() {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can initialize employee groups.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const definitions = {
    EmployeeGroups: [
      'GroupID',
      'GroupName',
      'CompanyID',
      'AssignedHrEmail',
      'Active',
      'CreatedAt',
      'CreatedBy',
      'UpdatedAt',
      'UpdatedBy'
    ],

    EmployeeGroupMembers: [
      'MembershipID',
      'GroupID',
      'EmployeeID',
      'Active',
      'CreatedAt',
      'CreatedBy',
      'UpdatedAt',
      'UpdatedBy'
    ]
  };

  const createdSheets = [];

  Object.keys(definitions).forEach(function(sheetName) {
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // Add headers only when the sheet is empty.
    // Never overwrite existing records.
    if (sheet.getLastRow() === 0) {
      const headers = definitions[sheetName];

      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers]);

      sheet.setFrozenRows(1);
      createdSheets.push(sheetName);
    }
  });

 
if (createdSheets.length > 0) {
  try {
    writeAuditLog_(
      ss,
      user.email,
      'EMPLOYEE_GROUP_SHEETS_INITIALIZED',
      '',
      'Initialized: ' + createdSheets.join(', ')
    );
  } catch (auditError) {
    console.error('Employee group sheet initialization audit logging failed.');
  }
}


  return {
    success: true,
    message: createdSheets.length > 0
      ? 'Initialized: ' + createdSheets.join(', ')
      : 'Group sheets already exist and were left unchanged.'
  };
}

function getEmployeeGroupManagementData() {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can manage employee groups.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const groupsSheet = ss.getSheetByName('EmployeeGroups');
  const rolesSheet = ss.getSheetByName('Roles');
  const companiesSheet = ss.getSheetByName('Companies');

  if (!groupsSheet || !rolesSheet || !companiesSheet) {
    throw new Error('A required group-management sheet is missing.');
  }

  const companies = companiesSheet.getDataRange().getValues()
    .slice(1)
    .map(function(row) {
      return {
        companyId: String(row[0] || '').trim(),
        companyName: String(row[1] || '').trim(),
        status: String(row[2] || '').trim().toUpperCase()
      };
    })
    .filter(function(company) {
      return company.companyId &&
        company.companyName &&
        company.status === 'ACTIVE';
    })
    .sort(function(a, b) {
      return a.companyName.localeCompare(b.companyName);
    });

  const companyNames = {};
  companies.forEach(function(company) {
    companyNames[company.companyId] = company.companyName;
  });

  const hrUsers = rolesSheet.getDataRange().getValues()
    .slice(1)
    .map(function(row) {
      const email = String(row[0] || '').trim().toLowerCase();
      const role = String(row[1] || '').trim().toUpperCase();
      const active = row[2] === true ||
        String(row[2] || '').trim().toUpperCase() === 'TRUE';
      const companyId = String(row[3] || '').trim();

      return {
        email: email,
        companyId: companyId,
        companyName: companyNames[companyId] || '',
        active: active,
        role: role
      };
    })
    .filter(function(hr) {
      return hr.email &&
        hr.role === 'HR' &&
        hr.active &&
        companyNames[hr.companyId];
    });

  const groups = groupsSheet.getDataRange().getValues()
    .slice(1)
    .map(function(row) {
      const companyId = String(row[2] || '').trim();

      return {
        groupId: String(row[0] || '').trim(),
        groupName: String(row[1] || '').trim(),
        companyId: companyId,
        companyName: companyNames[companyId] || 'Inactive or unavailable',
        assignedHrEmail: String(row[3] || '').trim(),
        active: row[4] === true ||
          String(row[4] || '').trim().toUpperCase() === 'TRUE'
      };
    })
    .filter(function(group) {
      return group.groupId && group.groupName;
    })
    .sort(function(a, b) {
      return a.groupName.localeCompare(b.groupName);
    });

  return {
    companies: companies,
    hrUsers: hrUsers,
    groups: groups
  };
}


function createEmployeeGroup(groupName, companyId) {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can create employee groups.');
  }

  groupName = String(groupName || '').trim();
  companyId = String(companyId || '').trim();

  if (groupName.length < 2 || groupName.length > 100) {
    throw new Error('Group name must contain 2 to 100 characters.');
  }

  if (!companyId) {
    throw new Error('Please select a company.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);

  const groupsSheet = ss.getSheetByName('EmployeeGroups');

  if (!groupsSheet) {
    throw new Error('EmployeeGroups sheet was not found.');
  }

  if (!isActiveCompany_(ss, companyId)) {
    throw new Error('Please select an active company.');
  }

  const rows = groupsSheet.getDataRange().getValues();

  const duplicate = rows.slice(1).some(function(row) {
    const existingName = String(row[1] || '').trim().toLowerCase();
    const existingCompany = String(row[2] || '').trim();
    const active = row[4] === true ||
      String(row[4] || '').trim().toUpperCase() === 'TRUE';

    return active &&
      existingCompany === companyId &&
      existingName === groupName.toLowerCase();
  });

  if (duplicate) {
    throw new Error('An active group with that name already exists for this company.');
  }

  const now = new Date();
  const groupId = 'GRP-' +
    Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();

  groupsSheet.appendRow([
    groupId,
    groupName,
    companyId,
    '',
    true,
    now,
    user.email,
    now,
    user.email
  ]);


try {
  
writeAuditLog_(
  ss,
  user.email,
  'EMPLOYEE_GROUP_CREATED',
  groupId,
  'Group: ' + groupName + '; CompanyID: ' + companyId,
  companyId
);

} catch (auditError) {
  console.error('Employee group creation audit logging failed.');
}


  return {
    success: true,
    groupId: groupId,
    message: 'Employee group created successfully.'
  };
}


function assignHrToEmployeeGroup(groupId, hrEmail) {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can assign HR to employee groups.');
  }

  groupId = String(groupId || '').trim();
  hrEmail = String(hrEmail || '').trim().toLowerCase();

  if (!groupId || !hrEmail) {
    throw new Error('Please select an employee group and an HR user.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const groupsSheet = ss.getSheetByName('EmployeeGroups');
  const rolesSheet = ss.getSheetByName('Roles');

  if (!groupsSheet || !rolesSheet) {
    throw new Error('EmployeeGroups or Roles sheet is missing.');
  }

  const groupRows = groupsSheet.getDataRange().getValues();
  let groupRowNumber = -1;
  let groupCompanyId = '';

  for (let i = 1; i < groupRows.length; i++) {
    if (
      String(groupRows[i][0]).trim() === groupId &&
      String(groupRows[i][4]).trim().toUpperCase() === 'TRUE'
    ) {
      groupRowNumber = i + 1;
      groupCompanyId = String(groupRows[i][2] || '').trim();
      break;
    }
  }

  if (groupRowNumber === -1) {
    throw new Error('Active employee group not found.');
  }

  const roleRows = rolesSheet.getDataRange().getValues();
  let validHr = false;

  for (let i = 1; i < roleRows.length; i++) {
    const email = String(roleRows[i][0] || '').trim().toLowerCase();
    const role = String(roleRows[i][1] || '').trim().toUpperCase();
    const active = String(roleRows[i][2] || '').trim().toUpperCase();
    const companyId = String(roleRows[i][3] || '').trim();


    if (
      email === hrEmail &&
      role === 'HR' &&
      companyId === groupCompanyId &&
      (active === 'TRUE' || active === 'ACTIVE')
    ) {
      validHr = true;
      break;
    }
  }

  if (!validHr) {
    throw new Error(
      'This HR user must be active and assigned to the same company as the group.'
    );
  }

  const now = new Date();

  groupsSheet.getRange(groupRowNumber, 4).setValue(hrEmail);
  groupsSheet.getRange(groupRowNumber, 8).setValue(now);
  groupsSheet.getRange(groupRowNumber, 9).setValue(user.email);

 
  try {
    
writeAuditLog_(
  ss,
  user.email,
  'EMPLOYEE_GROUP_HR_ASSIGNED',
  groupId,
  'Assigned HR: ' + hrEmail,
  groupCompanyId
);

  } catch (auditError) {
    console.error('HR group assignment audit logging failed.');
  }

  return {
    success: true,
    message: 'HR assigned to employee group successfully.'
  };
}


function assignEmployeeToGroup(groupId, employeeId) {
  const user = getCurrentUser();

if (!user || user.role !== 'SUPER_ADMIN') {
  throw new Error('Only the Super Admin can assign employees to groups.');
}

  groupId = String(groupId || '').trim();
  employeeId = String(employeeId || '').trim();

  if (!groupId || !employeeId) {
    throw new Error('Please select an employee group and an employee.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const groupsSheet = ss.getSheetByName('EmployeeGroups');
  const membersSheet = ss.getSheetByName('EmployeeGroupMembers');
  const employeesSheet = ss.getSheetByName('Employees');

  if (!groupsSheet || !membersSheet || !employeesSheet) {
    throw new Error('A required employee group or employee sheet is missing.');
  }

  const groupRows = groupsSheet.getDataRange().getValues();
  let groupCompanyId = '';

  for (let i = 1; i < groupRows.length; i++) {
    if (
      String(groupRows[i][0]).trim() === groupId &&
      String(groupRows[i][4]).trim().toUpperCase() === 'TRUE'
    ) {
      groupCompanyId = String(groupRows[i][2] || '').trim();
      break;
    }
  }

  if (!groupCompanyId) {
    throw new Error('Active employee group not found.');
  }

  const employeeRows = employeesSheet.getDataRange().getValues();
  let employeeExists = false;

  for (let i = 1; i < employeeRows.length; i++) {
    const row = employeeRows[i];
   
const rowEmployeeId = String(row[0] || '').trim();
const rowCompanyId = String(row[6] || '').trim();
const rowActive = String(row[4] || '').trim().toUpperCase();

    if (
      rowEmployeeId === employeeId &&
      rowCompanyId === groupCompanyId &&
      rowActive !== 'FALSE' &&
      rowActive !== 'INACTIVE'
    ) {
      employeeExists = true;
      break;
    }
  }

  if (!employeeExists) {
    throw new Error('Active employee not found in the group company.');
  }

  const memberRows = membersSheet.getDataRange().getValues();

  for (let i = 1; i < memberRows.length; i++) {
    if (
      String(memberRows[i][1]).trim() === groupId &&
      String(memberRows[i][2]).trim() === employeeId &&
      String(memberRows[i][3]).trim().toUpperCase() === 'TRUE'
    ) {
      throw new Error('This employee is already assigned to this group.');
    }
  }

  const now = new Date();
  const membershipId = 'MEM-' +
    Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();

  membersSheet.appendRow([
    membershipId,
    groupId,
    employeeId,
    true,
    now,
    user.email,
    now,
    user.email
  ]);

 
  try {
   
writeAuditLog_(
  ss,
  user.email,
  'EMPLOYEE_ASSIGNED_TO_GROUP',
  membershipId,
  'GroupID: ' + groupId + '; EmployeeID: ' + employeeId,
  groupCompanyId
);

  } catch (auditError) {
    console.error('Employee group assignment audit logging failed.');
  }


  return {
    success: true,
    message: 'Employee assigned to group successfully.'
  };
}

function getEmployeesForGroupAssignment() {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can manage employee group assignments.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const employeesSheet = ss.getSheetByName('Employees');

  if (!employeesSheet) {
    throw new Error('Employees sheet was not found.');
  }

  const rows = employeesSheet.getDataRange().getValues();

  const employees = rows.slice(1)
    .filter(function(row) {
      const employeeId = String(row[0] || '').trim();
      const name = String(row[1] || '').trim();
      const active = String(row[4] || '').trim().toUpperCase();
      const companyId = String(row[6] || '').trim();

      return employeeId &&
        name &&
        companyId &&
        active === 'TRUE';
    })
    .map(function(row) {
      return {
        employeeId: String(row[0]).trim(),
        name: String(row[1]).trim(),
        companyId: String(row[6]).trim()
      };
    });

  return employees;
}


function writeAuditLog_(ss, actorEmail, action, recordId, details, companyId) {
  const auditSheet = ss.getSheetByName('AuditLog');

  if (!auditSheet) {
    console.error('AuditLog sheet was not found.');
    return;
  }

  // Ensure the CompanyID header exists without changing existing records.
  const lastColumn = auditSheet.getLastColumn();

  if (
    lastColumn < 6 ||
    String(auditSheet.getRange(1, 6).getValue()).trim() !== 'CompanyID'
  ) {
    auditSheet.getRange(1, 6).setValue('CompanyID');
  }

  auditSheet.appendRow([
    new Date(),
    String(actorEmail || ''),
    String(action || ''),
    String(recordId || ''),
    String(details || ''),
    String(companyId || '')
  ]);
}




function prepareEmployeePasswordLogin() {
  const user = getCurrentUser();

  if (!isSuperAdmin_(user)) {
    throw new Error('Only the Super Admin can prepare employee login fields.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('Employees');

  if (!sheet) {
    throw new Error('Employees sheet was not found.');
  }

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });

  const requiredHeaders = [
    'PasswordHash',
    'MustChangePassword',
    'PasswordUpdatedAt'
  ];

  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  });

  return {
    success: true,
    message: 'Employee password fields are ready. Existing employee records were preserved.'
  };
}

function getEmployeePasswordColumnMap_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });

  const required = [
    'EmployeeID',
    'Name',
    'Active',
    'CompanyID',
    'PasswordHash',
    'MustChangePassword',
    'PasswordUpdatedAt'
  ];

  const map = {};

  required.forEach(function(header) {
    const index = headers.indexOf(header);
    if (index === -1) {
      throw new Error('Missing Employees column: ' + header);
    }
    map[header] = index;
  });

  return map;
}


function getEmployeePasswordSecret_() {
  const properties = PropertiesService.getScriptProperties();
  let secret = properties.getProperty('EMPLOYEE_PASSWORD_SECRET');

  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid() +
      Utilities.getUuid();
    properties.setProperty('EMPLOYEE_PASSWORD_SECRET', secret);
  }

  return secret;
}


function hashEmployeePassword_(password, salt) {
  const secret = getEmployeePasswordSecret_();
  const value = secret + '|' + salt + '|' + password;

  const digest = Utilities.computeHmacSha256Signature(value, secret);

  return digest.map(function(byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join('');
}


function setEmployeePassword(employeeId, newPassword) {
  const user = getCurrentUser();

  if (!hasHrAccess_(user)) {
    throw new Error('Only authorized HR users can set employee passwords.');
  }

  employeeId = String(employeeId || '').trim();
  newPassword = String(newPassword || '');

  if (!employeeId) {
    throw new Error('Employee ID is required.');
  }

  if (newPassword.length < 12) {
    throw new Error('Password must contain at least 12 characters.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('Employees');

  if (!sheet) {
    throw new Error('Employees sheet was not found.');
  }

  const map = getEmployeePasswordColumnMap_(sheet);
  const rows = sheet.getDataRange().getValues();
  let rowNumber = -1;
  let employeeCompanyId = '';

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][map.EmployeeID] || '').trim() !== employeeId) {
      continue;
    }

    const active = rows[i][map.Active] === true ||
      String(rows[i][map.Active] || '').trim().toUpperCase() === 'TRUE';

    if (!active) {
      throw new Error('This employee account is inactive.');
    }

    employeeCompanyId = String(rows[i][map.CompanyID] || '').trim();
    rowNumber = i + 1;
    break;
  }

  if (rowNumber === -1) {
    throw new Error('Employee ID was not found.');
  }

  if (!isSuperAdmin_(user) &&
      employeeCompanyId !== String(user.companyId || '').trim()) {
    throw new Error('You cannot set a password for another company.');
  }

  const salt = Utilities.getUuid() + Utilities.getUuid();
  const passwordHash = salt + ':' + hashEmployeePassword_(newPassword, salt);

  sheet.getRange(rowNumber, map.PasswordHash + 1).setValue(passwordHash);
  sheet.getRange(rowNumber, map.MustChangePassword + 1).setValue(true);
  sheet.getRange(rowNumber, map.PasswordUpdatedAt + 1).setValue(new Date());

  return {
    success: true,
    message: 'Temporary password set. The employee must change it after signing in.'
  };
}


function employeePasswordLogin(employeeId, password) {
  employeeId = String(employeeId || '').trim().toUpperCase();
  password = String(password || '');

  if (!employeeId || !password) {
    throw new Error('Enter your Employee ID and password.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('Employees');

  if (!sheet) {
    throw new Error('Employees sheet was not found.');
  }

  const map = getEmployeePasswordColumnMap_(sheet);
  const rows = sheet.getDataRange().getValues();
  let employee = null;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][map.EmployeeID] || '').trim().toUpperCase()
        !== employeeId) {
      continue;
    }

    const active = rows[i][map.Active] === true ||
      String(rows[i][map.Active] || '').trim().toUpperCase() === 'TRUE';

    if (!active) {
      throw new Error('Invalid Employee ID or password.');
    }

    employee = {
      rowNumber: i + 1,
      employeeId: employeeId,
      name: String(rows[i][map.Name] || employeeId),
      companyId: String(rows[i][map.CompanyID] || '').trim(),
      passwordHash: String(rows[i][map.PasswordHash] || ''),
      mustChangePassword: rows[i][map.MustChangePassword] === true ||
        String(rows[i][map.MustChangePassword] || '').trim().toUpperCase() === 'TRUE'
    };

    break;
  }

  if (!employee || !employee.passwordHash || !employee.companyId) {
    throw new Error('Invalid Employee ID or password.');
  }

  if (!isActiveCompany_(ss, employee.companyId)) {
    throw new Error('This employee account is not assigned to an active company.');
  }

  const parts = employee.passwordHash.split(':');

  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    throw new Error('A password has not been configured for this account.');
  }

  const expectedHash = hashEmployeePassword_(password, parts[0]);

  if (expectedHash !== parts[1]) {
    throw new Error('Invalid Employee ID or password.');
  }

  const token = Utilities.getUuid() + Utilities.getUuid() +
    Utilities.getUuid();

  const sessionUser = {
    employeeId: employee.employeeId,
    name: employee.name,
    role: 'Employee',
    companyId: employee.companyId,
    mustChangePassword: employee.mustChangePassword
  };

  CacheService.getScriptCache().put(
    'EMPLOYEE_SESSION_' + token,
    JSON.stringify(sessionUser),
    21600
  );

  return {
    success: true,
    token: token,
    user: sessionUser,
    mustChangePassword: employee.mustChangePassword
  };
}


function getEmployeeSessionUser_(token) {
  token = String(token || '').trim();

  if (!token || token.length > 200) {
    throw new Error('Your session has expired. Please sign in again.');
  }

  const cached = CacheService.getScriptCache().get(
    'EMPLOYEE_SESSION_' + token
  );

  if (!cached) {
    throw new Error('Your session has expired. Please sign in again.');
  }

  const sessionUser = JSON.parse(cached);
  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('Employees');

  if (!sheet) {
    throw new Error('Employees sheet was not found.');
  }

  const map = getEmployeePasswordColumnMap_(sheet);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][map.EmployeeID] || '').trim() !==
        String(sessionUser.employeeId || '').trim()) {
      continue;
    }

    const active = rows[i][map.Active] === true ||
      String(rows[i][map.Active] || '').trim().toUpperCase() === 'TRUE';

    const companyId = String(rows[i][map.CompanyID] || '').trim();

    if (!active || companyId !== sessionUser.companyId ||
        !isActiveCompany_(ss, companyId)) {
      CacheService.getScriptCache().remove('EMPLOYEE_SESSION_' + token);
      throw new Error('This employee session is no longer authorized.');
    }

  sessionUser.mustChangePassword =
  rows[i][map.MustChangePassword] === true ||
  String(rows[i][map.MustChangePassword] || '')
    .trim()
    .toUpperCase() === 'TRUE';

CacheService.getScriptCache().put(
  'EMPLOYEE_SESSION_' + token,
  JSON.stringify(sessionUser),
  21600
);

return sessionUser;
  }

  CacheService.getScriptCache().remove('EMPLOYEE_SESSION_' + token);
  throw new Error('Employee account was not found.');
}


function employeePasswordLogout(token) {
  token = String(token || '').trim();

  if (token && token.length <= 200) {
    CacheService.getScriptCache().remove('EMPLOYEE_SESSION_' + token);
  }

  return { success: true };
}

function changeEmployeePassword(sessionToken, currentPassword, newPassword) {
  const user = getEmployeeSessionUser_(sessionToken);

  if (!user || user.role !== 'Employee' || !user.employeeId) {
    throw new Error('Employee access is required.');
  }

  currentPassword = String(currentPassword || '');
  newPassword = String(newPassword || '');

  if (!currentPassword || newPassword.length < 12) {
    throw new Error('The new password must contain at least 12 characters.');
  }

  if (currentPassword === newPassword) {
    throw new Error('Choose a different password.');
  }

  const config = getPortalConfig_();
  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = ss.getSheetByName('Employees');

  if (!sheet) {
    throw new Error('Employees sheet was not found.');
  }

  const map = getEmployeePasswordColumnMap_(sheet);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (
      String(rows[i][map.EmployeeID] || '').trim() !==
      String(user.employeeId).trim()
    ) {
      continue;
    }

    const active = rows[i][map.Active] === true ||
      String(rows[i][map.Active] || '').trim().toUpperCase() === 'TRUE';

    if (!active ||
        String(rows[i][map.CompanyID] || '').trim() !==
        String(user.companyId || '').trim()) {
      throw new Error('This employee account is not authorized.');
    }

    const stored = String(rows[i][map.PasswordHash] || '').split(':');

    if (
      stored.length !== 2 ||
      hashEmployeePassword_(currentPassword, stored[0]) !== stored[1]
    ) {
      throw new Error('Your current password is incorrect.');
    }

    const salt = Utilities.getUuid() + Utilities.getUuid();
    const newHash = salt + ':' + hashEmployeePassword_(newPassword, salt);

    sheet.getRange(i + 1, map.PasswordHash + 1).setValue(newHash);
    sheet.getRange(i + 1, map.MustChangePassword + 1).setValue(false);
    sheet.getRange(i + 1, map.PasswordUpdatedAt + 1).setValue(new Date());
    const sessionKey = 'EMPLOYEE_SESSION_' + sessionToken;
const cachedSession = CacheService.getScriptCache().get(sessionKey);

if (cachedSession) {
  const updatedSession = JSON.parse(cachedSession);
  updatedSession.mustChangePassword = false;

  CacheService.getScriptCache().put(
    sessionKey,
    JSON.stringify(updatedSession),
    21600
  );
}

    return {
      success: true,
      message: 'Password changed successfully.'
    };
  }

  throw new Error('Employee account was not found.');
}
function diagnosePortalTestEmployee() {
  const employeeId = 'EMP-12309798';
  const documents = getEmployeeDocuments_(employeeId);

  Logger.log('Employee ID: ' + employeeId);
  Logger.log('Documents found: ' + documents.length);
  Logger.log(JSON.stringify(documents));
}

