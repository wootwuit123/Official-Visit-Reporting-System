// --- GLOBAL CONFIGURATION ---
const IMAGE_FOLDER_ID = '1Hmxo6lWzy9uQyh668Oa8M_1uTt-lPhSU-j9cmM0WNLs'; // เปลี่ยนเป็น ID ของโฟลเดอร์ Google Drive ของคุณ
const USERS_SHEET_NAME = 'users';
const REPORTS_SHEET_NAME = 'reports';

// ตั้งค่าโรงเรียน - แก้ไขตรงนี้
const SCHOOL_CONFIG = {
  name: 'โรงเรียนตัวอย่าง',
  email_domain: '@school.ac.th',
  district: 'สพป.จังหวัด เขต 1',
  province: 'จังหวัด'
};

// --- MAIN WEB APP FUNCTION ---
function doGet(e) {
  // Serve the HTML file when the web app URL is accessed.
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle(`ระบบรายงานการไปราชการ - ${SCHOOL_CONFIG.name}`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- SHEET & DATA SETUP ---

/**
 * Checks if required sheets and headers exist, creates them if not.
 */
function checkAndSetupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup Users Sheet
  const usersSheet = ss.getSheetByName(USERS_SHEET_NAME) || ss.insertSheet(USERS_SHEET_NAME);
  if (usersSheet.getLastRow() === 0) {
    const headers = ['id', 'name', 'email', 'password', 'role'];
    usersSheet.appendRow(headers);
    Logger.log(`Created headers in '${USERS_SHEET_NAME}' sheet.`);
  }

  // Setup Reports Sheet
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET_NAME) || ss.insertSheet(REPORTS_SHEET_NAME);
  if (reportsSheet.getLastRow() === 0) {
    const headers = ['id', 'userId', 'title', 'objective', 'datetime', 'location', 'travel', 'summary', 'images', 'submittedAt', 'status', 'comments'];
    reportsSheet.appendRow(headers);
    Logger.log(`Created headers in '${REPORTS_SHEET_NAME}' sheet.`);
  }
}

/**
 * Populates the sheets with sample data for testing.
 * Run this function manually from the Apps Script editor.
 */
function createSampleData() {
  checkAndSetupSheets(); // Ensure sheets and headers are ready
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET_NAME);

  // Clear existing data (except headers)
  if(usersSheet.getLastRow() > 1) usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, usersSheet.getLastColumn()).clearContent();
  if(reportsSheet.getLastRow() > 1) reportsSheet.getRange(2, 1, reportsSheet.getLastRow() - 1, reportsSheet.getLastColumn()).clearContent();

  // Sample Users - แก้ไขชื่อและอีเมลตามโรงเรียนของคุณ
  const users = [
    [1, 'อ.สมชาย ใจดี', `teacher1${SCHOOL_CONFIG.email_domain}`, 'password123', 'teacher'],
    [2, 'อ.สมศรี มีสุข', `teacher2${SCHOOL_CONFIG.email_domain}`, 'password123', 'teacher'],
    [3, 'อ.มานะ อดทน', `teacher3${SCHOOL_CONFIG.email_domain}`, 'password123', 'teacher'],
    [4, 'ผอ.เก่งกาจ สามารถ', `director${SCHOOL_CONFIG.email_domain}`, 'password123', 'director']
  ];
  usersSheet.getRange(2, 1, users.length, users[0].length).setValues(users);
  Logger.log('Sample users created.');

  // Sample Reports - แก้ไขข้อมูลตัวอย่างตามพื้นที่ของคุณ
  const reports = [
    [1, 1, 'ประชุมผู้บริหารสถานศึกษา', 'รับนโยบายการศึกษาใหม่', '2025-07-15T09:00', SCHOOL_CONFIG.district, 'เดินทางโดยรถยนต์ส่วนตัว', 'ได้รับทราบนโยบายและแนวทางการปฏิบัติงานใหม่ จะนำมาปรับใช้ในโรงเรียนต่อไป', JSON.stringify([]), new Date('2025-07-16T10:00:00Z').toISOString(), '✅ ผอ.รับทราบแล้ว', JSON.stringify([{ userId: 4, name: 'ผอ.เก่งกาจ สามารถ', text: 'รับทราบครับ', timestamp: new Date('2025-07-17T11:00:00Z').toISOString() }])],
    [2, 2, 'อบรมการใช้งานระบบ e-GP', 'เรียนรู้ระบบจัดซื้อจัดจ้างภาครัฐ', '2025-07-20T08:30', `โรงแรม${SCHOOL_CONFIG.province}`, 'เดินทางโดยรถโรงเรียน', 'เข้าใจขั้นตอนการทำงานมากขึ้น สามารถดำเนินการจัดซื้อได้ถูกต้อง', JSON.stringify([]), new Date('2025-07-21T14:00:00Z').toISOString(), '⏳ รอตรวจสอบ', JSON.stringify([])],
    [3, 1, 'นำนักเรียนแข่งขันทักษะวิชาการ', 'เข้าร่วมการแข่งขันระดับภาค', '2025-06-10T07:00', `มหาวิทยาลัย${SCHOOL_CONFIG.province}`, 'เดินทางโดยรถบัส', 'นักเรียนได้รับรางวัลชนะเลิศ 1 รายการ และรองชนะเลิศ 2 รายการ', JSON.stringify([]), new Date('2025-06-11T16:00:00Z').toISOString(), '💬 มีข้อเสนอแนะ', JSON.stringify([{ userId: 4, name: 'ผอ.เก่งกาจ สามารถ', text: 'ยอดเยี่ยมมากครับ ช่วยสรุปค่าใช้จ่ายแนบมาด้วยนะครับ', timestamp: new Date('2025-06-12T09:30:00Z').toISOString() }])]
  ];
  reportsSheet.getRange(2, 1, reports.length, reports[0].length).setValues(reports);
  Logger.log('Sample reports created.');
}

// --- HELPER for Data Conversion ---
function sheetDataToObjects(data) {
  if (data.length < 2) return [];
  const headers = data[0];
  const objects = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      let value = row[i];
      // FIX: Check if the value is a Date object and convert it to a serializable ISO string.
      if (value instanceof Date) {
        value = value.toISOString();
      }
      
      // Try to parse JSON strings for images and comments
      if ((header === 'images' || header === 'comments') && typeof value === 'string' && value.startsWith('[')) {
        try {
          obj[header] = JSON.parse(value);
        } catch (e) {
          obj[header] = []; // Default to empty array on parse error
        }
      } else {
        obj[header] = value;
      }
    });
    return obj;
  });
  return objects;
}

// --- FRONTEND API FUNCTIONS ---

/**
 * Logs in a user based on credentials.
 * @param {object} credentials - The user's email, password, and role.
 * @returns {object} A response object with success status and user data or error message.
 */
function loginUser(credentials) {
  let response = { success: false, message: 'อีเมล, รหัสผ่าน หรือบทบาทไม่ถูกต้อง' };
  try {
    checkAndSetupSheets(); // Ensure sheets exist before proceeding
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
    const usersData = usersSheet.getDataRange().getValues();
    const users = sheetDataToObjects(usersData);

    const foundUser = users.find(u => 
      u.email === credentials.email && 
      u.password === credentials.password && 
      u.role === credentials.role
    );

    if (foundUser) {
      const { password, ...userSafe } = foundUser;
      response = { success: true, user: userSafe };
    }
  } catch (e) {
    Logger.log(`Error in loginUser: ${e.message}`);
    response.message = `เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ${e.message}`;
  }
  return response;
}

/**
 * Gets all necessary data for the initial app load.
 * @returns {object} An object containing lists of all users and reports.
 */
function getInitialData() {
   let response = { users: [], reports: [] }; // Default response
   try {
    checkAndSetupSheets(); // Ensure sheets exist before proceeding
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
    const reportsSheet = ss.getSheetByName(REPORTS_SHEET_NAME);
    
    const usersData = usersSheet.getDataRange().getValues();
    const reportsData = reportsSheet.getDataRange().getValues();

    const users = sheetDataToObjects(usersData).map(({ password, ...user }) => user); // Exclude passwords
    const reports = sheetDataToObjects(reportsData);

    response = { users: users, reports: reports };
  } catch (e) {
    Logger.log(`Error in getInitialData: ${e.message}`);
    // On error, the default empty response is returned.
  }
  return response; // Always return a valid object
}

/**
 * Submits a new report, including handling image uploads.
 * @param {object} reportData - The report data from the frontend form.
 * @returns {object} A response object with success status and the new report data.
 */
function submitReport(reportData) {
  let response = { success: false, message: 'เกิดข้อผิดพลาดในการบันทึกรายงาน' };
  try {
    checkAndSetupSheets(); // Ensure sheets exist before proceeding
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportsSheet = ss.getSheetByName(REPORTS_SHEET_NAME);
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);

    // 1. Handle Image Uploads
    const uploadedImageInfo = [];
    if (reportData.images && reportData.images.length > 0) {
      reportData.images.forEach(img => {
        const decoded = Utilities.base64Decode(img.base64);
        const blob = Utilities.newBlob(decoded, img.type, img.name);
        const file = folder.createFile(blob);
        
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        uploadedImageInfo.push({
          id: file.getId(),
          name: file.getName(),
          type: file.getMimeType(),
          // FIX: Use a more reliable URL format for direct image embedding
          url: `https://lh3.googleusercontent.com/d/${file.getId()}`
        });
      });
    }

    // 2. Prepare data for sheet
    const newId = Date.now();
    const newReport = {
      id: newId,
      userId: reportData.userId,
      title: reportData.title,
      objective: reportData.objective,
      datetime: reportData.datetime,
      location: reportData.location,
      travel: reportData.travel,
      summary: reportData.summary,
      images: uploadedImageInfo,
      submittedAt: new Date().toISOString(),
      status: '⏳ รอตรวจสอบ',
      comments: []
    };

    const newRow = [
      newReport.id, newReport.userId, newReport.title, newReport.objective,
      newReport.datetime, newReport.location, newReport.travel, newReport.summary,
      JSON.stringify(newReport.images), newReport.submittedAt, newReport.status,
      JSON.stringify(newReport.comments)
    ];

    // 3. Append to sheet
    reportsSheet.appendRow(newRow);

    // 4. Return success response
    response = { success: true, newReport: newReport };

  } catch (e) {
    Logger.log(`Error in submitReport: ${e.message}`);
    response.message = `เกิดข้อผิดพลาดในการบันทึกรายงาน: ${e.message}`;
  }
  return response;
}

/**
 * Saves feedback from the director on a report.
 * @param {object} feedbackData - Contains reportId, commentText, and isApproved status.
 * @returns {object} The fully updated report object.
 */
function saveDirectorFeedback(feedbackData) {
  try {
    checkAndSetupSheets(); // Ensure sheets exist before proceeding
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REPORTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const reportIdCol = headers.indexOf('id') + 1;
    const statusCol = headers.indexOf('status') + 1;
    const commentsCol = headers.indexOf('comments') + 1;

    let reportRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][reportIdCol - 1] == feedbackData.reportId) {
        reportRowIndex = i + 1;
        break;
      }
    }

    if (reportRowIndex === -1) {
      throw new Error('ไม่พบรายงานที่ต้องการอัปเดต');
    }

    const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
    const users = sheetDataToObjects(usersSheet.getDataRange().getValues());
    const director = users.find(u => u.role === 'director');

    let comments = JSON.parse(sheet.getRange(reportRowIndex, commentsCol).getValue() || '[]');
    if (feedbackData.commentText && director) {
      comments.push({
        userId: director.id,
        name: director.name,
        text: feedbackData.commentText,
        timestamp: new Date().toISOString()
      });
    }
    sheet.getRange(reportRowIndex, commentsCol).setValue(JSON.stringify(comments));

    let currentStatus = sheet.getRange(reportRowIndex, statusCol).getValue();
    let newStatus = currentStatus;
    if (feedbackData.isApproved) {
      newStatus = '✅ ผอ.รับทราบแล้ว';
    } else {
      if (feedbackData.commentText && currentStatus !== '💬 มีข้อเสนอแนะ') {
        newStatus = '💬 มีข้อเสนอแนะ';
      } else if (!feedbackData.commentText && currentStatus === '✅ ผอ.รับทราบแล้ว') {
        newStatus = '⏳ รอตรวจสอบ';
      }
    }
    sheet.getRange(reportRowIndex, statusCol).setValue(newStatus);

    const updatedRow = sheet.getRange(reportRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const updatedReport = sheetDataToObjects([headers, updatedRow])[0];

    return updatedReport;

  } catch (e) {
    Logger.log(e);
    throw new Error(`เกิดข้อผิดพลาดในการบันทึกความเห็น: ${e.message}`);
  }
}

/**
 * ฟังก์ชันสำหรับอัปเดตการตั้งค่าโรงเรียน
 * @param {object} config - ข้อมูลการตั้งค่าใหม่
 */
function updateSchoolConfig(config) {
  // คุณสามารถบันทึกการตั้งค่าไว้ใน Properties Service หรือ Sheet แยกต่างหาก
  const properties = PropertiesService.getScriptProperties();
  properties.setProperties({
    'SCHOOL_NAME': config.name || SCHOOL_CONFIG.name,
    'SCHOOL_EMAIL_DOMAIN': config.email_domain || SCHOOL_CONFIG.email_domain,
    'SCHOOL_DISTRICT': config.district || SCHOOL_CONFIG.district,
    'SCHOOL_PROVINCE': config.province || SCHOOL_CONFIG.province
  });
  
  return { success: true, message: 'อัปเดตการตั้งค่าโรงเรียนเรียบร้อย' };
}

/**
 * ฟังก์ชันสำหรับดึงการตั้งค่าโรงเรียน
 */
function getSchoolConfig() {
  const properties = PropertiesService.getScriptProperties();
  const config = {
    name: properties.getProperty('SCHOOL_NAME') || SCHOOL_CONFIG.name,
    email_domain: properties.getProperty('SCHOOL_EMAIL_DOMAIN') || SCHOOL_CONFIG.email_domain,
    district: properties.getProperty('SCHOOL_DISTRICT') || SCHOOL_CONFIG.district,
    province: properties.getProperty('SCHOOL_PROVINCE') || SCHOOL_CONFIG.province
  };
  return config;
}
