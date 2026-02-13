

// -----------------------------------------------------------------------
// CẤU HÌNH & THIẾT LẬP (CONFIGURATION)
// -----------------------------------------------------------------------

// ID của file Google Sheet "Choice-Management"
const SPREADSHEET_ID = '1xalAKVKbLBUfjv8uKuYluyqpizl2fU9b9WfmHzj4HWQ';
// ID của file Google Sheet "DISC-Test" (Cấu hình & Câu hỏi)
const DISC_SPREADSHEET_ID = '1WxwDJF0cdJMBu9zptD-7S4QU7gHPM7lwBOcDxRdBQMQ';
// ID của file Google Sheet "InHouse-History" (Lưu lịch sử kết quả & Bài làm)
const HISTORY_SPREADSHEET_ID = '1Cu2B78YxjGhomZiPdsI5_CN7mB9_M_qzJHnNxcQsgsc';
// ID của file Google Sheet "Acc-Control"
const ACC_SPREADSHEET_ID = '13XQmkMf6FZDatJsph24i8Z2H0-Zw9D6E5r0yMxPyCSU';

// Tên các Sheet
const DB_SHEETS = {
  ACC: 'Acc-Management',        
  TESTS: 'Tests-Management',    
  BANK: 'Test-Bank',            
  RESULTS: 'Result-Test-History', // Updated Sheet Name
  QUIZ_ANSWERS: 'Test-Sheet-Answers-Record', // New Sheet for Detailed Answers
  DISC_TESTS: 'DISC-Tests-Management',
  DISC_QUESTIONS: 'DISC-Questions',
  DISC_ANSWERS: 'DISC-Answer-Sheet-Record',
  DISC_RESULTS: 'DISC-Result-History',
  ROLES: 'SOL-Role-List'
};

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getDiscSS() {
  return SpreadsheetApp.openById(DISC_SPREADSHEET_ID);
}

function getHistorySS() {
  return SpreadsheetApp.openById(HISTORY_SPREADSHEET_ID);
}

// VIEW: Trả về HTML khi truy cập trực tiếp bằng trình duyệt
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Hệ Thống Thi Trắc Nghiệm & DISC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// API: Xử lý request từ bên ngoài (AI Studio / Postman / Mobile App)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const args = data.args || [];
    let result;

    switch (action) {
      case 'getTestsList':
        result = getTestsList();
        break;
      case 'validateAdmin':
        result = validateAdmin(args[0], args[1]);
        break;
      case 'getTestResults':
        // args[0]: testName, args[1]: occasion (optional)
        result = getTestResults(args[0], args[1]);
        break;
      case 'getRecentResults': 
        // args[0]: testName, args[1]: filterDept, args[2]: filterArea
        result = getRecentResults(args[0], args[1], args[2]);
        break;
      case 'getQuizData':
        // args[0]: testName, args[1]: context object (dept, area, occasion, startTime)
        result = getQuizData(args[0], args[1]);
        break;
      case 'submitQuiz':
        result = submitQuiz(args[0]);
        break;
      case 'getDiscTestConfig': 
        // args[0]: testName, args[1]: context object
        result = getDiscTestConfig(args[0], args[1]);
        break;
      case 'getDiscQuestions': 
        // args[0]: testName, args[1]: context object
        result = getDiscQuestions(args[0], args[1]);
        break;
      case 'submitDiscTest': 
        result = submitDiscTest(args[0]);
        break;
      case 'getDiscResultsForAdmin':
        result = getDiscResultsForAdmin();
        break;
      case 'getRoleList':
        result = getRoleList();
        break;
      default:
        result = { error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -----------------------------------------------------------------------
// LOGIC NGHIỆP VỤ (BUSINESS LOGIC)
// -----------------------------------------------------------------------

function getTestsList() {
  const combinedList = [];
  const now = new Date();
  const timeZone = Session.getScriptTimeZone();

  // 1. FETCH QUIZ TESTS (Tests-Management)
  const ss = getSS();
  const sheet = ss.getSheetByName(DB_SHEETS.TESTS);
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    // Headers detection
    const colMap = { name: 1, limit: 2, duration: 3, status: 4, endTime: -1, startTime: -1, occasion: -1, saveAnswer: -1, department: -1, forArea: -1 };
    
    const headers = data[0];
    headers.forEach((h, i) => {
      const header = String(h).toLowerCase().trim();
      if (header.includes('end time') || header.includes('hạn nộp')) colMap.endTime = i;
      if (header.includes('start time') || header.includes('bắt đầu')) colMap.startTime = i;
      if (header.includes('save answer') || header.includes('lưu câu trả lời')) colMap.saveAnswer = i;
      if (header.includes('occassion') || header.includes('occasion') || header.includes('dịp')) colMap.occasion = i;
      // Updated to detect 'department test' or 'for department'
      if (header.includes('department test') || header.includes('for department') || header === 'department' || header.includes('phòng ban') || header.includes('bộ phận')) colMap.department = i;
      if (header.includes('for area') || header.includes('khu vực')) colMap.forArea = i;
    });

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[colMap.name]) continue;
      let status = String(row[colMap.status] || "").toLowerCase();

      // Time Logic
      if (status === 'pending' && colMap.startTime > -1) {
        const startVal = row[colMap.startTime];
        if (startVal instanceof Date && now >= startVal) {
          sheet.getRange(i + 1, colMap.status + 1).setValue('Active');
          status = 'active';
        }
      }
      if (status === 'active' && colMap.endTime > -1) {
        const endVal = row[colMap.endTime];
        if (endVal instanceof Date && now > endVal) {
          sheet.getRange(i + 1, colMap.status + 1).setValue('InActive');
          status = 'inactive';
        }
      }

      // Format Start Time for Display
      let startTimeStr = "";
      if (colMap.startTime > -1 && row[colMap.startTime] instanceof Date) {
        startTimeStr = Utilities.formatDate(row[colMap.startTime], timeZone, "dd/MM/yyyy HH:mm");
      }

      combinedList.push({
        type: 'QUIZ',
        name: row[colMap.name],
        duration: row[colMap.duration],
        questionCount: row[colMap.limit],
        status: status,
        occasion: (colMap.occasion > -1) ? row[colMap.occasion] : "",
        department: (colMap.department > -1) ? String(row[colMap.department]).trim() : "",
        forArea: (colMap.forArea > -1) ? String(row[colMap.forArea]).trim() : "",
        saveAnswer: (colMap.saveAnswer > -1) ? row[colMap.saveAnswer] : "",
        startTime: startTimeStr
      });
    }
  }

  // 2. FETCH DISC TESTS
  try {
    const discSS = getDiscSS();
    const discSheet = discSS.getSheetByName(DB_SHEETS.DISC_TESTS);
    if (discSheet) {
      const data = discSheet.getDataRange().getValues();
      const headers = data[0];
      const colMap = { name: -1, count: -1, duration: -1, blockDuration: -1, status: -1, occasion: -1, saveAnswer: -1, start: -1, end: -1, department: -1, forArea: -1 };
      
      headers.forEach((h, i) => {
        const header = String(h).toLowerCase().trim();
        if (header === 'test name') colMap.name = i;
        if (header === 'number of questions') colMap.count = i;
        if (header === 'duration test') colMap.duration = i;
        if (header === 'duration each block question') colMap.blockDuration = i;
        if (header === 'status') colMap.status = i;
        if (header === 'occassion test') colMap.occasion = i;
        if (header === 'save answer sheet') colMap.saveAnswer = i;
        if (header === 'start time') colMap.start = i;
        if (header === 'end time') colMap.end = i;
        // Updated to detect 'department test' or 'for department'
        if (header.includes('department test') || header.includes('for department') || header === 'department' || header.includes('phòng ban')) colMap.department = i;
        if (header.includes('for area') || header.includes('khu vực')) colMap.forArea = i;
      });

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[colMap.name] || colMap.name === -1) continue;
        let status = String(row[colMap.status] || "").toLowerCase();

        // DISC Time Logic
        if (status === 'pending' && colMap.start > -1) {
          const startVal = row[colMap.start];
          if (startVal instanceof Date && now >= startVal) {
            discSheet.getRange(i + 1, colMap.status + 1).setValue('Active');
            status = 'active';
          }
        }
        if (status === 'active' && colMap.end > -1) {
          const endVal = row[colMap.end];
          if (endVal instanceof Date && now > endVal) {
            discSheet.getRange(i + 1, colMap.status + 1).setValue('InActive');
            status = 'inactive';
          }
        }

        // Format Start Time for Display
        let startTimeStr = "";
        if (colMap.start > -1 && row[colMap.start] instanceof Date) {
          startTimeStr = Utilities.formatDate(row[colMap.start], timeZone, "dd/MM/yyyy HH:mm");
        }

        combinedList.push({
          type: 'DISC',
          name: row[colMap.name],
          duration: row[colMap.duration],
          questionCount: row[colMap.count],
          status: status,
          occasion: row[colMap.occasion],
          department: (colMap.department > -1) ? String(row[colMap.department]).trim() : "",
          forArea: (colMap.forArea > -1) ? String(row[colMap.forArea]).trim() : "",
          blockDuration: row[colMap.blockDuration],
          saveAnswer: row[colMap.saveAnswer],
          startTime: startTimeStr
        });
      }
    }
  } catch(e) { console.error("Error fetching DISC: " + e.toString()); }

  return combinedList.reverse();
}

function validateAdmin(username, password) {
  let ss;
  try { ss = SpreadsheetApp.openById(ACC_SPREADSHEET_ID); } catch(e) { return { valid: false }; }
  const sheet = ss.getSheetByName(DB_SHEETS.ACC);
  if (!sheet) return { valid: false };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const inputUser = String(username || "").trim().toLowerCase();
  
  // Dynamic Header Mapping to ensure Role and Area are found correctly
  const colMap = { 
    user: 1, pass: 2, role: -1, dept: 4, quiz: 5, disc: 6, area: -1 
  };

  headers.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    if (header === 'username' || header === 'user' || header === 'tài khoản') colMap.user = i;
    if (header === 'password' || header === 'pass' || header === 'mật khẩu') colMap.pass = i;
    if (header === 'role' || header === 'chức vụ' || header === 'quyền hạn') colMap.role = i;
    if (header === 'department' || header.includes('phòng ban')) colMap.dept = i;
    if (header.includes('quiz permission') || header === 'quiz') colMap.quiz = i;
    if (header.includes('disc permission') || header === 'disc') colMap.disc = i;
    if (header === 'area' || header.includes('khu vực') || header === 'region') colMap.area = i;
  });

  for (let i = 1; i < data.length; i++) {
    const storedUser = String(data[i][colMap.user] || "").toLowerCase();
    const storedPass = String(data[i][colMap.pass] || "");
    
    if (storedUser === inputUser && storedPass === password) {
      const department = (colMap.dept > -1) ? String(data[i][colMap.dept] || "").trim() : "";
      const role = (colMap.role > -1) ? String(data[i][colMap.role] || "").trim() : "";
      const area = (colMap.area > -1) ? String(data[i][colMap.area] || "").trim() : "";
      
      const quizPerm = (colMap.quiz > -1) ? String(data[i][colMap.quiz] || "").toUpperCase().trim() === 'X' : false;
      const discPerm = (colMap.disc > -1) ? String(data[i][colMap.disc] || "").toUpperCase().trim() === 'X' : false;
      
      return { 
        valid: true, 
        permissions: { quiz: quizPerm, disc: discPerm },
        department: department,
        role: role,
        area: area
      };
    }
  }
  return { valid: false };
}

function getTestResults(testName, occasion) {
  const ss = getSS(); // Choice-Management for Metadata
  const historySS = getHistorySS(); // History SS for Results
  
  const resultSheet = historySS.getSheetByName(DB_SHEETS.RESULTS);
  if (!resultSheet) return { passScore: 0, totalQuestions: 0, duration: 0, results: [] };
  
  // Metadata fetch (from Choice-Management)
  const testSheet = ss.getSheetByName(DB_SHEETS.TESTS);
  const testData = testSheet.getDataRange().getValues();
  let passScore = 0, totalQuestions = 0, duration = 0;
  
  // Basic lookup for metadata
  for(let i=1; i<testData.length; i++) {
    if(testData[i][1] === testName) {
       // Assuming standard column structure or fallback
       duration = Number(testData[i][3]); 
       totalQuestions = Number(testData[i][2]); // Default Limit col
       // Attempt to find Score Pass
       const headers = testData[0];
       const spIdx = headers.indexOf("Score Pass");
       if(spIdx > -1) passScore = Number(testData[i][spIdx]);
       break;
    }
  }

  const data = resultSheet.getDataRange().getValues();
  data.shift(); 
  
  const results = data
    .filter(row => {
        // Filter by Test Name
        if (row[2] !== testName) return false;
        // Filter by Occasion if provided (Index 9)
        if (occasion && String(row[9]) !== String(occasion)) return false;
        return true;
    })
    .map(row => ({
      timestamp: row[1],
      timestampStr: formatDate(row[1]),
      candidate: row[3],
      position: row[4],
      timeSaving: Number(row[6]),
      score: Number(row[7]),
      hintsUsed: Number(row[8]) || 0,
      occasion: row[9],
      department: String(row[10] || ""),
      forArea: String(row[11] || "")
    }));
    
  return { passScore, totalQuestions, duration, results };
}

function getRecentResults(testName, filterDept, filterArea) {
  const ss = getHistorySS(); // Use History SS
  const sheet = ss.getSheetByName(DB_SHEETS.RESULTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  
  const fDept = String(filterDept || "").trim().toLowerCase();
  const fArea = String(filterArea || "").trim().toLowerCase();

  // Filter based on TestName AND (Optional) Department AND (Optional) Area
  const filtered = data.filter(row => {
     // Row 2 is TestName
     if (row[2] !== testName) return false;
     
     // Row 10 is Department (Index 10)
     if (fDept) {
       const rDept = String(row[10] || "").trim().toLowerCase();
       if (rDept !== fDept) return false;
     }
     
     // Row 11 is Area (Index 11)
     if (fArea) {
       const rArea = String(row[11] || "").trim().toLowerCase();
       if (rArea !== fArea) return false;
     }
     
     return true;
  });

  return filtered.slice(-15).reverse().map(row => ({
      candidate: row[3],
      score: Number(row[7]).toFixed(2),
      timestamp: formatDate(row[1])
  }));
}

function formatDate(dateObj) {
  try { return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss"); } 
  catch (e) { return String(dateObj); }
}

function getQuizData(testName, context) {
  const ss = getSS();
  const metaSheet = ss.getSheetByName(DB_SHEETS.TESTS);
  const metaData = metaSheet.getDataRange().getValues();
  let testConfig = null;
  
  // Find Test Config with Context Matching
  const headers = metaData[0];
  let deptIdx = -1, areaIdx = -1, occasionIdx = -1, startTimeIdx = -1;
  
  headers.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    // Updated to detect 'department test' or 'for department'
    if (header.includes('department test') || header.includes('for department') || header === 'department' || header.includes('phòng ban')) deptIdx = i;
    if (header.includes('for area') || header.includes('khu vực')) areaIdx = i;
    if (header.includes('occassion') || header.includes('occasion') || header.includes('dịp')) occasionIdx = i;
    if (header.includes('start time') || header.includes('bắt đầu')) startTimeIdx = i;
  });

  for (let i = 1; i < metaData.length; i++) {
    // 1. Match Name
    if (metaData[i][1] !== testName) continue;
    
    // 2. Strict Context Matching if context is provided
    if (context) {
        if (context.department && deptIdx > -1 && String(metaData[i][deptIdx]) !== context.department) continue;
        if (context.forArea && areaIdx > -1 && String(metaData[i][areaIdx]) !== context.forArea) continue;
        if (context.occasion && occasionIdx > -1 && String(metaData[i][occasionIdx]) !== context.occasion) continue;
        // StartTime match is tricky due to date formats, but we can check if passed
        // For simplicity, Department + Area + Occasion + Name is usually unique enough.
    }

    testConfig = {
      name: metaData[i][1],
      limit: Number(metaData[i][2]), // Default col C
      duration: metaData[i][3],
      passScore: 0,
      department: (deptIdx > -1) ? metaData[i][deptIdx] : "",
      forArea: (areaIdx > -1) ? metaData[i][areaIdx] : "",
      occasion: (occasionIdx > -1) ? metaData[i][occasionIdx] : "",
      startTime: (startTimeIdx > -1 && metaData[i][startTimeIdx] instanceof Date) 
                 ? Utilities.formatDate(metaData[i][startTimeIdx], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : ""
    };
    
    // Try to find Pass Score
    const spIdx = headers.indexOf("Score Pass");
    if(spIdx > -1) testConfig.passScore = Number(metaData[i][spIdx]);
    
    break; // Found the specific test
  }
  
  if (!testConfig) throw new Error("Không tìm thấy cấu hình bài thi này (với thông tin chi tiết đã chọn).");

  const bankSheet = ss.getSheetByName(DB_SHEETS.BANK);
  const bankData = bankSheet.getDataRange().getValues();
  
  // Scan headers to find 'Section' column
  const bankHeaders = bankData[0];
  let sectionIdx = -1;
  bankHeaders.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    if(header === 'section' || header === 'phân vùng' || header.includes('kiến thức')) sectionIdx = i;
  });

  bankData.shift(); 
  
  let questions = bankData
    .filter(row => row[1] === testName)
    .map(row => ({
      question: row[2],
      type: row[3],
      choices: [row[4], row[5], row[6], row[7]].filter(c => c !== "" && c !== undefined), 
      hint: row[8],
      section: (sectionIdx > -1) ? String(row[sectionIdx]) : ""
    }));
    
  questions = shuffleArray(questions).slice(0, testConfig.limit);
  return { config: testConfig, questions: questions };
}

function normalizeStr(str) { return String(str || "").trim().toLowerCase(); }

function submitQuiz(payload) {
  // Use LockService to prevent race conditions during ID generation and write
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait up to 30s for the lock
  
  try {
      const ss = getSS(); // Use Choice-Management for Bank & Tests
      const bankSheet = ss.getSheetByName(DB_SHEETS.BANK);
      const bankData = bankSheet.getDataRange().getValues();
      const historySS = getHistorySS();

      // Detect Section Header in Bank
      const bankHeaders = bankData[0];
      let sectionIdx = -1;
      bankHeaders.forEach((h, i) => {
        const header = String(h).toLowerCase().trim();
        if(header === 'section' || header === 'phân vùng' || header.includes('kiến thức')) sectionIdx = i;
      });
      
      let correctCount = 0;
      const resultDetails = []; 
      const keyLabels = ['A', 'B', 'C', 'D'];

      payload.answers.forEach(ans => {
        const qKey = normalizeStr(ans.question);
        const targetRow = bankData.find(row => String(row[1]) === payload.testName && normalizeStr(row[2]) === qKey);
        
        // Extract Section info
        const qSection = (sectionIdx > -1 && targetRow) ? String(targetRow[sectionIdx] || "") : "";

        if (!targetRow) {
          resultDetails.push({ question: ans.question, isCorrect: false, userSelected: ans.selected, correctTexts: ["Lỗi"], hint: "", section: "" });
          return;
        }

        const optionTexts = [targetRow[4], targetRow[5], targetRow[6], targetRow[7]]; 
        const hint = String(targetRow[8] || ""); 
        const correctRaw = String(targetRow[9]).trim(); 
        let correctKeys = correctRaw.toUpperCase().split(/[^A-D]+/).filter(k => k).sort();
        
        let userSelectedArr = Array.isArray(ans.selected) ? ans.selected : [ans.selected];
        let userKeys = [];
        userSelectedArr.forEach(userTxt => {
          const idx = optionTexts.findIndex(opt => normalizeStr(opt) === normalizeStr(userTxt));
          if (idx !== -1) userKeys.push(keyLabels[idx]);
        });
        userKeys.sort(); 

        const isCorrect = (userKeys.join(" ") === correctKeys.join(" ")) && (userKeys.length > 0);
        if (isCorrect) correctCount++;

        const correctTexts = correctKeys.map(k => String(optionTexts[keyLabels.indexOf(k)] || ""));

        resultDetails.push({ question: targetRow[2], isCorrect: isCorrect, userSelected: userSelectedArr, correctTexts: correctTexts, hint: hint, section: qSection });
      });
      
      // Calculate Bonus & Fetch Occasion & Check Save Answer
      const metaSheet = ss.getSheetByName(DB_SHEETS.TESTS);
      const metaData = metaSheet.getDataRange().getValues();
      let numQuestions = 0, duration = 1;
      let occasion = payload.occasion || ""; 
      let saveAnswerConfig = "";
      let department = payload.department || "";
      let forArea = payload.forArea || "";

      // Dynamic header check
      const headers = metaData[0];
      let occasionIdx = -1;
      let saveAnsIdx = -1;
      let deptIdx = -1;
      let areaIdx = -1;
      let startTimeIdx = -1;

      headers.forEach((h, i) => {
        const header = String(h).toLowerCase().trim();
        if (header.includes('occassion') || header.includes('occasion') || header.includes('dịp')) occasionIdx = i;
        if (header.includes('save answer') || header.includes('lưu câu trả lời')) saveAnsIdx = i;
        // Updated to detect 'department test' or 'for department'
        if (header.includes('department test') || header.includes('for department') || header === 'department' || header.includes('phòng ban') || header.includes('bộ phận')) deptIdx = i;
        if (header.includes('for area') || header.includes('khu vực')) areaIdx = i;
        if (header.includes('start time') || header.includes('bắt đầu')) startTimeIdx = i;
      });

      // STRICT LOOKUP LOOP
      // We want to find the config row that matches NOT JUST the name, but ALL the context fields provided.
      for (let i = 1; i < metaData.length; i++) {
        const rowName = metaData[i][1];
        if (rowName !== payload.testName) continue;

        // Check strict context provided by payload
        if (payload.department && deptIdx > -1 && String(metaData[i][deptIdx]) !== payload.department) continue;
        if (payload.forArea && areaIdx > -1 && String(metaData[i][areaIdx]) !== payload.forArea) continue;
        if (payload.occasion && occasionIdx > -1 && String(metaData[i][occasionIdx]) !== payload.occasion) continue;
        
        // Check start time if provided (formatted string comparison)
        if (payload.startTime && startTimeIdx > -1 && metaData[i][startTimeIdx] instanceof Date) {
            const rowTimeStr = Utilities.formatDate(metaData[i][startTimeIdx], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
            if (rowTimeStr !== payload.startTime) continue;
        }

        // Match Found!
        numQuestions = Number(metaData[i][2]); 
        duration = Number(metaData[i][3]);     
        
        // Use the values from the sheet to be safe, or fallback to payload
        if (occasionIdx > -1) occasion = metaData[i][occasionIdx];
        if (saveAnsIdx > -1) saveAnswerConfig = String(metaData[i][saveAnsIdx]);
        if (deptIdx > -1) department = String(metaData[i][deptIdx]);
        if (areaIdx > -1) forArea = String(metaData[i][areaIdx]);
        break;
      }

      const timeSaving = parseFloat(payload.timeSaving) || 0;
      let bonus = (duration > 0 && numQuestions > 0) ? 0.02 * (numQuestions * timeSaving / duration) : 0;
      const totalResultStr = (correctCount + bonus).toFixed(2);
      
      // Ensure numeric types for saving
      const timeSavingVal = parseFloat(timeSaving.toFixed(2));
      const scoreVal = parseFloat(totalResultStr);
      const hintsUsedVal = Number(payload.hintsUsed || 0);

      // Save to History Spreadsheet (Summary)
      const resultSheet = historySS.getSheetByName(DB_SHEETS.RESULTS);
      const lastRow = resultSheet.getLastRow();
      let newId = 1;
      if (lastRow > 1) {
        // Read explicitly from Column 1 to avoid confusion with other columns
        const lastIdVal = resultSheet.getRange(lastRow, 1).getValue();
        newId = Number(lastIdVal) + 1;
        if (isNaN(newId)) newId = 1;
      }
      
      const now = new Date();
      resultSheet.appendRow([
        newId, 
        now, 
        payload.testName, 
        payload.candidate, 
        payload.position, 
        correctCount, 
        timeSavingVal, // Save as Number
        scoreVal,      // Save as Number
        hintsUsedVal,  // Save as Number
        occasion, 
        department, 
        forArea
      ]);
      
      // Save Detailed Answers if Configured
      if (saveAnswerConfig.toUpperCase().trim() === 'X') {
        const answerSheet = historySS.getSheetByName(DB_SHEETS.QUIZ_ANSWERS);
        if (answerSheet) {
          const detailRows = [];
          const timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yy HH:mm");
          resultDetails.forEach(d => {
            const userAnsStr = Array.isArray(d.userSelected) ? d.userSelected.join(", ") : d.userSelected;
            // Cols: ID | Timestamp | Occasion | TestName | Name | Position | Question | Section | User Answer | Correct | Note | Department Test | For Area
            detailRows.push([
              newId,
              timestampStr,
              occasion,
              payload.testName,
              payload.candidate,
              payload.position,
              d.question,
              d.section,
              userAnsStr,
              d.isCorrect ? "X" : "",
              "", // Note
              department,
              forArea
            ]);
          });
          if(detailRows.length > 0) {
            answerSheet.getRange(answerSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length).setValues(detailRows);
          }
        }
      }

      return { score: totalResultStr, hintsUsed: hintsUsedVal, details: resultDetails };
  } catch (e) {
      throw e;
  } finally {
      lock.releaseLock();
  }
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

// -----------------------------------------------------------------------
// DISC SPECIFIC LOGIC
// -----------------------------------------------------------------------

function getDiscTestConfig(testName, context) {
  const ss = getDiscSS();
  const sheet = ss.getSheetByName(DB_SHEETS.DISC_TESTS);
  if (!sheet) throw new Error("DISC Database not found");
  const data = sheet.getDataRange().getValues();
  
  // Headers check
  const headers = data[0];
  const colMap = { name: 1, duration: 3, blockDuration: 4, occasion: 6, saveAnswer: 7, department: -1, forArea: -1, start: -1 };
  headers.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    if(header === 'test name') colMap.name = i;
    if(header === 'duration test') colMap.duration = i;
    if(header === 'duration each block question') colMap.blockDuration = i;
    if(header === 'occassion test') colMap.occasion = i;
    if(header === 'save answer sheet') colMap.saveAnswer = i;
    if(header === 'start time') colMap.start = i;
    
    // Updated to detect 'department test' or 'for department'
    if (header.includes('department test') || header.includes('for department') || header === 'department' || header.includes('phòng ban')) colMap.department = i;
    if (header.includes('for area') || header.includes('khu vực')) colMap.forArea = i;
  });

  for (let i = 1; i < data.length; i++) {
    // 1. Match Name
    if (data[i][colMap.name] !== testName) continue;
    
    // 2. Strict Context Matching
    if (context) {
      if (context.department && colMap.department > -1 && String(data[i][colMap.department]) !== context.department) continue;
      if (context.forArea && colMap.forArea > -1 && String(data[i][colMap.forArea]) !== context.forArea) continue;
      if (context.occasion && colMap.occasion > -1 && String(data[i][colMap.occasion]) !== context.occasion) continue;
    }

    let startTimeStr = "";
    if (colMap.start > -1 && data[i][colMap.start] instanceof Date) {
        startTimeStr = Utilities.formatDate(data[i][colMap.start], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    }

    return {
      type: 'DISC',
      name: data[i][colMap.name],
      duration: data[i][colMap.duration],
      blockDuration: data[i][colMap.blockDuration],
      occasion: data[i][colMap.occasion],
      saveAnswer: data[i][colMap.saveAnswer],
      department: (colMap.department > -1) ? data[i][colMap.department] : "",
      forArea: (colMap.forArea > -1) ? data[i][colMap.forArea] : "",
      startTime: startTimeStr
    };
  }
  throw new Error("Không tìm thấy thông tin bài test DISC (với thông tin chi tiết đã chọn).");
}

function getDiscQuestions(testName, context) {
  // === MODIFIED HERE: Use History Spreadsheet for Questions ===
  const ss = getHistorySS(); 
  // Pass context to config getter
  const config = getDiscTestConfig(testName, context);
  const qSheet = ss.getSheetByName(DB_SHEETS.DISC_QUESTIONS);
  if (!qSheet) throw new Error("Sheet DISC-Questions not found");
  
  const qData = qSheet.getDataRange().getValues();
  const headers = qData[0];
  const colMap = { id: 0, question: 1, most: 2, least: 3, fit: -1, nonFit: -1 }; // Default
  
  headers.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    if (header === 'id') colMap.id = i;
    if (header.includes('question') || header.includes('nội dung')) colMap.question = i;
    if (header.includes('most') || header.includes('giống')) colMap.most = i;
    if (header.includes('least') || header.includes('khác')) colMap.least = i;
    if (header.includes('fit culture') || header.includes('văn hóa')) colMap.fit = i;
    if (header.includes('critical non-fit') || header.includes('non-fit')) colMap.nonFit = i;
  });

  const questions = [];
  for(let i=1; i<qData.length; i++) {
    const row = qData[i];
    if(!row[colMap.id]) continue;
    
    // Check flags (X indicates true)
    const isFit = (colMap.fit > -1) && (String(row[colMap.fit] || "").toUpperCase().trim() === 'X');
    const isNonFit = (colMap.nonFit > -1) && (String(row[colMap.nonFit] || "").toUpperCase().trim() === 'X');

    questions.push({
      id: row[colMap.id],
      content: row[colMap.question],
      mostValue: row[colMap.most],
      leastValue: row[colMap.least],
      isFit: isFit,
      isNonFit: isNonFit
    });
  }
  return { config: config, questions: questions };
}

function submitDiscTest(payload) {
  // Use LockService to prevent race conditions during ID generation
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait up to 30s

  try {
      const ss = getHistorySS();
      
      const context = {
        department: payload.department,
        forArea: payload.forArea,
        occasion: payload.occasion,
        startTime: payload.startTime // Optional matching
      };
      
      // Fetch authoritative config to get the Department/Area directly from sheet
      // instead of relying solely on payload, ensuring data integrity.
      let authDept = payload.department;
      let authArea = payload.forArea;
      let saveAnswerConfig = "";

      try {
        const config = getDiscTestConfig(payload.testName, context);
        if (config) {
            authDept = config.department || "";
            authArea = config.forArea || "";
            saveAnswerConfig = String(config.saveAnswer || "");
        }
      } catch (e) {
        console.warn("Could not fetch authoritative DISC config for submission: " + e.toString());
      }
      
      // We use getDiscQuestions just to get the valid questions/IDs for calculation
      const data = getDiscQuestions(payload.testName, context);
      const questions = data.questions;

      // Generate Submission ID
      const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
      let submissionId = "";
      for (let i = 0; i < 2; i++) submissionId += chars.charAt(Math.floor(Math.random() * chars.length));
      
      const now = new Date();
      const timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yy HH:mm");

      // 1. SAVE RAW ANSWERS
      if (saveAnswerConfig.toUpperCase().trim() === 'X') {
        const sheetAns = ss.getSheetByName(DB_SHEETS.DISC_ANSWERS);
        if(sheetAns) {
          const rows = [];
          questions.forEach(q => {
            const qId = String(q.id);
            const sel = payload.selections[qId];
            let mostVal = (sel === 'MOST') ? q.mostValue : "";
            let leastVal = (sel === 'LEAST') ? q.leastValue : "";
            rows.push([
                submissionId, 
                timestampStr, 
                payload.occasion,  // Use payload exact occasion
                payload.candidate, 
                payload.position, 
                q.content, 
                mostVal, 
                leastVal, 
                "", // Note Column
                authDept, // Use Authoritative Dept
                authArea  // Use Authoritative Area
            ]);
          });
          if(rows.length > 0) sheetAns.getRange(sheetAns.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        }
      }

      // 2. SAVE CALCULATED RESULTS
      const sheetRes = ss.getSheetByName(DB_SHEETS.DISC_RESULTS);
      let newResultId = 1;
      if(sheetRes) {
        const lastRow = sheetRes.getLastRow();
        if(lastRow > 1) {
          const lastIdVal = sheetRes.getRange(lastRow, 1).getValue();
          newResultId = Number(lastIdVal) + 1;
          if(isNaN(newResultId)) newResultId = 1;
        }
        
        // Calculate counts for D,I,S,C
        const counts = { MOST: { D:0, I:0, S:0, C:0 }, LEAST: { D:0, I:0, S:0, C:0 } };
        const normalizeKey = (v) => {
          const s = String(v).toUpperCase().trim();
          if(s==='1'||s==='D') return 'D'; if(s==='2'||s==='I') return 'I'; if(s==='3'||s==='S') return 'S'; if(s==='4'||s==='C') return 'C';
          return null;
        };

        if(payload.selections) {
          for(const [qId, type] of Object.entries(payload.selections)) {
            const q = questions.find(item => String(item.id) === String(qId));
            if(q) {
              const val = (type === 'MOST') ? q.mostValue : q.leastValue;
              const key = normalizeKey(val);
              if(key) counts[type][key]++;
            }
          }
        }

        // Extract calculated culture strings
        const cultureAnalysis = payload.cultureAnalysis || {};
        const strMostFit = cultureAnalysis.mostFit || "";
        const strLeastFit = cultureAnalysis.leastFit || "";
        const strMostNonFit = cultureAnalysis.mostNonFit || "";

        sheetRes.appendRow([
          newResultId, 
          timestampStr, 
          payload.occasion, // Use Payload Occasion
          payload.candidate, 
          payload.position,
          payload.currentDisc || "", 
          payload.trendDisc || "",
          counts.LEAST.D, counts.LEAST.I, counts.LEAST.S, counts.LEAST.C,
          counts.MOST.D, counts.MOST.I, counts.MOST.S, counts.MOST.C,
          strMostFit,      // Most-Fit Culture (Col 16)
          strLeastFit,     // Least-Fit Cuture (Col 17)
          strMostNonFit,   // Most-NonFit Culture (Col 18)
          "",              // Note (Col 19)
          authDept,        // Use Authoritative Dept from Config
          authArea         // Use Authoritative Area from Config
        ]);
      }

      return { status: 'success', timestamp: timestampStr };
  } catch (e) {
      throw e;
  } finally {
      lock.releaseLock();
  }
}

function getDiscResultsForAdmin() {
  const ss = getHistorySS();
  const sheet = ss.getSheetByName(DB_SHEETS.DISC_RESULTS);
  if (!sheet) return { results: [], meta: { totalFit: 0, totalNonFit: 0 } };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Dynamic Header Mapping for Robustness (specifically handling 'For Department')
  let deptIdx = 19; // Default fallback
  let areaIdx = 20; // Default fallback
  
  headers.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    if (header === 'for department' || header.includes('department') || header.includes('phòng ban')) deptIdx = i;
    if (header.includes('for area') || header.includes('khu vực')) areaIdx = i;
  });

  data.shift(); // Remove header
  
  const results = data.map(row => ({
    id: row[0], timestamp: row[1], occasion: row[2], name: row[3], position: row[4],
    currentDiscStr: row[5], trendDiscStr: row[6],
    least: { D: row[7], I: row[8], S: row[9], C: row[10] },
    most: { D: row[11], I: row[12], S: row[13], C: row[14] },
    strMostFit: row[15] || "",
    strLeastFit: row[16] || "",
    strMostNonFit: row[17] || "",
    department: row[deptIdx], // Use Dynamic Index
    forArea: row[areaIdx],    // Use Dynamic Index
    fitPoint: 0, 
    violatePoint: 0
  })).reverse();

  // CALCULATE META DATA (Total Fit / NonFit counts from DB)
  let totalFit = 0;
  let totalNonFit = 0;
  try {
    // === MODIFIED HERE: Use History Spreadsheet for Questions ===
    const discSS = getHistorySS(); 
    const qSheet = discSS.getSheetByName(DB_SHEETS.DISC_QUESTIONS);
    if (qSheet) {
      const qData = qSheet.getDataRange().getValues();
      const headers = qData[0];
      let fIdx = -1, nfIdx = -1;
      // Dynamic header find
      headers.forEach((h, i) => {
         const s = String(h).toLowerCase().trim();
         if (s.includes('fit culture') || s.includes('văn hóa')) fIdx = i;
         if (s.includes('critical non-fit') || s.includes('non-fit')) nfIdx = i;
      });
      
      for(let i=1; i<qData.length; i++) {
         if (fIdx > -1 && String(qData[i][fIdx]).toUpperCase().trim() === 'X') totalFit++;
         if (nfIdx > -1 && String(qData[i][nfIdx]).toUpperCase().trim() === 'X') totalNonFit++;
      }
    }
  } catch(e) { console.error("Error counting culture totals: " + e.toString()); }

  return { results: results, meta: { totalFit: totalFit, totalNonFit: totalNonFit } };
}

function getRoleList() {
  let ss;
  try { ss = SpreadsheetApp.openById(ACC_SPREADSHEET_ID); } catch(e) { return []; }
  const sheet = ss.getSheetByName(DB_SHEETS.ROLES);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  // Assume Row 1 is header: ID | Role | Note
  const roles = [];
  // Start from row index 1 (second row)
  for (let i = 1; i < data.length; i++) {
    const role = String(data[i][1] || "").trim(); // Column B is index 1
    if (role) roles.push(role);
  }
  return roles;
}