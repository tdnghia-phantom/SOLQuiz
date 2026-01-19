
// -----------------------------------------------------------------------
// CẤU HÌNH & THIẾT LẬP (CONFIGURATION)
// -----------------------------------------------------------------------

// ID của file Google Sheet "Choice-Management"
const SPREADSHEET_ID = '1xalAKVKbLBUfjv8uKuYluyqpizl2fU9b9WfmHzj4HWQ';

// Tên các Sheet
const DB_SHEETS = {
  ACC: 'Acc-Management',        
  TESTS: 'Tests-Management',    
  BANK: 'Test-Bank',            
  RESULTS: 'Result-Test'        
};

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// VIEW: Trả về HTML khi truy cập trực tiếp bằng trình duyệt
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Hệ Thống Thi Trắc Nghiệm')
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
        result = getTestResults(args[0]);
        break;
      case 'getRecentResults': // New Action
        result = getRecentResults(args[0]);
        break;
      case 'getQuizData':
        result = getQuizData(args[0]);
        break;
      case 'submitQuiz':
        result = submitQuiz(args[0]);
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
  const ss = getSS();
  const sheet = ss.getSheetByName(DB_SHEETS.TESTS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  data.shift(); // Bỏ dòng tiêu đề
  
  // Col 4: Status (Active)
  const activeTests = data
    .filter(row => String(row[4]).toLowerCase() === 'active') 
    .map(row => ({
      name: row[1],      
      duration: row[3]   
    }))
    .reverse(); 
    
  return activeTests;
}

function validateAdmin(username, password) {
  const ss = getSS();
  const sheet = ss.getSheetByName(DB_SHEETS.ACC);
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === username && String(data[i][2]) === password) {
      return true;
    }
  }
  return false;
}

function getTestResults(testName) {
  const ss = getSS();
  
  // 1. Get Pass Score & Total Questions from Tests-Management
  const testSheet = ss.getSheetByName(DB_SHEETS.TESTS);
  const testData = testSheet.getDataRange().getValues();
  const headers = testData[0];
  const scorePassIdx = headers.indexOf("Score Pass");
  
  // Find column for Total Questions dynamically
  // Headers: ID, Test Name, Limit/TotalQuestions, Duration, Status, ...
  let limitIdx = -1;
  ["Limit", "Number of Questions", "Số câu"].forEach(key => {
     if (limitIdx === -1) limitIdx = headers.indexOf(key);
  });
  if (limitIdx === -1) limitIdx = 2; // Fallback to column index 2 (C)

  let passScore = 0;
  let totalQuestions = 0;
  let duration = 0;
  
  for (let i = 1; i < testData.length; i++) {
    if (testData[i][1] === testName) {
      passScore = (scorePassIdx > -1) ? Number(testData[i][scorePassIdx]) : 0;
      totalQuestions = Number(testData[i][limitIdx]);
      duration = Number(testData[i][3]); // Get Duration
      break;
    }
  }

  // 2. Get Results
  const resultSheet = ss.getSheetByName(DB_SHEETS.RESULTS);
  if (!resultSheet) return { passScore: passScore, totalQuestions: totalQuestions, duration: duration, results: [] };
  
  const data = resultSheet.getDataRange().getValues();
  data.shift(); 
  
  const results = data
    .filter(row => row[2] === testName)
    .map(row => ({
      timestamp: row[1], // Keep raw object for sorting
      timestampStr: formatDate(row[1]),
      candidate: row[3],
      position: row[4],
      timeSaving: Number(row[6]), // Col 6 is TimeSaving
      score: Number(row[7]),      // Col 7 is Score
      hintsUsed: Number(row[8]) || 0 // Col 8 (9th column) is Hints Used
    }));
    
  // Return metadata and list
  return {
    passScore: passScore,
    totalQuestions: totalQuestions,
    duration: duration, // Return duration
    results: results
  };
}

// NEW FUNCTION: Get recent results for live feed
function getRecentResults(testName) {
  const ss = getSS();
  const sheet = ss.getSheetByName(DB_SHEETS.RESULTS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  
  // Filter by test name and take the last 15 entries (most recent)
  // Assuming data is appended, the last rows are the newest.
  const recent = data
    .filter(row => row[2] === testName)
    .slice(-15) // Take last 15
    .reverse() // Newest first
    .map(row => ({
      candidate: row[3],
      score: Number(row[7]).toFixed(2),
      timestamp: formatDate(row[1]) // Optional for display
    }));

  return recent;
}

function formatDate(dateObj) {
  try {
    return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  } catch (e) {
    return String(dateObj);
  }
}

function getQuizData(testName) {
  const ss = getSS();
  const metaSheet = ss.getSheetByName(DB_SHEETS.TESTS);
  const metaData = metaSheet.getDataRange().getValues();
  const headers = metaData[0];
  
  const scorePassIdx = headers.indexOf("Score Pass");
  
  let limitIdx = -1;
  ["Limit", "Number of Questions", "Số câu"].forEach(key => {
     if (limitIdx === -1) limitIdx = headers.indexOf(key);
  });
  if (limitIdx === -1) limitIdx = 2;

  let testConfig = null;
  
  for (let i = 1; i < metaData.length; i++) {
    if (metaData[i][1] === testName) {
      testConfig = {
        name: metaData[i][1],
        limit: Number(metaData[i][limitIdx]),
        duration: metaData[i][3],
        passScore: (scorePassIdx > -1) ? Number(metaData[i][scorePassIdx]) : 0
      };
      break;
    }
  }
  
  if (!testConfig) throw new Error("Không tìm thấy cấu hình bài thi này.");

  const bankSheet = ss.getSheetByName(DB_SHEETS.BANK);
  const bankData = bankSheet.getDataRange().getValues();
  bankData.shift(); 
  
  let questions = bankData
    .filter(row => row[1] === testName)
    .map(row => ({
      question: row[2],
      type: row[3],
      choices: [row[4], row[5], row[6], row[7]].filter(c => c !== "" && c !== undefined), 
      hint: row[8]
    }));
    
  questions = shuffleArray(questions).slice(0, testConfig.limit);
  
  return {
    config: testConfig,
    questions: questions
  };
}

// Hàm chuẩn hóa chuỗi để so sánh chính xác hơn
function normalizeStr(str) {
  return String(str || "").trim().toLowerCase();
}

function submitQuiz(payload) {
  const ss = getSS();
  const bankSheet = ss.getSheetByName(DB_SHEETS.BANK);
  const bankData = bankSheet.getDataRange().getValues();
  
  let correctCount = 0;
  const resultDetails = []; 

  const keyLabels = ['A', 'B', 'C', 'D'];

  payload.answers.forEach(ans => {
    const qKey = normalizeStr(ans.question);
    
    const targetRow = bankData.find(row => 
      String(row[1]) === payload.testName && 
      normalizeStr(row[2]) === qKey
    );
    
    if (!targetRow) {
      resultDetails.push({
        question: ans.question,
        isCorrect: false,
        userSelected: ans.selected,
        correctTexts: ["Lỗi"],
        hint: "Vui lòng kiểm tra lại dữ liệu"
      });
      return;
    }

    const optionTexts = [targetRow[4], targetRow[5], targetRow[6], targetRow[7]]; 
    const hint = String(targetRow[8] || ""); 
    const correctRaw = String(targetRow[9]).trim(); 

    let correctKeys = correctRaw.toUpperCase().split(/[^A-D]+/).filter(k => k);
    correctKeys.sort(); 

    let userSelectedArr = Array.isArray(ans.selected) ? ans.selected : [ans.selected];
    let userKeys = [];

    userSelectedArr.forEach(userTxt => {
      const idx = optionTexts.findIndex(opt => normalizeStr(opt) === normalizeStr(userTxt));
      if (idx !== -1) {
        userKeys.push(keyLabels[idx]);
      }
    });
    userKeys.sort(); 

    const correctStr = correctKeys.join(" ");
    const userStr = userKeys.join(" ");

    const isCorrect = (userStr === correctStr) && (userStr.length > 0);

    if (isCorrect) {
      correctCount++;
    }

    const correctTexts = correctKeys.map(k => {
      const idx = keyLabels.indexOf(k);
      return String(optionTexts[idx] || ""); 
    });

    resultDetails.push({
      question: targetRow[2], 
      isCorrect: isCorrect,
      userSelected: userSelectedArr, 
      correctTexts: correctTexts,    
      hint: hint
    });
  });
  
  // --- Tính Điểm ---
  const metaSheet = ss.getSheetByName(DB_SHEETS.TESTS);
  const metaData = metaSheet.getDataRange().getValues();
  let numQuestions = 0;
  let duration = 1;
  
  // Find column for Total Questions dynamically
  const headers = metaData[0];
  let limitIdx = -1;
  ["Limit", "Number of Questions", "Số câu"].forEach(key => {
     if (limitIdx === -1) limitIdx = headers.indexOf(key);
  });
  if (limitIdx === -1) limitIdx = 2;

  for (let i = 1; i < metaData.length; i++) {
    if (metaData[i][1] === payload.testName) {
      numQuestions = Number(metaData[i][limitIdx]); 
      duration = Number(metaData[i][3]);     
      break;
    }
  }

  const timeSaving = parseFloat(payload.timeSaving) || 0;
  let bonus = 0;
  if (duration > 0 && numQuestions > 0) {
    bonus = 0.02 * (numQuestions * timeSaving / duration);
  }
  
  const totalResult = correctCount + bonus;
  const scoreStr = totalResult.toFixed(2);
  const hintsUsed = payload.hintsUsed || 0; // Get hints used from payload
  
  // Lưu kết quả vào Sheet
  const resultSheet = ss.getSheetByName(DB_SHEETS.RESULTS);
  const lastRow = resultSheet.getLastRow();
  let newId = 1;
  if (lastRow > 1) {
    const lastId = resultSheet.getRange(lastRow, 1).getValue();
    newId = Number(lastId) + 1;
  }
  if (isNaN(newId)) newId = 1;
  
  resultSheet.appendRow([
    newId,
    new Date(),
    payload.testName,
    payload.candidate,
    payload.position,
    correctCount,
    timeSaving.toFixed(2),
    scoreStr,
    hintsUsed // Col 9: Hints Used
  ]);
  
  return {
    score: scoreStr,
    hintsUsed: hintsUsed, // Return hints used so frontend can display it
    details: resultDetails
  };
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}
