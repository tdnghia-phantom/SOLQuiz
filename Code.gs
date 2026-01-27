
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
  RESULTS: 'Result-Test',
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
        result = getTestResults(args[0]);
        break;
      case 'getRecentResults': 
        result = getRecentResults(args[0]);
        break;
      case 'getQuizData':
        result = getQuizData(args[0]);
        break;
      case 'submitQuiz':
        result = submitQuiz(args[0]);
        break;
      case 'getDiscTestConfig': 
        result = getDiscTestConfig(args[0]);
        break;
      case 'getDiscQuestions': 
        result = getDiscQuestions(args[0]);
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

  // 1. FETCH QUIZ TESTS
  const ss = getSS();
  const sheet = ss.getSheetByName(DB_SHEETS.TESTS);
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    // Headers: ID, Test Name, Limit, Duration, Status, End Time, Start Time...
    // Fallback mapping
    const colMap = { name: 1, limit: 2, duration: 3, status: 4, endTime: -1, startTime: -1 };
    
    // Simple dynamic header check
    const headers = data[0];
    headers.forEach((h, i) => {
      const header = String(h).toLowerCase().trim();
      if (header.includes('end time') || header.includes('hạn nộp')) colMap.endTime = i;
      if (header.includes('start time') || header.includes('bắt đầu')) colMap.startTime = i;
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

      combinedList.push({
        type: 'QUIZ',
        name: row[colMap.name],
        duration: row[colMap.duration],
        questionCount: row[colMap.limit],
        status: status
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
      const colMap = { name: -1, count: -1, duration: -1, blockDuration: -1, status: -1, occasion: -1, saveAnswer: -1, start: -1, end: -1 };
      
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

        combinedList.push({
          type: 'DISC',
          name: row[colMap.name],
          duration: row[colMap.duration],
          questionCount: row[colMap.count],
          status: status,
          occasion: row[colMap.occasion],
          blockDuration: row[colMap.blockDuration],
          saveAnswer: row[colMap.saveAnswer]
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
  const inputUser = String(username || "").trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const storedUser = String(data[i][1]).toLowerCase();
    const storedPass = String(data[i][2]);
    if (storedUser === inputUser && storedPass === password) {
      const quizPerm = String(data[i][4] || "").toUpperCase().trim() === 'X';
      const discPerm = String(data[i][5] || "").toUpperCase().trim() === 'X';
      return { valid: true, permissions: { quiz: quizPerm, disc: discPerm } };
    }
  }
  return { valid: false };
}

function getTestResults(testName) {
  const ss = getSS();
  const resultSheet = ss.getSheetByName(DB_SHEETS.RESULTS);
  if (!resultSheet) return { passScore: 0, totalQuestions: 0, duration: 0, results: [] };
  
  // Metadata fetch
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
    .filter(row => row[2] === testName)
    .map(row => ({
      timestamp: row[1],
      timestampStr: formatDate(row[1]),
      candidate: row[3],
      position: row[4],
      timeSaving: Number(row[6]),
      score: Number(row[7]),
      hintsUsed: Number(row[8]) || 0
    }));
    
  return { passScore, totalQuestions, duration, results };
}

function getRecentResults(testName) {
  const ss = getSS();
  const sheet = ss.getSheetByName(DB_SHEETS.RESULTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.filter(row => row[2] === testName).slice(-15).reverse().map(row => ({
      candidate: row[3],
      score: Number(row[7]).toFixed(2),
      timestamp: formatDate(row[1])
  }));
}

function formatDate(dateObj) {
  try { return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), "dd/MM/yy HH:mm:ss"); } 
  catch (e) { return String(dateObj); }
}

function getQuizData(testName) {
  const ss = getSS();
  const metaSheet = ss.getSheetByName(DB_SHEETS.TESTS);
  const metaData = metaSheet.getDataRange().getValues();
  let testConfig = null;
  
  // Find Test Config
  for (let i = 1; i < metaData.length; i++) {
    if (metaData[i][1] === testName) {
      testConfig = {
        name: metaData[i][1],
        limit: Number(metaData[i][2]), // Default col C
        duration: metaData[i][3],
        passScore: 0
      };
      // Try to find Pass Score
      const headers = metaData[0];
      const spIdx = headers.indexOf("Score Pass");
      if(spIdx > -1) testConfig.passScore = Number(metaData[i][spIdx]);
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
  return { config: testConfig, questions: questions };
}

function normalizeStr(str) { return String(str || "").trim().toLowerCase(); }

function submitQuiz(payload) {
  const ss = getSS();
  const bankSheet = ss.getSheetByName(DB_SHEETS.BANK);
  const bankData = bankSheet.getDataRange().getValues();
  
  let correctCount = 0;
  const resultDetails = []; 
  const keyLabels = ['A', 'B', 'C', 'D'];

  payload.answers.forEach(ans => {
    const qKey = normalizeStr(ans.question);
    const targetRow = bankData.find(row => String(row[1]) === payload.testName && normalizeStr(row[2]) === qKey);
    
    if (!targetRow) {
      resultDetails.push({ question: ans.question, isCorrect: false, userSelected: ans.selected, correctTexts: ["Lỗi"], hint: "" });
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

    resultDetails.push({ question: targetRow[2], isCorrect: isCorrect, userSelected: userSelectedArr, correctTexts: correctTexts, hint: hint });
  });
  
  // Calculate Bonus
  const metaSheet = ss.getSheetByName(DB_SHEETS.TESTS);
  const metaData = metaSheet.getDataRange().getValues();
  let numQuestions = 0, duration = 1;
  for (let i = 1; i < metaData.length; i++) {
    if (metaData[i][1] === payload.testName) {
      numQuestions = Number(metaData[i][2]); 
      duration = Number(metaData[i][3]);     
      break;
    }
  }

  const timeSaving = parseFloat(payload.timeSaving) || 0;
  let bonus = (duration > 0 && numQuestions > 0) ? 0.02 * (numQuestions * timeSaving / duration) : 0;
  const totalResult = (correctCount + bonus).toFixed(2);
  
  // Save
  const resultSheet = ss.getSheetByName(DB_SHEETS.RESULTS);
  const lastRow = resultSheet.getLastRow();
  let newId = 1;
  if (lastRow > 1) {
    newId = Number(resultSheet.getRange(lastRow, 1).getValue()) + 1;
    if (isNaN(newId)) newId = 1;
  }
  
  resultSheet.appendRow([newId, new Date(), payload.testName, payload.candidate, payload.position, correctCount, timeSaving.toFixed(2), totalResult, payload.hintsUsed || 0]);
  
  return { score: totalResult, hintsUsed: payload.hintsUsed || 0, details: resultDetails };
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

function getDiscTestConfig(testName) {
  const ss = getDiscSS();
  const sheet = ss.getSheetByName(DB_SHEETS.DISC_TESTS);
  if (!sheet) throw new Error("DISC Database not found");
  const data = sheet.getDataRange().getValues();
  
  // Simple mapping based on known columns in PDF (Name at B/1)
  // Cols: Test Name(1), Duration(3), BlockDur(4), Occassion(6), SaveAnswer(7)
  // Headers check
  const headers = data[0];
  const colMap = { name: 1, duration: 3, blockDuration: 4, occasion: 6, saveAnswer: 7 };
  headers.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    if(header === 'test name') colMap.name = i;
    if(header === 'duration test') colMap.duration = i;
    if(header === 'duration each block question') colMap.blockDuration = i;
    if(header === 'occassion test') colMap.occasion = i;
    if(header === 'save answer sheet') colMap.saveAnswer = i;
  });

  for (let i = 1; i < data.length; i++) {
    if (data[i][colMap.name] === testName) {
      return {
        type: 'DISC',
        name: data[i][colMap.name],
        duration: data[i][colMap.duration],
        blockDuration: data[i][colMap.blockDuration],
        occasion: data[i][colMap.occasion],
        saveAnswer: data[i][colMap.saveAnswer]
      };
    }
  }
  throw new Error("Không tìm thấy thông tin bài test DISC");
}

function getDiscQuestions(testName) {
  const ss = getDiscSS();
  const config = getDiscTestConfig(testName);
  const qSheet = ss.getSheetByName(DB_SHEETS.DISC_QUESTIONS);
  if (!qSheet) throw new Error("Sheet DISC-Questions not found");
  
  const qData = qSheet.getDataRange().getValues();
  const headers = qData[0];
  const colMap = { id: 0, question: 1, most: 2, least: 3 }; // Default
  
  headers.forEach((h, i) => {
    const header = String(h).toLowerCase().trim();
    if (header === 'id') colMap.id = i;
    if (header.includes('question') || header.includes('nội dung')) colMap.question = i;
    if (header.includes('most') || header.includes('giống')) colMap.most = i;
    if (header.includes('least') || header.includes('khác')) colMap.least = i;
  });

  const questions = [];
  for(let i=1; i<qData.length; i++) {
    const row = qData[i];
    if(!row[colMap.id]) continue;
    questions.push({
      id: row[colMap.id],
      content: row[colMap.question],
      mostValue: row[colMap.most],
      leastValue: row[colMap.least]
    });
  }
  return { config: config, questions: questions };
}

function submitDiscTest(payload) {
  const ss = getHistorySS();
  const data = getDiscQuestions(payload.testName);
  const config = data.config;
  const questions = data.questions;

  // Generate Submission ID
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let submissionId = "";
  for (let i = 0; i < 2; i++) submissionId += chars.charAt(Math.floor(Math.random() * chars.length));
  
  const now = new Date();
  const timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yy HH:mm");

  // 1. SAVE RAW ANSWERS
  if (String(config.saveAnswer).toUpperCase().trim() === 'X') {
    const sheetAns = ss.getSheetByName(DB_SHEETS.DISC_ANSWERS);
    if(sheetAns) {
      const rows = [];
      questions.forEach(q => {
        const qId = String(q.id);
        const sel = payload.selections[qId];
        let mostVal = (sel === 'MOST') ? q.mostValue : "";
        let leastVal = (sel === 'LEAST') ? q.leastValue : "";
        rows.push([submissionId, timestampStr, config.occasion, payload.candidate, payload.position, q.content, mostVal, leastVal]);
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
       newResultId = Number(sheetRes.getRange(lastRow, 1).getValue()) + 1;
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

    sheetRes.appendRow([
      newResultId, timestampStr, config.occasion, payload.candidate, payload.position,
      payload.currentDisc || "", payload.trendDisc || "",
      counts.LEAST.D, counts.LEAST.I, counts.LEAST.S, counts.LEAST.C,
      counts.MOST.D, counts.MOST.I, counts.MOST.S, counts.MOST.C
    ]);
  }

  return { status: 'success', timestamp: timestampStr };
}

function getDiscResultsForAdmin() {
  const ss = getHistorySS();
  const sheet = ss.getSheetByName(DB_SHEETS.DISC_RESULTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  
  return data.map(row => ({
    id: row[0], timestamp: row[1], occasion: row[2], name: row[3], position: row[4],
    currentDiscStr: row[5], trendDiscStr: row[6],
    least: { D: row[7], I: row[8], S: row[9], C: row[10] },
    most: { D: row[11], I: row[12], S: row[13], C: row[14] },
    fitPoint: row[15] || 0, violatePoint: row[16] || 0
  })).reverse();
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
