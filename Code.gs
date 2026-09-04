/**
 * BACKEND CBT SYSTEM - SMAI AL AZHAR 15 SEMARANG
 * Menggunakan Google Sheets sebagai Database
 * Versi: Optimized (LockService & Strict Equality Check)
 */

const USERS_SHEET = "Users";
const EXAMS_SHEET = "Exams";
const SCORES_SHEET = "Scores";
const TOKENS_SHEET = "Tokens";
const SESSIONS_SHEET = "Sessions";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Portal CBT SMAI Al Azhar 15")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0",
    );
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let wsUsers = ss.getSheetByName(USERS_SHEET);
  if (!wsUsers) {
    wsUsers = ss.insertSheet(USERS_SHEET);
    wsUsers.appendRow([
      "Username",
      "Password",
      "Name",
      "Role",
      "Class",
      "Mapel",
    ]);
    wsUsers.appendRow([
      "siswa1",
      "123",
      "Budi Santoso",
      "user",
      "XI - MIPA 1",
      "",
    ]);
    wsUsers.appendRow([
      "guru1",
      "123",
      "Drs. H. Sugeng",
      "admin",
      "",
      "Sosiologi",
    ]);
    wsUsers.appendRow([
      "superadmin",
      "123",
      "Admin Sekolah",
      "superadmin",
      "",
      "",
    ]);
    wsUsers.getRange("A1:F1").setFontWeight("bold").setBackground("#d1d5db");
  }

  let wsExams = ss.getSheetByName(EXAMS_SHEET);
  if (!wsExams) {
    wsExams = ss.insertSheet(EXAMS_SHEET);
    wsExams.appendRow([
      "ExamID",
      "Title",
      "Mapel",
      "Duration",
      "TargetClassesJSON",
      "QuestionsJSON",
      "StartTime",
      "EndTime",
    ]);
    wsExams.getRange("A1:H1").setFontWeight("bold").setBackground("#d1d5db");
  }

  let wsScores = ss.getSheetByName(SCORES_SHEET);
  if (!wsScores) {
    wsScores = ss.insertSheet(SCORES_SHEET);
    wsScores.appendRow([
      "ExamID",
      "Username",
      "Score",
      "AnswersJSON",
      "Timestamp",
    ]);
    wsScores.getRange("A1:E1").setFontWeight("bold").setBackground("#d1d5db");
  }

  let wsTokens = ss.getSheetByName(TOKENS_SHEET);
  if (!wsTokens) {
    wsTokens = ss.insertSheet(TOKENS_SHEET);
    wsTokens.appendRow([
      "Token",
      "Type",
      "ExamID",
      "Username",
      "Status",
      "CreatedAt",
      "UsedAt",
    ]);
    wsTokens.getRange("A1:G1").setFontWeight("bold").setBackground("#d1d5db");
  }

  let wsSessions = ss.getSheetByName(SESSIONS_SHEET);
  if (!wsSessions) {
    wsSessions = ss.insertSheet(SESSIONS_SHEET);
    wsSessions.appendRow([
      "ExamID",
      "Username",
      "Status",
      "CheatCount",
      "UnlockCount",
      "LastUpdate",
    ]);
    wsSessions.getRange("A1:F1").setFontWeight("bold").setBackground("#d1d5db");
  }

  return "Database berhasil disiapkan!";
}

// ==========================================
// 1. MANAJEMEN AUTENTIKASI
// ==========================================

function apiLogin(payload) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let sheetUsername = String(data[i][0]).trim();
    let sheetPassword = String(data[i][1]).trim();
    let inputUsername = String(payload.username).trim();
    let inputPassword = String(payload.password).trim();

    if (sheetUsername === inputUsername && sheetPassword === inputPassword) {
      return {
        username: data[i][0],
        name: data[i][2],
        role: data[i][3],
        class: data[i][4],
        mapel: data[i][5],
      };
    }
  }
  return null;
}

function apiChangePassword(payload) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let sheetUsername = String(data[i][0]).trim();
    let sheetPassword = String(data[i][1]).trim();

    if (sheetUsername === String(payload.username).trim()) {
      // PERBAIKAN: Pastikan input password lama dari payload dijadikan string sebelum dikomparasi
      if (sheetPassword === String(payload.oldPassword).trim()) {
        sheet.getRange(i + 1, 2).setValue(String(payload.newPassword).trim());
        return { success: true };
      } else {
        return {
          success: false,
          message: "Password lama yang Anda masukkan salah.",
        };
      }
    }
  }
  return { success: false, message: "Pengguna tidak ditemukan." };
}

// ==========================================
// 2. MANAJEMEN USER (SUPERADMIN)
// ==========================================

function apiGetAllUsers() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  let users = [];
  for (let i = 1; i < data.length; i++) {
    users.push({
      username: String(data[i][0]).trim(),
      password: String(data[i][1]).trim(),
      name: String(data[i][2]).trim(),
      role: String(data[i][3]).trim(),
      class: String(data[i][4]).trim(),
      mapel: String(data[i][5]).trim(),
    });
  }
  return users;
}

function apiSaveUser(payload) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.username).trim()) {
      if (!payload.isNew) {
        sheet.getRange(i + 1, 2).setValue(payload.password);
        sheet.getRange(i + 1, 3).setValue(payload.name);
        sheet.getRange(i + 1, 4).setValue(payload.role);
        sheet.getRange(i + 1, 5).setValue(payload.class);
        sheet.getRange(i + 1, 6).setValue(payload.mapel);
        found = true;
        break;
      } else {
        return {
          success: false,
          message: "Username sudah terdaftar di sistem.",
        };
      }
    }
  }

  if (!found) {
    if (payload.isNew) {
      sheet.appendRow([
        payload.username,
        payload.password,
        payload.name,
        payload.role,
        payload.class,
        payload.mapel,
      ]);
    } else {
      return {
        success: false,
        message: "User tidak ditemukan untuk diupdate.",
      };
    }
  }
  return { success: true };
}

function apiDeleteUser(payload) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.username).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "User tidak ditemukan." };
}

// ==========================================
// 3. MANAJEMEN UJIAN & SOAL
// ==========================================

function apiGetClasses() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  const data = sheet.getDataRange().getValues();
  let classes = [];
  for (let i = 1; i < data.length; i++) {
    let role = String(data[i][3]).trim();
    let className = String(data[i][4]).trim();
    if (role === "user" && className !== "") {
      if (!classes.includes(className)) classes.push(className);
    }
  }
  return classes.sort();
}

function apiGetExams(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EXAMS_SHEET);
  const data = sheet.getDataRange().getValues();

  let completedExams = [];
  if (payload.role === "user" && payload.username) {
    const scoresSheet = ss.getSheetByName(SCORES_SHEET);
    if (scoresSheet) {
      const scoresData = scoresSheet.getDataRange().getValues();
      for (let s = 1; s < scoresData.length; s++) {
        if (String(scoresData[s][1]).trim() === String(payload.username).trim())
          completedExams.push(String(scoresData[s][0]).trim());
      }
    }
  }

  let results = [];
  for (let i = 1; i < data.length; i++) {
    let classes = [];
    try {
      classes = JSON.parse(data[i][4]);
    } catch (e) {}

    let examId = String(data[i][0]).trim();
    let exam = {
      id: examId,
      title: data[i][1],
      mapel: data[i][2],
      duration: data[i][3],
      targetClasses: classes,
      questionCount: 0,
      isCompleted: completedExams.includes(examId),
      startTime: data[i][6] ? String(data[i][6]).replace(/^'/, "") : "",
      endTime: data[i][7] ? String(data[i][7]).replace(/^'/, "") : "",
    };

    if (exam.startTime && exam.endTime) {
      const now = new Date();
      const start = new Date(exam.startTime);
      const end = new Date(exam.endTime);
      if (now < start) exam.timeStatus = "Belum Mulai";
      else if (now > end) exam.timeStatus = "Susulan";
      else exam.timeStatus = "Aktif";
    } else {
      exam.timeStatus = "Aktif";
    }

    try {
      let qs = JSON.parse(String(data[i][5] || "[]"));
      if (Array.isArray(qs)) exam.questionCount = qs.length;
    } catch (e) {}

    if (payload.role === "superadmin") results.push(exam);
    else if (payload.role === "admin" && exam.mapel === payload.filterValue)
      results.push(exam);
    else if (
      payload.role === "user" &&
      exam.targetClasses.includes(payload.filterValue)
    )
      results.push(exam);
  }
  return results;
}

function apiSaveExamMetadata(payload) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXAMS_SHEET);
  const data = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    // PERBAIKAN: Cast to String
    if (String(data[i][0]).trim() === String(payload.id).trim()) {
      sheet.getRange(i + 1, 2).setValue(payload.title);
      sheet.getRange(i + 1, 3).setValue(payload.mapel);
      sheet.getRange(i + 1, 4).setValue(payload.duration);
      sheet.getRange(i + 1, 5).setValue(payload.targetClassesJSON);
      sheet
        .getRange(i + 1, 7)
        .setValue(payload.startTime ? "'" + payload.startTime : "");
      sheet
        .getRange(i + 1, 8)
        .setValue(payload.endTime ? "'" + payload.endTime : "");
      found = true;
      break;
    }
  }
  if (!found)
    sheet.appendRow([
      payload.id,
      payload.title,
      payload.mapel,
      payload.duration,
      payload.targetClassesJSON,
      "[]",
      payload.startTime ? "'" + payload.startTime : "",
      payload.endTime ? "'" + payload.endTime : "",
    ]);
  return true;
}

function apiDeleteExam(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsExams = ss.getSheetByName(EXAMS_SHEET);
  const data = wsExams.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    // PERBAIKAN: Cast to String
    if (String(data[i][0]).trim() === String(payload.examId).trim()) {
      wsExams.deleteRow(i + 1);
      break;
    }
  }

  const wsScores = ss.getSheetByName(SCORES_SHEET);
  if (wsScores) {
    const scores = wsScores.getDataRange().getValues();
    for (let i = scores.length - 1; i >= 1; i--) {
      // PERBAIKAN: Cast to String
      if (String(scores[i][0]).trim() === String(payload.examId).trim())
        wsScores.deleteRow(i + 1);
    }
  }
  return true;
}

function getMediaFolder() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty("CBT_MEDIA_FOLDER_ID");

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      console.warn(
        "Folder media default tidak valid, akan dibuat ulang: " + e.message,
      );
    }
  }

  const folders = DriveApp.getFoldersByName("CBT Media");
  if (folders.hasNext()) {
    return folders.next();
  }

  const folder = DriveApp.createFolder("CBT Media");
  props.setProperty("CBT_MEDIA_FOLDER_ID", folder.getId());
  return folder;
}

function apiUploadMedia(file) {
  try {
    if (!file || !file.getName || !file.getBlob) {
      return { success: false, message: "File upload tidak valid." };
    }

    const mimeType =
      file.getContentType && file.getContentType()
        ? file.getContentType()
        : "image/png";
    if (!mimeType.startsWith("image/")) {
      return { success: false, message: "Hanya file gambar yang diizinkan." };
    }

    const folder = getMediaFolder();
    const extension =
      file.getName() && file.getName().includes(".")
        ? file.getName().split(".").pop()
        : mimeType === "image/png"
          ? "png"
          : "jpg";

    const safeName =
      "cbt-media-" +
      Date.now() +
      "-" +
      Math.random().toString(36).slice(2, 8) +
      "." +
      extension;
    const blob = file.getBlob ? file.getBlob() : file;
    blob.setName(safeName);

    const savedFile = folder.createFile(blob);
    if (!savedFile || !savedFile.getId()) {
      return {
        success: false,
        message: "File tidak berhasil disimpan ke Drive.",
      };
    }

    savedFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW,
    );

    return {
      success: true,
      fileId: savedFile.getId(),
      url: "https://drive.google.com/uc?export=view&id=" + savedFile.getId(),
      fileName: savedFile.getName(),
      message: "Gambar berhasil diunggah.",
    };
  } catch (e) {
    console.error("apiUploadMedia error:", e);
    return {
      success: false,
      message:
        "Upload gambar gagal: " +
        (e && e.message ? e.message : "error tidak diketahui"),
    };
  }
}

function apiGetQuestions(payload) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXAMS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // PERBAIKAN: Cast to String
    if (String(data[i][0]).trim() === String(payload.examId).trim()) {
      try {
        return JSON.parse(String(data[i][5] || "[]"));
      } catch (e) {
        return [];
      }
    }
  }
  return [];
}

function apiGetQuestionsForExam(payload) {
  const questions = apiGetQuestions(payload);
  return questions.map(function (q) {
    return {
      type: q.type,
      text: q.text,
      options: q.options || [],
      prompts: q.prompts || [],
    };
  });
}

function apiSaveQuestions(payload) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXAMS_SHEET);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // PERBAIKAN: Cast to String
    if (String(data[i][0]).trim() === String(payload.examId).trim()) {
      sheet.getRange(i + 1, 6).setValue(payload.questionsJSON);
      return true;
    }
  }
  return false;
}

// ==========================================
// 4. PENILAIAN & SCORING
// ==========================================

function calculateScore(questions, answers) {
  if (!Array.isArray(questions) || questions.length === 0) return 0;
  const poinPerSoal = 100 / questions.length;
  let totalPoin = 0;

  questions.forEach(function (q, idx) {
    const jawabanSiswa = answers[idx];
    if (jawabanSiswa === undefined || jawabanSiswa === null) return;

    if (q.type === "pg") {
      if (jawabanSiswa === q.answerKey) totalPoin += poinPerSoal;
    } else if (q.type === "pgk") {
      const kunci = Array.isArray(q.answerKey) ? q.answerKey.slice() : [];
      const jawaban = Array.isArray(jawabanSiswa) ? jawabanSiswa.slice() : [];
      if (kunci.length > 0) {
        let benar = 0;
        jawaban.forEach(function (j) {
          if (kunci.includes(j)) benar++;
          else benar--;
        });
        let finalPGK = (benar / kunci.length) * poinPerSoal;
        if (finalPGK > 0) totalPoin += finalPGK;
      }
    } else if (q.type === "bs" || q.type === "jodoh") {
      const kunci = q.answerKey || {};
      const totalPasangan = Object.keys(kunci).length;
      if (totalPasangan === 0) return;
      let benar = 0;
      for (const idxPasangan in kunci) {
        if (jawabanSiswa[idxPasangan] === kunci[idxPasangan]) benar++;
      }
      totalPoin += (benar / totalPasangan) * poinPerSoal;
    } else if (q.type === "uraian") {
      if (
        typeof jawabanSiswa === "object" &&
        jawabanSiswa !== null &&
        jawabanSiswa.score !== undefined
      ) {
        totalPoin += (parseFloat(jawabanSiswa.score) / 100) * poinPerSoal;
      }
    }
  });
  return Math.round(totalPoin * 100) / 100;
}

function apiSaveEssayScore(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsScores = ss.getSheetByName(SCORES_SHEET);
  if (!wsScores) return { success: false };

  let questions = apiGetQuestions({ examId: payload.examId });

  const scoreData = wsScores.getDataRange().getValues();
  for (let i = 1; i < scoreData.length; i++) {
    if (
      String(scoreData[i][0]).trim() === String(payload.examId).trim() &&
      String(scoreData[i][1]).trim() === String(payload.username).trim()
    ) {
      let answers = {};
      try {
        answers = JSON.parse(scoreData[i][3]) || {};
      } catch (e) {
        answers = {};
      }

      if (!answers[payload.qIndex])
        answers[payload.qIndex] = { text: "", score: 0 };
      else if (typeof answers[payload.qIndex] === "string")
        answers[payload.qIndex] = { text: answers[payload.qIndex], score: 0 };

      answers[payload.qIndex].score = payload.score;

      const finalScore = calculateScore(questions, answers);

      wsScores.getRange(i + 1, 3).setValue(finalScore);
      wsScores.getRange(i + 1, 4).setValue(JSON.stringify(answers));

      return { success: true, newTotalScore: finalScore };
    }
  }
  return { success: false, message: "Data pengerjaan siswa tidak ditemukan." };
}

function apiSubmitExam(payload) {
  // PERBAIKAN: Menggunakan LockService untuk mencegah Race Conditions saat ujian selesai masal
  const lock = LockService.getScriptLock();

  try {
    // Tunggu antrean eksekusi sampai maksimum 10 detik
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wsScores = ss.getSheetByName(SCORES_SHEET);
    const wsSessions = ss.getSheetByName(SESSIONS_SHEET);
    const timestamp = new Date().toLocaleString("id-ID");

    let questions = apiGetQuestions({ examId: payload.examId });
    let answers = {};
    try {
      answers = JSON.parse(payload.answersJSON) || {};
    } catch (e) {
      answers = {};
    }

    const finalScore = calculateScore(questions, answers);
    wsScores.appendRow([
      payload.examId,
      payload.username,
      finalScore,
      payload.answersJSON,
      timestamp,
    ]);

    if (wsSessions) {
      const sessData = wsSessions.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < sessData.length; i++) {
        // PERBAIKAN: Cast to string
        if (
          String(sessData[i][0]).trim() === String(payload.examId).trim() &&
          String(sessData[i][1]).trim() === String(payload.username).trim()
        ) {
          wsSessions.getRange(i + 1, 3).setValue("FINISHED");
          found = true;
          break;
        }
      }
      if (!found)
        wsSessions.appendRow([
          payload.examId,
          payload.username,
          "FINISHED",
          0,
          0,
          timestamp,
        ]);
    }
    return true;
  } catch (e) {
    // Jika lock timeout atau error
    return {
      success: false,
      message:
        "Server sangat sibuk karena antrean pengumpulan. Silakan coba klik submit lagi.",
    };
  } finally {
    // PERBAIKAN: Lock wajib dilepaskan agar antrean berikutnya jalan
    lock.releaseLock();
  }
}

// ==========================================
// 5. ANALITIK DATA
// ==========================================

function apiGetAnalytics(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsUsers = ss.getSheetByName(USERS_SHEET);
  const wsExams = ss.getSheetByName(EXAMS_SHEET);
  const wsScores = ss.getSheetByName(SCORES_SHEET);

  const userData = wsUsers.getDataRange().getValues();
  const examData = wsExams.getDataRange().getValues();
  const scoreData = wsScores.getDataRange().getValues();

  let targetExams = [];
  for (let i = 1; i < examData.length; i++) {
    if (
      payload.role === "superadmin" ||
      String(examData[i][2]).trim() === payload.mapel
    ) {
      if (
        payload.examId === "all" ||
        String(examData[i][0]).trim() === String(payload.examId).trim()
      ) {
        let cls = [];
        try {
          cls = JSON.parse(examData[i][4]);
        } catch (e) {}
        targetExams.push({
          id: String(examData[i][0]).trim(),
          title: examData[i][1],
          classes: cls,
          completed: 0,
          avg: 0,
        });
      }
    }
  }
  const targetExamIds = targetExams.map((e) => e.id);

  let classByUser = {};
  let targetUsers = [];
  for (let i = 1; i < userData.length; i++) {
    if (String(userData[i][3]).trim() === "user") {
      const uname = String(userData[i][0]).trim();
      const cls = String(userData[i][4]).trim();
      classByUser[uname] = cls;
      if (payload.className === "all" || cls === payload.className)
        targetUsers.push(uname);
    }
  }

  let nilaiList = [];
  let examStatsMap = {};
  targetExams.forEach((e) => (examStatsMap[e.id] = { sum: 0, count: 0 }));

  for (let i = 1; i < scoreData.length; i++) {
    const examId = String(scoreData[i][0]).trim();
    const username = String(scoreData[i][1]).trim();
    const nilai = parseFloat(scoreData[i][2]);

    if (
      !targetExamIds.includes(examId) ||
      !targetUsers.includes(username) ||
      isNaN(nilai)
    )
      continue;

    nilaiList.push(nilai);
    examStatsMap[examId].sum += nilai;
    examStatsMap[examId].count++;
  }

  targetExams.forEach((e) => {
    e.completed = examStatsMap[e.id].count;
    e.avg =
      e.completed > 0
        ? Math.round((examStatsMap[e.id].sum / e.completed) * 100) / 100
        : 0;
  });

  let avgScore = 0;
  let distData = [0, 0, 0, 0];
  let pieData = [0, 0, 0, 0];
  if (nilaiList.length > 0) {
    const total = nilaiList.reduce((a, b) => a + b, 0);
    avgScore = Math.round((total / nilaiList.length) * 100) / 100;
    nilaiList.forEach((n) => {
      if (n < 55) distData[0]++;
      else if (n < 70) distData[1]++;
      else if (n < 85) distData[2]++;
      else distData[3]++;
    });
    pieData = [distData[3], distData[2], distData[1], distData[0]];
  }

  if (payload.className !== "all")
    targetExams = targetExams.filter((e) =>
      e.classes.includes(payload.className),
    );

  return {
    examStats: targetExams,
    totalStudentData: targetUsers.length,
    avgScore: avgScore,
    distData: distData,
    pieData: pieData,
  };
}

function apiGetDetailedAnalytics(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsUsers = ss.getSheetByName(USERS_SHEET);
  const wsExams = ss.getSheetByName(EXAMS_SHEET);
  const wsScores = ss.getSheetByName(SCORES_SHEET);
  const wsSessions = ss.getSheetByName(SESSIONS_SHEET);

  const userData = wsUsers.getDataRange().getValues();
  const examData = wsExams.getDataRange().getValues();
  const scoreData = wsScores.getDataRange().getValues();

  let exam = null;
  for (let i = 1; i < examData.length; i++) {
    // PERBAIKAN: Cast to String
    if (String(examData[i][0]).trim() === String(payload.examId).trim()) {
      let cls = [];
      try {
        cls = JSON.parse(examData[i][4]);
      } catch (e) {}
      let qs = apiGetQuestions({ examId: payload.examId });
      exam = {
        id: examData[i][0],
        title: examData[i][1],
        targetClasses: cls,
        questions: qs,
        questionCount: qs.length,
      };
      break;
    }
  }
  if (!exam) return null;

  let sessionMap = {};
  if (wsSessions) {
    const sessData = wsSessions.getDataRange().getValues();
    for (let i = 1; i < sessData.length; i++) {
      if (String(sessData[i][0]).trim() === String(payload.examId).trim()) {
        sessionMap[String(sessData[i][1]).trim()] = {
          status: sessData[i][2],
          cheatCount: sessData[i][3] || 0,
        };
      }
    }
  }

  let students = [];
  let classStatsMap = {};
  exam.targetClasses.forEach(
    (c) => (classStatsMap[c] = { total: 0, completed: 0, notCompleted: 0 }),
  );

  for (let i = 1; i < userData.length; i++) {
    if (
      String(userData[i][3]).trim() === "user" &&
      exam.targetClasses.includes(String(userData[i][4]).trim())
    ) {
      let uname = String(userData[i][0]).trim();
      let sess = sessionMap[uname] || { status: "NOT_STARTED", cheatCount: 0 };

      let u = {
        username: uname,
        name: String(userData[i][2]).trim(),
        class: String(userData[i][4]).trim(),
        status: "Belum Selesai",
        sessionStatus: sess.status,
        cheatCount: sess.cheatCount,
        score: null,
        answers: {},
      };
      students.push(u);
      if (classStatsMap[u.class]) classStatsMap[u.class].total++;
    }
  }

  for (let i = 1; i < scoreData.length; i++) {
    if (String(scoreData[i][0]).trim() === String(payload.examId).trim()) {
      let username = String(scoreData[i][1]).trim();
      let student = students.find((s) => s.username === username);
      if (student) {
        student.status = "Selesai";
        student.score = parseFloat(scoreData[i][2]);
        try {
          student.answers = JSON.parse(scoreData[i][3]) || {};
        } catch (e) {
          student.answers = {};
        }
        if (classStatsMap[student.class]) {
          classStatsMap[student.class].completed++;
          classStatsMap[student.class].notCompleted =
            classStatsMap[student.class].total -
            classStatsMap[student.class].completed;
        }
      }
    }
  }

  let classStats = Object.keys(classStatsMap).map((k) => ({
    className: k,
    total: classStatsMap[k].total,
    completed: classStatsMap[k].completed,
    notCompleted: classStatsMap[k].notCompleted,
  }));
  students.sort(
    (a, b) => a.class.localeCompare(b.class) || a.name.localeCompare(b.name),
  );

  return {
    examDetails: exam,
    classStats: classStats,
    studentDetails: students,
  };
}

// ==========================================
// 6. MANAJEMEN SESI & TOKEN
// ==========================================

function apiUpdateSession(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsSessions = ss.getSheetByName(SESSIONS_SHEET);
  if (!wsSessions) return false;

  const data = wsSessions.getDataRange().getValues();
  const ts = new Date().toLocaleString("id-ID");

  for (let i = 1; i < data.length; i++) {
    // PERBAIKAN: Cast to String
    if (
      String(data[i][0]).trim() === String(payload.examId).trim() &&
      String(data[i][1]).trim() === String(payload.username).trim()
    ) {
      if (payload.status)
        wsSessions.getRange(i + 1, 3).setValue(payload.status);
      if (payload.cheatCount !== undefined)
        wsSessions.getRange(i + 1, 4).setValue(payload.cheatCount);
      wsSessions.getRange(i + 1, 6).setValue(ts);
      return true;
    }
  }

  wsSessions.appendRow([
    payload.examId,
    payload.username,
    payload.status || "STARTED",
    payload.cheatCount || 0,
    0,
    ts,
  ]);
  return true;
}

function apiGetTokenCandidates(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsUsers = ss.getSheetByName(USERS_SHEET);
  const wsExams = ss.getSheetByName(EXAMS_SHEET);
  const wsScores = ss.getSheetByName(SCORES_SHEET);
  const wsSessions = ss.getSheetByName(SESSIONS_SHEET);
  const wsTokens = ss.getSheetByName(TOKENS_SHEET);

  let exam = null;
  const examData = wsExams.getDataRange().getValues();
  for (let i = 1; i < examData.length; i++) {
    // PERBAIKAN: Cast to String
    if (String(examData[i][0]).trim() === String(payload.examId).trim()) {
      let cls = [];
      try {
        cls = JSON.parse(examData[i][4]);
      } catch (e) {}
      let et = examData[i][7] ? String(examData[i][7]).replace(/^'/, "") : null;
      exam = { id: examData[i][0], targetClasses: cls, endTime: et };
      break;
    }
  }
  if (!exam) return { candidates: [], tokens: [] };

  const isPastTime = exam.endTime ? new Date() > new Date(exam.endTime) : false;

  const scoresData = wsScores.getDataRange().getValues();
  let finishedUsers = [];
  for (let i = 1; i < scoresData.length; i++) {
    if (String(scoresData[i][0]).trim() === String(exam.id).trim())
      finishedUsers.push(String(scoresData[i][1]).trim());
  }

  const sessionsData = wsSessions ? wsSessions.getDataRange().getValues() : [];
  let userSessions = {};
  for (let i = 1; i < sessionsData.length; i++) {
    if (String(sessionsData[i][0]).trim() === String(exam.id).trim()) {
      userSessions[String(sessionsData[i][1]).trim()] = {
        status: sessionsData[i][2],
        cheatCount: sessionsData[i][3],
        unlockCount: sessionsData[i][4],
      };
    }
  }

  const usersData = wsUsers.getDataRange().getValues();
  let candidates = [];
  for (let i = 1; i < usersData.length; i++) {
    const uRole = String(usersData[i][3]).trim();
    const uClass = String(usersData[i][4]).trim();

    if (
      uRole === "user" &&
      uClass === payload.className &&
      exam.targetClasses.includes(uClass)
    ) {
      const uname = String(usersData[i][0]).trim();
      if (finishedUsers.includes(uname)) continue;

      let sess = userSessions[uname] || {
        status: "NOT_STARTED",
        cheatCount: 0,
        unlockCount: 0,
      };
      let needsUnlock = sess.status === "LOCKED";
      let needsSusulan =
        (sess.status === "NOT_STARTED" || sess.status === "STARTED") &&
        isPastTime;

      if (needsUnlock || needsSusulan) {
        candidates.push({
          username: uname,
          name: String(usersData[i][2]).trim(),
          class: uClass,
          status: sess.status,
          cheatCount: sess.cheatCount,
          unlockCount: sess.unlockCount,
          needsUnlock: needsUnlock,
          needsSusulan: needsSusulan,
        });
      }
    }
  }

  let tokens = [];
  if (wsTokens) {
    const tokenData = wsTokens.getDataRange().getValues();
    let cNames = candidates.map((c) => c.username);
    for (let i = 1; i < tokenData.length; i++) {
      if (
        String(tokenData[i][2]).trim() === String(exam.id).trim() &&
        (!payload.className || cNames.includes(String(tokenData[i][3]).trim()))
      ) {
        tokens.push({
          token: String(tokenData[i][0]).trim(),
          type: tokenData[i][1],
          username: String(tokenData[i][3]).trim(),
          status: tokenData[i][4],
          createdAt: tokenData[i][5],
          usedAt: tokenData[i][6],
        });
      }
    }
  }

  return { candidates: candidates, tokens: tokens.reverse() };
}

function apiGenerateToken(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsTokens = ss.getSheetByName(TOKENS_SHEET);
  if (!wsTokens) return { success: false };

  const charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let tk = "";
  for (let i = 0; i < 6; i++)
    tk += charset.charAt(Math.floor(Math.random() * charset.length));

  const ts = new Date().toLocaleString("id-ID");
  wsTokens.appendRow([
    tk,
    payload.type,
    payload.examId,
    payload.username,
    "ACTIVE",
    ts,
    "",
  ]);

  return { success: true, token: tk };
}

function apiVerifyAndUseToken(payload) {
  // PERBAIKAN: Gunakan LockService karena token sering diketik siswa secara bersamaan
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000); // 5 Detik cukup untuk token

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wsTokens = ss.getSheetByName(TOKENS_SHEET);
    const data = wsTokens.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      // PERBAIKAN: Cast to String
      if (
        String(data[i][0]).trim() === String(payload.token).trim() &&
        String(data[i][1]).trim() === String(payload.type).trim() &&
        String(data[i][2]).trim() === String(payload.examId).trim() &&
        String(data[i][3]).trim() === String(payload.username).trim() &&
        data[i][4] === "ACTIVE"
      ) {
        wsTokens.getRange(i + 1, 5).setValue("USED");
        wsTokens
          .getRange(i + 1, 7)
          .setValue(new Date().toLocaleString("id-ID"));

        if (payload.type === "UNLOCK") {
          const wsSessions = ss.getSheetByName(SESSIONS_SHEET);
          if (wsSessions) {
            const sessData = wsSessions.getDataRange().getValues();
            for (let j = 1; j < sessData.length; j++) {
              if (
                String(sessData[j][0]).trim() ===
                  String(payload.examId).trim() &&
                String(sessData[j][1]).trim() ===
                  String(payload.username).trim()
              ) {
                wsSessions.getRange(j + 1, 3).setValue("STARTED");
                let currentUnlock = parseInt(sessData[j][4]) || 0;
                wsSessions.getRange(j + 1, 5).setValue(currentUnlock + 1);
                break;
              }
            }
          }
        }
        return { success: true };
      }
    }
    return {
      success: false,
      message: "Token tidak valid / sudah pernah digunakan.",
    };
  } catch (e) {
    return {
      success: false,
      message:
        "Sistem sedang memproses request lain, silakan klik Submit Token lagi.",
    };
  } finally {
    lock.releaseLock();
  }
}

function apiClearSystemData(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (payload.action === "clear_tokens") {
    const ws = ss.getSheetByName(TOKENS_SHEET);
    if (ws && ws.getLastRow() > 1) {
      ws.getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn()).clearContent();
    }
    return {
      success: true,
      message: "Semua data token ujian berhasil dihapus.",
    };
  }

  if (payload.action === "clear_sessions") {
    const ws = ss.getSheetByName(SESSIONS_SHEET);
    if (ws && ws.getLastRow() > 1) {
      ws.getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn()).clearContent();
    }
    return {
      success: true,
      message: "Semua data sesi pengerjaan aktif berhasil direset.",
    };
  }

  if (payload.action === "clear_scores_exam") {
    const ws = ss.getSheetByName(SCORES_SHEET);
    if (!ws) return { success: false };
    const data = ws.getDataRange().getValues();
    let deletedCount = 0;

    // Ini sudah benar secara algoritma menggunakan for loop mundur
    for (let i = data.length - 1; i >= 1; i--) {
      // PERBAIKAN: Cast to String
      if (String(data[i][0]).trim() === String(payload.examId).trim()) {
        ws.deleteRow(i + 1);
        deletedCount++;
      }
    }
    return {
      success: true,
      message: `Data nilai untuk ujian tersebut berhasil dihapus (${deletedCount} data).`,
    };
  }

  return { success: false };
}
