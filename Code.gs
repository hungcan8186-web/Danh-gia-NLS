/**
 * Hàm khởi chạy giao diện
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('EduDigital Rubric Pro - Đánh giá năng lực số')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * CẤU HÌNH DỮ LIỆU CỦA BẠN
 * 1. Thay bằng ID Google Sheets của bạn để lưu cơ sở dữ liệu
 * 2. Thay bằng ID Thư mục Google Drive để lưu file báo cáo
 */
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; 
const DRIVE_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE'; 
const DATA_SHEET = 'DanhSachTheoDoi';

// ==========================================
// 1. MODULE: QUẢN LÝ DATABASE VÀ TRẠNG THÁI
// ==========================================

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET);
    sheet.appendRow(["Email", "Họ và tên", "Vai trò", "Trường", "Trạng thái", "Ngày cập nhật", "File Nháp (JSON)", "Điểm số", "Xếp loại", "Link Kết quả Docs"]);
    sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  return sheet;
}

// Tìm dòng dữ liệu theo Email
function findUserRow(sheet, email) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
      return i + 1; // Hàng thứ i+1 (Google Sheet 1-indexed)
    }
  }
  return -1;
}

// API: Tải dữ liệu nháp
function loadDraftData(email) {
  try {
    const sheet = getSheet();
    const row = findUserRow(sheet, email);
    if (row > -1) {
      const data = sheet.getRange(row, 1, 1, 10).getValues()[0];
      return {
        success: true,
        email: data[0],
        name: data[1],
        role: data[2],
        school: data[3],
        status: data[4],
        draftJson: data[6] || null
      };
    }
    return { success: false, message: "Không tìm thấy dữ liệu cũ." };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// API: Lưu nháp
function saveDraftData(payload) {
  try {
    const sheet = getSheet();
    const row = findUserRow(sheet, payload.email);
    const dateStr = new Date().toLocaleString('vi-VN');
    const jsonStr = JSON.stringify(payload.draftData); // Chứa array điểm số
    
    // Nếu có rồi thì cập nhật
    if (row > -1) {
      sheet.getRange(row, 2, 1, 6).setValues([[
        payload.name,
        payload.role,
        payload.school,
        "Đang đánh giá", // status
        dateStr, // time
        jsonStr // JSON
      ]]);
    } else {
      // Nếu chưa có thì thêm dòng mới
      sheet.appendRow([
        payload.email, payload.name, payload.role, payload.school, "Đang đánh giá", dateStr, jsonStr, "", "", ""
      ]);
    }
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// API: Load Dashboard
function getDashboardStats() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const stats = [];
    // Skip header
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        stats.push({
          email: data[i][0],
          name: data[i][1],
          role: data[i][2],
          school: data[i][3],
          status: data[i][4], // Đang đánh giá, Đã hoàn thành
          updateTime: data[i][5],
          scoreStr: data[i][7],
          classification: data[i][8]
        });
      }
    }
    return { success: true, list: stats };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 2. MODULE: LƯU GOOGLE DOCS (KHI NỘP BÀI)
// ==========================================

function submitFinalAssessment(payload) {
  try {
    const dateStr = new Date().toLocaleString('vi-VN');
    
    // 1. Cập nhật Google Sheets
    const sheet = getSheet();
    const row = findUserRow(sheet, payload.email);
    const jsonStr = JSON.stringify(payload.draftData);
    const scoreStr = payload.stats.totalScore + " / " + payload.stats.maxScore;
    const classStr = payload.stats.classification;

    // 2. Tạo File Google Docs
    const docName = "Phiếu ĐGNLS " + (payload.role === 'GV' ? "GV - " : "CBQL - ") + payload.name + " (" + payload.email + ")";
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    
    // Định dạng tiêu đề
    const title = body.insertParagraph(0, "KẾT QUẢ ĐÁNH GIÁ NĂNG LỰC SỐ");
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    body.appendParagraph("Họ và tên: " + payload.name);
    body.appendParagraph("Email: " + payload.email);
    body.appendParagraph("Đơn vị: " + payload.school);
    body.appendParagraph("Thời gian đánh giá: " + dateStr);
    body.appendParagraph("Tổng điểm: " + scoreStr + " (" + payload.stats.percentage + "%)");
    body.appendParagraph("Xếp loại: " + classStr);
    body.appendParagraph("").setSpacingAfter(20);
    
    // Bỏ dữ liệu đánh giá chung (Comments)
    body.appendParagraph("I. ĐÁNH GIÁ CHUNG").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph(payload.evaluation.comments || "Không có");
    body.appendParagraph("Phát triển/Bồi dưỡng: " + (payload.evaluation.suggestions || payload.evaluation.development || "Không có"));
    body.appendParagraph("").setSpacingAfter(20);
    
    // Lặp chi tiết rubric đưa vào Docs
    body.appendParagraph("II. ĐIỂM CHI TIẾT TỪNG TIÊU CHÍ").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    
    const rubricData = payload.draftData; // Array của criteria
    rubricData.forEach(domain => {
      body.appendParagraph(domain.name).setBold(true);
      domain.items.forEach(item => {
        let text = item.id + ". " + item.name + " \n";
        text += "-> Điểm: " + (item.score || 0) + "đ\n";
        text += "-> Minh chứng: " + (item.evidence || "Không có");
        body.appendParagraph(text);
      });
      body.appendParagraph("");
    });
    
    doc.saveAndClose();
    
    // 3. Chuyển File tới Thư Mục mong muốn (DRIVE_FOLDER_ID)
    let docUrl = doc.getUrl();
    if (DRIVE_FOLDER_ID !== 'YOUR_DRIVE_FOLDER_ID_HERE') {
      try {
        const file = DriveApp.getFileById(doc.getId());
        const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        file.moveTo(folder); // Move to specific folder
      } catch (e) {
        // Ignored if API permission or folder lacking, let it stay in Root.
        docUrl += " (Warning: Cannot move to folder)";
      }
    }
    
    // 4. Update Row status
    const recordValues = [
      payload.name, payload.role, payload.school, "Hoàn thành", dateStr, jsonStr, scoreStr, classStr, docUrl
    ];
    if (row > -1) {
      sheet.getRange(row, 2, 1, 9).setValues([recordValues]);
    } else {
      sheet.appendRow([payload.email, ...recordValues]);
    }

    return { success: true, docUrl: docUrl };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}


