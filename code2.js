// --- FILE TỔNG TRÊN GITHUB ---07/07/26

// Lấy sheetId từ cột J (cột 10) hoặc cột I (cột 9 dự phòng) của sheet 'idgv' trong ssAdmin
function getSheetIdByIdgv(idgv) {
  if (!idgv) return "";
  try {
    if (typeof ssAdmin === "undefined" || !ssAdmin) return "";
    var sheet = ssAdmin.getSheetByName("idgv");
    if (!sheet) return "";
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "";
    
    // Lấy từ cột A đến cột J (cột 1 đến cột 10)
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var targetStr = String(idgv).replace(/'/g, "").trim().toUpperCase();
    var targetN9 = typeof N9 === "function" ? N9(idgv) : targetStr;

    for (var i = 0; i < data.length; i++) {
      var cellA = data[i][0];
      if (cellA === null || cellA === undefined || cellA === "") continue;

      var colAStr = String(cellA).replace(/'/g, "").trim().toUpperCase();
      var colAN9 = typeof N9 === "function" ? N9(cellA) : colAStr;

      // So sánh khớp IDGV
      var isMatch = (colAStr === targetStr) || 
                    (targetN9 && colAN9 === targetN9) || 
                    (typeof supper === "function" && supper(colAStr) === supper(targetStr));

      if (isMatch) {
        var sidJ = String(data[i][9] || "").trim(); // Cột J (Index 9)
        var sidI = String(data[i][8] || "").trim(); // Cột I (Index 8)

        // Lựa chọn ưu tiên: Kiểm tra cột J trước, nếu không hợp lệ thì dự phòng sang cột I
        if (sidJ && sidJ.length >= 25) {
          return sidJ;
        }
        if (sidI && sidI.length >= 25) {
          return sidI;
        }
      }
    }
  } catch (err) {
    Logger.log("Lỗi getSheetIdByIdgv: " + err.toString());
  }
  return "";
}

// Mở Spreadsheet của Giáo viên (ss2) theo sheetId hoặc tự động tra cứu từ idgv
function getSS2(sheetId, idgv) {
  var sid = String(sheetId || "").trim();
  
  // Loại bỏ các trường hợp truyền chuỗi "undefined" hoặc "null" từ Client
  if (sid === "undefined" || sid === "null" || sid.length < 25) {
    sid = "";
  }

  // Nếu không có sheetId chuẩn, tiến hành tra cứu qua idgv trong sheet 'idgv'
  if (!sid && idgv) {
    sid = getSheetIdByIdgv(idgv);
  }

  // Thực hiện mở Google Sheet theo ID tìm được
  if (sid && sid.length >= 25) {
    try {
      return SpreadsheetApp.openById(sid);
    } catch (e) {
      Logger.log("Không thể mở Spreadsheet theo sheetId [" + sid + "]: " + e.toString());
    }
  }

  // Fallback bắt buộc về File Sheet chính (ss) nếu không mở được file riêng
  return (typeof ss !== "undefined" && ss) ? ss : SpreadsheetApp.getActiveSpreadsheet();
}
// 2508sua1: Hàm getSS2Sheet mở Spreadsheet theo sheetId/idgv và trả về Sheet cần tìm
function getSS2Sheet(sheetId, idgv, sheetName) {
  // 2508sua1: Chuẩn hóa chuỗi sheetId
  var sid = String(sheetId || "").trim();
  
  // 2508sua1: Loại bỏ các trường hợp truyền chuỗi "undefined" hoặc "null" từ Client
  if (sid === "undefined" || sid === "null" || sid.length < 25) {
    // 2508sua1: Đặt lại sid rỗng nếu không hợp lệ
    sid = "";
  }

  // 2508sua1: Nếu không có sheetId chuẩn, tiến hành tra cứu qua idgv trong sheet 'idgv'
  if (!sid && idgv) {
    // 2508sua1: Tra cứu sheetId của giáo viên
    sid = getSheetIdByIdgv(idgv);
  }

  // 2508sua1: Thực hiện mở Google Sheet theo ID tìm được
  if (sid && sid.length >= 25) {
    // 2508sua1: Thử mở bảng tính theo ID
    try {
      // 2508sua1: Mở Spreadsheet ss2
      var ss2 = SpreadsheetApp.openById(sid);
      // 2508sua1: Lấy sheet theo tên từ ss2
      var sheet = ss2.getSheetByName(sheetName);
      // 2508sua1: Nếu chưa tồn tại sheet thì tạo mới
      if (!sheet) {
        // 2508sua1: Thêm sheet mới vào ss2
        sheet = ss2.insertSheet(sheetName);
      }
      // 2508sua1: Trả về đối tượng sheet
      return sheet; 
    } catch (e) {
      // 2508sua1: Ghi log nếu có lỗi khi mở file
      Logger.log("Không thể mở Spreadsheet theo sheetId [" + sid + "]: " + e.toString());
    }
  }

  // 2508sua1: Trả về null nếu không tìm thấy
  return null;
}

function mainDoGet(e) {
const params = e.parameter;
  const type = params.type;
  const action = params.action || e.parameter.action;
  // ACTION: Lấy dữ liệu đánh giá
  if (action === "getRating") {
    const sheetDG = ssAdmin.getSheetByName("danhgia");
    
    if (!sheetDG) {
      return ContentService.createTextOutput(JSON.stringify({
        avgStars: "5.0", totalReviews: 0, fiveStars: 0, fourStars: 0, threeStars: 0, twoStars: 0, oneStar: 0
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const values = sheetDG.getDataRange().getValues();
    let count5 = 0, count4 = 0, count3 = 0, count2 = 0, count1 = 0;
    let totalStars = 0;
    let totalReviews = 0;

    for (let i = 1; i < values.length; i++) {
      const rowStar = Number(values[i][1]);
      if (rowStar >= 1 && rowStar <= 5) {
        totalReviews++;
        totalStars += rowStar;
        if (rowStar === 5) count5++;
        if (rowStar === 4) count4++;
        if (rowStar === 3) count3++;
        if (rowStar === 2) count2++;
        if (rowStar === 1) count1++;
      }
    }

    const avgStars = totalReviews > 0 ? (totalStars / totalReviews).toFixed(1) : "5.0";

    return ContentService.createTextOutput(JSON.stringify({
      avgStars: avgStars,
      totalReviews: totalReviews,
      fiveStars: count5,
      fourStars: count4,
      threeStars: count3,
      twoStars: count2,
      oneStar: count1
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Xóa ảnh trong cloud của Giaovien
  // --- DAN NOI TIEP VAO TRONG HAM mainDoGet(e) ---
if (action === "adminResetCloudImages") {
  const tId = e.parameter.teacherId;
  const sId = e.parameter.subjectId;
  const cName = e.parameter.cloudName;
  const folderDe = e.parameter.folderDe; 
  const apiKey = e.parameter.apiKey;
  const apiSecret = e.parameter.apiSec;
  
  // Duong dan folder can quet sach anh: Giaovien/subjectId/teacherId
  const prefixFolder = folderDe + "/" + sId + "/" + tId;
  
  try {
    const timestamp = Math.floor(Date.now() / 1000).toString();
    
    // Tao chuoi ky xac thuc (Signature) dung chuan API Cloudinary de xoa tai nguyen
    const stringToSign = "prefix=" + prefixFolder + "&timestamp=" + timestamp + apiSecret;
    const signature = SHA256_(stringToSign); 
    
    // Goi API cua Cloudinary dung phuong thuc DELETE de xoa sach anh trong folder
    const url = "https://api.cloudinary.com/v1_1/" + cName + "/resources/image/upload";
    const payload = {
      "prefix": prefixFolder,
      "timestamp": timestamp,
      "api_key": apiKey,
      "signature": signature
    };
    
    const options = {
      "method": "delete",
      "payload": payload,
      "muteHttpExceptions": true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "message": "Da clear anh tren Cloud"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

  // Admin xóa dữ liệu
  // --- DAN ĐÈ HOẶC THÊM VÀO TRONG HÀM mainDoGet(e) ---
if (action === "adminResetSheet") {
  // Lay truc tiep tu tham so e.parameter (vi gui bang GET)
  const sheetName = e.parameter.sheet; 
  const sheet = ss.getSheetByName(sheetName);   
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": "Khong tim thay sheet: " + sheetName
    })).setMimeType(ContentService.MimeType.JSON);
  }     
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();   
  
  // Tien hanh xoa sach du lieu tu hang 2 neu co du lieu
  if (lastRow >= 2 && lastColumn > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
    SpreadsheetApp.flush(); // Ep Google update ngay lap tuc
  }   
  
  return ContentService.createTextOutput(JSON.stringify({
    "status": "success", 
    "message": "Da xoa sach du lieu sheet [" + sheetName + "] tu dong 2 den het!"
  })).setMimeType(ContentService.MimeType.JSON);                                 
}

  // Xóa ảnh cloud
  // --- THÊM NHÁNH NÀY VÀO TRONG HÀM mainDoGet(e) CỦA FILE TỔNG ---
if (action === "adminResetCloudImages") {
  const tId = e.parameter.teacherId;
  const sId = e.parameter.subjectId;
  const cName = e.parameter.cloudName;
  
  // 1. Điền thông tin cấu hình Cloudinary của thầy vào đây để hệ thống ký xác thực ngầm
  const apiKey = "THAY_DIEN_API_KEY_CLOUDINARY_VAO_DAY";
  const apiSecret = "THAY_DIEN_API_SECRET_CLOUDINARY_VAO_DAY";
  
  // Đường dẫn chính xác đến folder cần xóa: Giaovien/subjectId/teacherId
  const prefixFolder = "Giaovien/" + sId + "/" + tId;
  
  try {
    // Tạo mốc thời gian timestamp bắt buộc của Cloudinary
    const timestamp = Math.floor(Date.now() / 1000).toString();
    
    // Tạo chuỗi ký xác thực (Signature) theo chuẩn API Cloudinary để xóa thư mục
    const stringToSign = "prefix=" + prefixFolder + "&timestamp=" + timestamp + apiSecret;
    const signature = SHA256_(stringToSign); // Sử dụng hàm băm SHA256 phía dưới
    
    // Gọi API Cloudinary xóa tất cả tài nguyên (ảnh) có tiền tố đường dẫn này
    const url = "https://api.cloudinary.com/v1_1/" + cName + "/resources/image/upload";
    const payload = {
      "prefix": prefixFolder,
      "timestamp": timestamp,
      "api_key": apiKey,
      "signature": signature
    };
    
    const options = {
      "method": "delete",
      "payload": payload,
      "muteHttpExceptions": true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const resData = JSON.parse(response.getContentText());
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "message": "Da clear anh tren Cloud",
      "detail": resData
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}  
  
//#01
  // Xác minh và đăng nhập Game Show
  if (action === "registershow") {
    return register(e);
  }
  if (action === "loginshow") {
    return login(e);
  }
  if (action === "adminLoginshow") {
  return adminLogin(e);
  }
  if (action === "getUsersshow") {
    return getUsers(e);
  }
  if (action === "updatePassword") {
  return updatePassword(e);
}
  if (action === "getTeachers") {
  return getTeachers();
}
  // Xác minh giáo viên bên VBA
  // Xác minh chỉ idgv
  if (action === "getIdGV") {
  const sheet = ssAdmin.getSheetByName("idgv");
  const data = sheet.getRange("A2:A" + sheet.getLastRow())
                  .getValues()
                  .flat()
                  .map(item => String(item).slice(-9));
  
  // Lọc bỏ ô trống và chuyển về chữ thường
  const cleanData = data.filter(String).map(id => id.toString().toLowerCase().trim());
  
  // Trả về chuỗi thuần túy: "gv100,gv101,admin22"
  return ContentService.createTextOutput(cleanData.join(",")).setMimeType(ContentService.MimeType.TEXT);
}
 // Xác minh idgv và môn
  if (action === "getIdGVM") {
    const sheet = ssAdmin.getSheetByName("idgv");
    const data = sheet.getRange("G2:G" + sheet.getLastRow())
      .getValues()
      .flat()
      .map(item => String(item));

    // Lọc bỏ ô trống và chuyển về chữ thường
    const cleanData = data.filter(String).map(id => id.toString().trim());

    // Trả về chuỗi thuần túy: "gv100,gv101,admin22"
    return ContentService.createTextOutput(cleanData.join(",")).setMimeType(ContentService.MimeType.TEXT);
  }
  // Xác minh ADMIN VBA
  // Xác minh ADMIN VBA
  if (action === "getIdGVAD") {
    const sheet = ssAdmin.getSheetByName("idgv");
    // Lấy giá trị ô H2, ép kiểu chuỗi và cắt khoảng trắng
    const passAdmiN = supper(sheet.getRange("H2"));

    // Trả về trực tiếp chuỗi pass, không join gì cả
    return ContentService.createTextOutput(passAdmiN).setMimeType(ContentService.MimeType.TEXT);
  }
  // Ghi idinput vào sheet(danhsach)
  // Thêm vào trong hàm doGet(e) của thầy
// Ghi idinput vào sheet(danhsach)
  if (action === "getLastID") {  
    var ss2 = getSS2(e.parameter.sheetId, e.parameter.idgv || e.parameter.idnumber);
    var sheet = ss2.getSheetByName("danhsach"); 
    if (!sheet) return ContentService.createTextOutput("");
    var val = sheet.getRange("J2").getValue();
    return ContentService.createTextOutput(val.toString());
  }
  if (action === "normalize") {
    try {
      var result = normalizeQuestionBank();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        activeCount: result.activeCount,
        deletedCount: result.deletedCount
      })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (action === "saveLastID") {
    var idMoi = e.parameter.id;  
    var ss2 = getSS2(e.parameter.sheetId, e.parameter.idgv || e.parameter.idnumber);
    var sheet = ss2.getSheetByName("danhsach"); 
    if (sheet) {
      sheet.getRange("J2").setValue("'" + idMoi); 
      SpreadsheetApp.flush(); 
    }
    return ContentService.createTextOutput("Success");
  }
// Xác minh admin
  if (action === "checkAdminOTP") {
    var userOTP = e.parameter.otp;   
    var isCorrect = (supper(userOTP) === supper(passAdmin));
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      verified: isCorrect
    })).setMimeType(ContentService.MimeType.JSON);
  } 
  if (action === 'getSheetData') {    
    var sheet = ss.getSheetByName("dangcd");
    const lastRow = sheet.getLastRow();
    var range = sheet.getDataRange();
    const values = sheet.getRange(1, 1, lastRow, 10).getValues(); // Lấy toàn bộ hàng và cột
    
    return ContentService.createTextOutput(JSON.stringify(values))
      .setMimeType(ContentService.MimeType.JSON);
  } 
  // 6. XÁC MINH THÍ SINH
  // 2308sua1: Bắt cả type và action cho verifyStudent trong mainDoGet, hỗ trợ so khớp cả SBD/IDGV lẫn khóa H (sbd.idgv)
if (type === 'verifyStudent' || action === 'verifyStudent') {
  try {
    const idNumber = N9(params.idnumber || params.idgv || "");
    const sbd = supper(params.sbd || "");
    const pass = String(params.pass || "").trim();
    const reqSheetId = params.sheetId || "";

    // Bọc kiểm tra tham số bắt buộc từ client
    if (!sbd || !idNumber || !pass) {
      return createResponse("error", "Thiếu thông tin đăng nhập!");
    }

    const ss2 = getSS2(reqSheetId, idNumber);
    const sheet = ss2.getSheetByName("danhsach");
    if (!sheet) {
      return createResponse("error", "Không tìm thấy dữ liệu danh sách!");
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return createResponse("error", "Danh sách thí sinh trống!");
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const keyds = supper(sbd + "." + idNumber);

    for (let i = 0; i < data.length; i++) {
      const dbSbd = supper(data[i][0] || "");
      const dbIdNumber = N9(data[i][5] || "");
      const dbSbdKey = supper(data[i][7] || "");
      const dbPass = String(data[i][8] || "").trim();

      // Kiểm tra SBD & ID hoặc khóa ghép (Short-circuit evaluation)
      if ((dbSbd === sbd && (dbIdNumber === idNumber || supper(data[i][5] || "") === supper(params.idgv || ""))) || dbSbdKey === keyds) {
        
        // Ngăn chặn trường hợp Mật khẩu trong Sheet bị bỏ trống
        if (!dbPass) {
          return createResponse("error", "Tài khoản chưa được thiết lập mật khẩu!");
        }

        // Kiểm tra mật khẩu
        if (dbPass === pass) {
          const matchedSheetId = reqSheetId || getSheetIdByIdgv(idNumber);
          return createResponse("success", "OK", {
            name: data[i][1], 
            class: data[i][2], 
            limit: data[i][3],
            limittab: data[i][4], 
            taikhoanapp: data[i][6], 
            idnumber: idNumber, 
            sbd: "'" + sbd,
            sheetId: matchedSheetId
          });
        } else {
          return createResponse("error", "Mật khẩu không chính xác!");
        }
      }
    }

    // Chạy hết vòng lặp mà không tìm thấy
    return createResponse("error", "Số báo danh hoặc Số định danh không tồn tại!");
  } catch (error) {
    // Bắt toàn bộ lỗi phát sinh hệ thống, tránh sập app
    return createResponse("error", "Lỗi hệ thống: " + error.toString());
  }
}

// #02 Thi theo ma trận
// load ngân hàng đề
  if (action === "loadQuestions") {    
    const lastRow = sheetNH.getLastRow();
    if (lastRow < 2) {
    return createResponse("error", "Ngân hàng trống!");
      }  
    const values = sheetNH.getRange(2, 1, lastRow - 1, 8).getValues();
    // var headers = values[0]; // có cần lệnh này không?
    // var rows = values.slice(1);
    var result = rows.map(function (r) {
      var obj = {
        id: r[0],
        classTag: r[1],
        type: r[2],
        part: r[3],
        question: r[4]
      };

      if (r[2] === "mcq") {
        obj.o = r[5] ? JSON.parse(r[5]) : [];
        obj.a = r[6];
      }

      if (r[2] === "true-false") {
        obj.s = r[5] ? JSON.parse(r[5]) : [];
      }

      if (r[2] === "short-answer") {
        obj.a = r[6];
      }

      return obj;
    });

    return createResponse("success", "Load thành công", result);
  }
  //=========== Tìm lời giải ========================
if (action === 'getLG') {
  try {
    const idTraCuu = supper(params.id || "");
    if (!idTraCuu) return createResponse("error", "Thiếu ID tra cứu!");

    const lastRow = sheetNH.getLastRow();
    if (lastRow < 2) return createResponse("error", "Ngân hàng đang trống!");

    const data = sheetNH.getRange(2, 1, lastRow - 1, 8).getValues();
    const randomVersion = Math.floor(Math.random() * 9000) + 1000;

    for (let i = 0; i < data.length; i++) {
      const dbId = (data[i][0] || "").toString().trim();

      if (dbId === idTraCuu) {
        // Ép kiểu String() ngay từ đầu giúp hàm .replace() luôn an toàn 100%
        let qquestion = String(data[i][4] || "").replace(/\.png(?=['"]|\s|>|$)/g, ".png?v=" + randomVersion);
        let qoption   = String(data[i][5] || "");
        let qanswer   = String(data[i][6] || "");
        let qloigiai  = String(data[i][7] || "").replace(/\.png(?=['"]|\s|>|$)/g, ".png?v=" + randomVersion);

        const resultObj = {
          question: qquestion,
          option: qoption,
          answer: qanswer,
          loigiai: qloigiai
        };

        return ContentService.createTextOutput(JSON.stringify(resultObj))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    return createResponse("error", "Không tìm thấy ID câu hỏi này!");

  } catch (error) {
    return createResponse("error", "Lỗi xử lý getLG: " + error.toString());
  }
}
  // Tìm câu trùng
 if (action === 'findDuplicateQuestions') {
  const targetTag = e.parameter.targetClassTag;
  const res = findDuplicateQuestions(targetTag);
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}
  
  if (action == 'deleteQuestionRow') {
    var rowIdx = e.parameter.rowIdx;
    return ContentService.createTextOutput(JSON.stringify(deleteQuestionRow(rowIdx)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 7. LẤY CÂU HỎI THEO ID
  if (action === 'getQuestionById') {
    var id = supper(params.id);   
    const lastRow = sheetNH.getLastRow();
    if (lastRow < 2) {
    return createResponse("error", "Ngân hàng trống!");
      }  

    var dataNH = sheetNH.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < dataNH.length; i++) {
      if (dataNH[i][0].toString().trim() === id) {
        var answersheet = dataNH[i][6].toString().trim() || "";
        var typesheet = dataNH[i][2].toString().trim().toLowCase();
        var optionsheet = dataNH[i][5] || "";
        if (typesheet === "true-false") {
          answersheet = parseTfOptions(optionsheet);          
        }
        return createResponse("success", "OK", {
          idquestion: dataNH[i][0], 
          classTag: dataNH[i][1], 
          type: typesheet,
          question: dataNH[i][4],
          options: optionsheet,
          answer: answersheet,
          loigiai: dataNH[i][7],
          datetime: dataNH[i][8]
          
        });
      }
    }
    return resJSON({ status: 'error' });
  }

  // 8. LẤY MA TRẬN ĐỀ
  // 2308sua1: Hỗ trợ cả type và action cho getExamCodes, so khớp teacherId linh hoạt
  if (type === 'getExamCodes' || action === 'getExamCodes') {
  const teacherId = supper(params.idnumber || params.idgv || "");
  const ss2 = getSS2(params.sheetId, teacherId);
  const sheet = ss2.getSheetByName("matran");
  if (!sheet) {
    return createResponse("error", "Ma trận trống!");
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return createResponse("error", "Ma trận trống!");
  }  

  // 2208sua1: Quét 23 cột (từ cột A đến cột W - lớp áp dụng)
  const data = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
  const results = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // Kiểm tra ID giáo viên hoặc tài khoản hệ thống
    if ((row[0] || "").toString().trim() === teacherId || row[0].toString() === "SYSTEM") {
      try {
        // 🔥 ĐỌC GIÁ TRỊ NGÀY GIỜ TỪ CỘT T VÀ U (Chỉ số mảng là 19 và 20)
        const openDateVal = row[19] || ""; 
        const closeDateVal = row[20] || "";
        // 2208them1: Đọc lớp từ cột W (chỉ số mảng 22)
        const targetClass = (row[22] || "").toString().trim();

        // 🔥 CHẶN THỜI GIAN THEO HÀM opencloseDate CỦA ANH
        const isPastOpen = opencloseDate(openDateVal, 'open');
        const isPastClose = opencloseDate(closeDateVal, 'close');

        // Điều kiện: Đã đến giờ mở đề VÀ Chưa vượt quá giờ đóng đề thì mới cho hiện mã đề
        if (isPastOpen && !isPastClose) {
          results.push({
            code: row[1].toString(), 
            name: row[2].toString(), 
            topics: JSON.parse(row[3]),
            targetClass: targetClass, // 2208them1: Lớp dành cho mã đề
            fixedConfig: {
              duration: parseInt(row[4]), 
              numMC: JSON.parse(row[5]), 
              scoreMC: parseFloat(row[6]),
              mcL3: JSON.parse(row[7]), 
              mcL4: JSON.parse(row[8]), 
              numTF: JSON.parse(row[9]),
              scoreTF: parseFloat(row[10]), 
              tfL3: JSON.parse(row[11]), 
              tfL4: JSON.parse(row[12]),
              numSA: JSON.parse(row[13]), 
              scoreSA: parseFloat(row[14]), 
              saL3: JSON.parse(row[15]), 
              saL4: JSON.parse(row[16])
            }
          });
        }
      } catch (err) {
        // Bỏ qua dòng lỗi cấu trúc JSON để vòng lặp tiếp tục chạy
      }
    }
  }
  return createResponse("success", "OK", results);
}
  // 9. LẤY TẤT CẢ CÂU HỎI 
  if (action === "getQuestions") {  
  var rows = sheetNH.getDataRange().getValues();
  var questions = [];

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;

    var parsedOptions = null;
    try {
      parsedOptions = rows[i][5] ? JSON.parse(rows[i][5]) : null;
    } catch(e) {
      parsedOptions = null;
    }
    var qText = String(rows[i][4] || "");
    var qloigiai = String(rows[i][7] || "");
    var randomVersion = Math.floor(Math.random() * 9000) + 1000;

    if (qText.indexOf(".png'") !== -1) {
    qText = qText.replaceAll(".png'", ".png?v=" + randomVersion + "'");
    }
    if (qloigiai.indexOf(".png'") !== -1) {
    qloigiai = qloigiai.replaceAll(".png'", ".png?v=" + randomVersion + "'");
    }
    var qObj = {
      id: rows[i][0],
      classTag: rows[i][1] || "",
      type: rows[i][2] || "",
      part: rows[i][3] || "",
      question: qText,
      a: rows[i][6] || "",
      loigiai: qloigiai
    };

    if (qObj.type === "mcq") {
      qObj.o = parsedOptions;
    }

    if (qObj.type === "true-false") {
      qObj.s = parsedOptions;
    }

    if (qObj.type === "short-answer") {
      // không cần options
    }

    questions.push(qObj);
  }

  return createResponse("success", "OK", questions);
}

// #03 thi lẻ
//= TÌM CÂU HỎI LẺ=======
  if (action === "getSingleQuestion") {
  const idgv = N9(e.parameter.idgv || "");
  const ss2 = getSS2(e.parameter.sheetId, idgv);
  const sheet = ss2.getSheetByName("exam_data");
  if (!sheet) return createResponse("error", "Ngân hàng câu hỏi trống!");
  const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
    return createResponse("error", "Ngân hàng câu hỏi trống!");
      }  
  const examCodeInput = supper(e.parameter.examCode || "");
  const questionIdInput = supper(e.parameter.questionId || "");

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (let i = 0; i < data.length; i++) {    
    if (supper(data[i][0] || "") === examCodeInput && supper(data[i][1] || "") === questionIdInput && N9(data[i][7] || "") === idgv) {

  return createResponse(
    "success",
    "OK",
    {
      id: data[i][1],
      classTag: data[i][2],
      type: data[i][3],
      question: data[i][4],
      loigiai: data[i][5]
    }
  );

}
  }

  return createResponse("error", "Không tìm thấy câu hỏi");
} 
  // lấy ngân hàng theo idgv
  if (action === 'getQuestionsByCode') {
    const examCode = supper(params.examCode || "");
    const idgv = N9(params.idgv || "");
    const ss2 = getSS2(params.sheetId, idgv);
    const sheet = ss2.getSheetByName("exam_data");
    if (!sheet) return createResponse("error", "Chưa có dữ liệu exam_data");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
    return createResponse("error", "Ngân hàng trống!");
      }  
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const results = [];

    for (let i = 0; i < data.length; i++) {      
      if (supper(data[i][0] || "") === examCode && N9(data[i][7] || "") === idgv) {
        try {
          var qText = String(data[i][4] || "");
          var randomVersion = Math.floor(Math.random() * 9000) + 1000;
          if (qText.indexOf(".png'") !== -1) {
          qText = qText.replaceAll(".png'", ".png?v=" + randomVersion + "'");
          }
          results.push(JSON.parse(qText));
        } catch (err) {
          results.push(qText);
        }
      }
    }
    return createResponse("success", "OK", results);
  }

// #04 chung
  // Tải điểm
if (action === "downloadScores") {
    const idgv = (e.parameter.idgv || ""); 
    const exams = (e.parameter.exams || "");     
    const ss2 = getSS2(e.parameter.sheetId, idgv);
    const sheet = ss2.getSheetByName("ketqua");
    
    const keycheck = supper(exams + "." + idgv);
    
    if (!sheet) return ContentService.createTextOutput("Sheet ketqua không tồn tại").setMimeType(ContentService.MimeType.TEXT);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return ContentService.createTextOutput("Dữ liệu kết quả đang trống").setMimeType(ContentService.MimeType.TEXT);

    // --- ĐIỀU CHỈNH TẠI ĐÂY ---
    // Lấy header từ cột 1 đến cột 8 (Cột H)
    const header = sheet.getRange(1, 1, 1, 11).getValues()[0];
    
    // Lấy dữ liệu từ dòng 2, cột 1, đến dòng cuối, lấy 10 cột để vẫn lọc được cột I (cột 9)
    // Nhưng chúng ta sẽ dùng slice để cắt bớt trước khi trả về
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();   
    
    const filteredData = data
      .filter(row => supper(row[9] || "otrong") === keycheck) // Lọc dựa trên cột I (index 9)
      .map(row => row.slice(0, 9)); // Cắt bỏ cột J, chỉ lấy từ cột A (index 0) đến I (index 8)
    
    return ContentService.createTextOutput(JSON.stringify({
      header: header,
      data: filteredData
    })).setMimeType(ContentService.MimeType.JSON);
}
  
// ===== LẤY LIST EXAMS =====
  if (action === "getExamsList") {
    return getExamsList(e.parameter.type, e.parameter.idgv, e.parameter.sheetId);
  }

  // ===== RESET DATA =====
  if (action === "resetData") {
    const key = supper(e.parameter.password + "." + e.parameter.idgv );
      const sheetIdGV = ssAdmin.getSheetByName("idgv");
      const datapass = sheetIdGV.getRange("F2:F" + sheetIdGV.getLastRow()).getValues();
      let kiemtra = 0;
      for (let i = 0; i < datapass.length; i++) {

        if (datapass[i][0] && datapass[i][0].toString().trim() === key) {

        kiemtra = 1;

      break;
        }
      }
      if (kiemtra === 0) {
  return createResponse("error", "⚠️ Sai mật khẩu hoặc ID rồi thầy/cô ơi!");
}
    return resetData(
      e.parameter.type,
      e.parameter.password,
      e.parameter.mode,
      e.parameter.exams,
      e.parameter.idgv,
      e.parameter.sheetId
    );
  }

  // xem điểm
  if (action === "getScore") {
    return getScore(e);
  }
  // lấy dạng câu hỏi
  if (action === 'getAppConfig') {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: getAppConfig()
    })).setMimeType(ContentService.MimeType.JSON);
  }
// THÊM NHÁNH NÀY CHO MA TRẬN
if (action === 'getAppConfigmt') {
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: getAppConfigmt()
  })).setMimeType(ContentService.MimeType.JSON);
}

// 2408sua2: Route lấy danh sách mã đề chấm lại & chấm lại (GET) hỗ trợ lọc theloai (Matrix, Word, PDF)
if (action === "getRegradeExamsList") {
  const targetIdgv = e.parameter.idgv || e.parameter.idnumber || "";
  const reqTheloai = e.parameter.theloai || "";
  const sheetId2 = getSheetIdByIdgv(targetIdgv);
  const reqSheetId = sheetId2 || e.parameter.sheetId || "";
  return getRegradeExamsList(targetIdgv, reqSheetId, reqTheloai);
}  

  // 4. THỐNG KÊ ĐÁNH GIÁ
  if (type === 'getStats') {
    const stats = { ratings: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 } };
    const sheetRate = ss.getSheetByName("danhgia");
    if (sheetRate) {
      const rateData = sheetRate.getDataRange().getValues();
      for (let i = 1; i < rateData.length; i++) {
        const star = parseInt(rateData[i][1]);
        if (star >= 1 && star <= 5) stats.ratings[star]++;
      }
    }
    return createResponse("success", "OK", stats);
  }

// Kết thúc Doget
  return createResponse("error", "Yêu cầu không hợp lệ");
} 


// Hết Doget ###


  
function mainDoPost(e) {
// #05 Xác minh
const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    //const data = JSON.parse(e.postData.contents || "{}");
    // Chấp nhận cả dữ liệu gửi dạng JSON hoặc gửi dạng Form Parameter thông thường
    var data = {};
    if (e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(c) { data = e.parameter || {}; }
    } else {
      data = e.parameter || {};
    }
    const action = (data.action || (e.parameter && e.parameter.action) || data.type || (e.parameter && e.parameter.type) || "").toString();

    // 2308them1: Bổ sung xử lý action "verifyStudent" trong mainDoPost để học sinh xác minh qua POST không bị lỗi "Không khớp lệnh nào!"
    if (action === "verifyStudent" || data.type === "verifyStudent") {
      try {
        const idNumber = N9(data.idnumber || data.idgv || "");
        const sbd = supper(data.sbd || "");
        const pass = String(data.pass || "").trim();
        const reqSheetId = data.sheetId || "";

        if (!sbd || !idNumber || !pass) {
          return createResponse("error", "Thiếu thông tin đăng nhập!");
        }

        const ss2 = getSS2(reqSheetId, idNumber);
        const sheet = ss2.getSheetByName("danhsach");
        if (!sheet) {
          return createResponse("error", "Không tìm thấy dữ liệu danh sách!");
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
          return createResponse("error", "Danh sách thí sinh trống!");
        }

        const dataRows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
        const keyds = supper(sbd + "." + idNumber);

        for (let i = 0; i < dataRows.length; i++) {
          const dbSbd = supper(dataRows[i][0] || "");
          const dbIdNumber = N9(dataRows[i][5] || "");
          const dbSbdKey = supper(dataRows[i][7] || "");
          const dbPass = String(dataRows[i][8] || "").trim();

          if ((dbSbd === sbd && (dbIdNumber === idNumber || supper(dataRows[i][5] || "") === supper(data.idgv || ""))) || dbSbdKey === keyds) {
            if (!dbPass) {
              return createResponse("error", "Tài khoản chưa được thiết lập mật khẩu!");
            }

            if (dbPass === pass) {
              const matchedSheetId = reqSheetId || getSheetIdByIdgv(idNumber);
              return createResponse("success", "OK", {
                name: dataRows[i][1], 
                class: dataRows[i][2], 
                limit: dataRows[i][3],
                limittab: dataRows[i][4], 
                taikhoanapp: dataRows[i][6], 
                idnumber: idNumber, 
                sbd: "'" + sbd,
                sheetId: matchedSheetId
              });
            } else {
              return createResponse("error", "Mật khẩu không chính xác!");
            }
          }
        }
        return createResponse("error", "Số báo danh hoặc ID Giáo viên không chính xác!");
      } catch (err) {
        return createResponse("error", "Lỗi xác minh học sinh: " + err.toString());
      }
    }

    // 2308them1: Bổ sung xử lý action "getExamCodes" trong mainDoPost
    if (action === "getExamCodes" || data.type === "getExamCodes") {
      try {
        const teacherId = supper(data.idnumber || data.idgv || "");
        const ss2 = getSS2(data.sheetId, teacherId);
        const sheet = ss2.getSheetByName("matran");
        if (!sheet) {
          return createResponse("error", "Ma trận trống!");
        }
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
          return createResponse("error", "Ma trận trống!");
        }
        const mdata = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
        const results = [];
        for (let i = 0; i < mdata.length; i++) {
          const row = mdata[i];
          const rowTId = (row[0] || "").toString().trim();
          if (rowTId === teacherId || supper(rowTId) === teacherId || N9(rowTId) === N9(teacherId) || rowTId === "SYSTEM") {
            try {
              const openDateVal = row[19] || "";
              const closeDateVal = row[20] || "";
              const isPastOpen = opencloseDate(openDateVal, 'open');
              const isPastClose = opencloseDate(closeDateVal, 'close');
              if (isPastOpen && !isPastClose) {
                results.push({
                  code: row[1].toString(),
                  name: row[2].toString(),
                  topics: JSON.parse(row[3]),
                  fixedConfig: {
                    duration: parseInt(row[4]),
                    numMC: JSON.parse(row[5]),
                    scoreMC: parseFloat(row[6]),
                    mcL3: JSON.parse(row[7]),
                    mcL4: JSON.parse(row[8]),
                    numTF: JSON.parse(row[9]),
                    scoreTF: parseFloat(row[10]),
                    tfL3: JSON.parse(row[11]),
                    tfL4: JSON.parse(row[12]),
                    numSA: JSON.parse(row[13]),
                    scoreSA: parseFloat(row[14]),
                    saL3: JSON.parse(row[15]),
                    saL4: JSON.parse(row[16])
                  }
                });
              }
            } catch (err) {}
          }
        }
        return createResponse("success", "OK", results);
      } catch (err) {
        return createResponse("error", "Lỗi lấy mã đề: " + err.toString());
      }
    }

    // ACTION: Lưu đánh giá học sinh
    if (action === "submitRating" || action === "rate") {
      let sheetDG = ssAdmin.getSheetByName("danhgia");
      
      // Nếu chưa có sheet thì tự tạo và viết Header
      if (!sheetDG) {
        sheetDG = ssAdmin.insertSheet("danhgia");
        sheetDG.appendRow([
          "Timestamp", "stars", "name", "class", 
          "idNumber", "account", "comment", 
          "fullstars", "5stars", "4stars", "3stars", "2stars", "1stars"
        ]);
      }

      const timestamp = new Date();
      const stars = Number(data.stars) || 5;
      const name = data.name || "";
      const studentClass = data.class || "";
      const idNumber = data.idNumber || "";
      const account = data.account || "";
      const comment = data.comment || "";

      // Ghi một dòng mới vào Sheet
      sheetDG.appendRow([
        timestamp, stars, name, studentClass, 
        idNumber, account, comment, 
        "", "", "", "", "", ""
      ]);

      // Tính toán lại thống kê
      const values = sheetDG.getDataRange().getValues();
      let count5 = 0, count4 = 0, count3 = 0, count2 = 0, count1 = 0;
      let totalStars = 0;
      let totalReviews = 0;

      for (let i = 1; i < values.length; i++) {
        const rowStar = Number(values[i][1]);
        if (rowStar >= 1 && rowStar <= 5) {
          totalReviews++;
          totalStars += rowStar;
          if (rowStar === 5) count5++;
          if (rowStar === 4) count4++;
          if (rowStar === 3) count3++;
          if (rowStar === 2) count2++;
          if (rowStar === 1) count1++;
        }
      }

      const avgStars = totalReviews > 0 ? (totalStars / totalReviews).toFixed(1) : "5.0";

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        stats: {
          avgStars: avgStars,
          totalReviews: totalReviews,
          fiveStars: count5,
          fourStars: count4,
          threeStars: count3,
          twoStars: count2,
          oneStar: count1
        }
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2408sua2: Route chấm lại bài thi (POST)
    if (action === "regradeExams") {
      const reqIdgv = data.idgv || e.parameter.idgv || "";
      const reqPass = data.password || e.parameter.password || "";
      const reqExamCode = data.examCode || data.exams || e.parameter.examCode || e.parameter.exams || "";
      const reqTheloai = data.theloai || e.parameter.theloai || "";
      const sheetId2 = getSheetIdByIdgv(reqIdgv);
      const reqSheetId = sheetId2 || data.sheetId || e.parameter.sheetId || "";     
      return regradeExams(reqIdgv, reqPass, reqExamCode, reqSheetId, reqTheloai);
    }
    const res = (status, message, payload) =>
      ContentService.createTextOutput(
        JSON.stringify({ status, message, data: payload || null })
      ).setMimeType(ContentService.MimeType.JSON);
    //2308sua3: Xử lý lưu ảnh an toàn
    // --- CHÈN NHÁNH XỬ LÝ LƯU ẢNH VÀO ĐÂY ---
    if (action === "uploadImage") {
      var base64Data = data.fileData; 
      var fileName = data.fileName;   
      
      if (!base64Data || !fileName) {
        return res("error", "GAS không nhận được dữ liệu fileData hoặc fileName!");
      }
      
      var folderId = "1Gk_9n0JWveBlwXQDlqwVTpSceYqv_WNI";
      var folder = DriveApp.getFolderById(folderId);
      
      // Khử các khoảng trắng sinh ra do quá trình truyền chuỗi
      base64Data = base64Data.replace(/ /g, '+');
      
      var decodedData = Utilities.base64Decode(base64Data);
      var blob = Utilities.newBlob(decodedData, 'image/png', fileName); 
      var file = folder.createFile(blob);
      
      return res("success", "Đã lưu thành công file: " + fileName, { "fileUrl": file.getUrl() });
    }

    //2308them3: Bắt toàn bộ các action submit đề thi (Word và Ma Trận), khắc phục hoàn toàn lỗi sheetKq is not defined
    if (action === "submitExam" || action === "submitExamMatrix" || action === "submitMatrix" || action === "submit" || (data && (data.type === "submitExam" || data.type === "submitExamMatrix" || data.type === "submitMatrix"))) {
      try {
        // 1. LẤY SHEET TẢI ĐIỂM (Mở Sheet riêng của GV nếu có, nếu không lấy Sheet hiện tại)
        var idgv2 = (data.idgv || data.idnumber || "").toString().trim();
        var reqSheetId = (data.sheetId || "").toString().trim();
        var sheetId2 = (typeof getSheetIdByIdgv === "function" && idgv2) ? getSheetIdByIdgv(idgv2) : "";
        var finalSheetId = sheetId2 || reqSheetId;
        var targetSS = (typeof getSS2 === "function") ? getSS2(finalSheetId, idgv2) : null;
        
        if (!targetSS) {
          targetSS = (typeof ss !== "undefined" && ss) ? ss : SpreadsheetApp.getActiveSpreadsheet(); // Fallback an toàn
        }

        var sheetKq = targetSS.getSheetByName("ketqua") || targetSS.getSheetByName("KetQua");    
        if (!sheetKq) {
          sheetKq = targetSS.insertSheet("ketqua");
        }
        
        // Đảm bảo tiêu đề cột chuẩn nếu sheet mới tạo hoặc trống
        if (sheetKq.getLastRow() === 0) {
          sheetKq.appendRow(["Timestamp", "Mã đề", "SBD", "Họ tên", "Lớp", "Tổng điểm", "Thời gian làm", "IDGV", "Vi phạm", "exams.idgv", "exams.sbd.idgv", "Thể loại", "Detail", "Nhận xét", "Ghi chú", "Điểm cũ"]);
        }

        // 2. CHUẨN HÓA DỮ LIỆU
        var exams = (data.exams || data.examCode || "").toString().toUpperCase();
        var idgv = idgv2;
        var rawScore = data.tongdiem !== undefined ? data.tongdiem : (data.score !== undefined ? data.score : 0);
        var diem = rawScore;
        var className = (data.class || data.className || "Tự do").toString();
        var thoiGian = (data.time !== undefined && data.time !== null) ? data.time : (data.totalTime || 0);
        var sbd = (data.sbd || "").toString().trim();
        var tabCount = data.tabSwitches !== undefined ? data.tabSwitches : (data.vipham !== undefined ? data.vipham : (data.tabCount || 0));
        var theloai = data.theloai || (action === "submitExamMatrix" || action === "submitMatrix" ? "Matrix" : "Word");
        
        // Tính điểm dạng số để lấy nhận xét chuẩn xác
        var numericScore = typeof rawScore === 'number' ? rawScore : (parseFloat(String(rawScore).replace(',', '.')) || 0);
        var nx = (typeof layNhanXet === "function") ? layNhanXet(numericScore) : "Không có nhận xét nào";

        // Xử lý chuỗi chi tiết an toàn tránh vượt giới hạn ký tự cell Google Sheets
        // Xử lý detailsStr an toàn
        var detailsStr = data.details || "";
        if (typeof detailsStr === "object") {
          detailsStr = JSON.stringify(detailsStr);
          } else {
          detailsStr = String(detailsStr);
            }

        if (detailsStr.length > 45000) {
        detailsStr = detailsStr.substring(0, 45000) + "...";
      }

        // 3. TÌM HÀNG TRỐNG TIẾP THEO (Quét tìm ô trống thực tế tại Cột B - Mã đề)
        var vals = sheetKq.getDataRange().getValues();
        var nextRow = -1;
        for (var i = 1; i < vals.length; i++) {
          var cellValue = vals[i][1] !== undefined && vals[i][1] !== null ? String(vals[i][1]).trim() : "";
          if (cellValue === "") {
            nextRow = i + 1; 
            break;
          }
        }

        if (nextRow === -1) {
          nextRow = sheetKq.getLastRow() + 1;
        }

        if (nextRow < 2) nextRow = 2;

        // 4. CHUẨN BỊ MẢNG DỮ LIỆU 1 HÀNG
        var rowData = [
          data.timestamp || new Date().toLocaleString('vi-VN'), // A: Timestamp
          "'" + supper(exams),                                  // B: Mã đề
          "'" + supper(sbd),                                    // C: SBD
          supper(data.name || ""),                              // D: Họ tên
          supper(className),                                    // E: Lớp
          diem,                                                 // F: Tổng điểm
          thoiGian,                                             // G: Thời gian làm            
          "'" + supper(idgv),                                   // H: IDGV
          tabCount,                                             // I: Vi phạm
          "'" + supper(exams + "." + idgv),                     // J: exams.idgv
          "'" + supper(exams + "." + sbd + "." + idgv),         // K: exams.sbd.idgv
          theloai,                                              // L: Thể loại
          detailsStr,                                           // M: Detail
          nx || "Không có nhận xét nào"                         // N: Nhận xét
        ];

        // 5. GHI DỮ LIỆU VÀ ĐỊNH DẠNG
        sheetKq.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
        
        try { 
          sheetKq.getRange(nextRow, 13).setWrap(true); 
        } catch(eWrap) {}
       
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "success", 
          message: "Ghi điểm thành công!",
          rowRecorded: nextRow
        })).setMimeType(ContentService.MimeType.JSON);

      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ 
          status: "error", 
          message: "Lỗi ghi điểm: " + err.toString() 
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    //2308ketthucsua3

    if (data.type === 'register') {
      var sheetUser = ss.getSheetByName("users");
      sheetUser.appendRow([new Date(), data.phone, data.pass]);
      return ContentService.createTextOutput("Đã đăng ký thành công");
    }
// 4. XÁC MINH GIÁO VIÊN (verifyGV)
    if (action === "verifyGV") {
      var sheetGV = ssAdmin.getSheetByName("idgv");
      var rows = sheetGV.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][0].toString().trim() === data.idnumber.toString().trim() && rows[i][1].toString().trim() === data.password.toString().trim()) {
          var teacherSheetId = rows[i][9] ? rows[i][9].toString().trim() : "";
          return resJSON({ status: "success", sheetId: teacherSheetId });
        }
      }
      return resJSON({ status: "error", message: "ID hoặc Mật khẩu GV không đúng!" });
    }
// 6. XÁC MINH ADMIN (verifyAdmin)
    if (action === "verifyAdmin") {      
      if (supper(data.password) === suppere(passAdmin)) return resJSON({ status: "success", message: "Chào Admin!" });
      return resJSON({ status: "error", message: "Sai mật khẩu!" });
    }
// #06 Ma trận
// 1. NHÁNH LỜI GIẢI (saveLG)
    if (action === 'saveLG') {
  var lastRow = sheetNH.getLastRow();
  if (lastRow < 2) return ContentService.createTextOutput("⚠️ Sheet rỗng!").setMimeType(ContentService.MimeType.TEXT);

  var idValues = sheetNH.getRange(1, 1, lastRow, 1).getValues().map(function(r) { 
    return r[0].toString().trim(); 
  });

  var count = 0;
  data.forEach(function (item) {
    var idToFind = item.id.toString().trim();
    
    var rowIndex = idValues.indexOf(idToFind);

    if (rowIndex !== -1) {
      var targetRow = rowIndex + 1;
      var rawLG = item.loigiai || "";

      sheetNH.getRange(targetRow, 8).setValue(rawLG);
      count++;
    }
  });

  sheetNH.getRange("H:H").setWrap(true);
  return ContentService.createTextOutput("🚀 Thành công! Đã cập nhật " + count + " lời giải vào đúng hàng theo ID.");
}
    // 2. NHÁNH MA TRẬN (saveMatrix)
    if (action === "saveMatrix") {
      var now = new Date();      
      const gvId = data.gvId || data.idgv || "";
      const reqSheetId = data.sheetId || "";
      const ss2 = getSS2(reqSheetId, gvId);

      const sheetMatran = ss2.getSheetByName("matran") || ss2.insertSheet("matran");
      const toStr = (v) => (v != null) ? String(v).trim() : "";
      const toNum = (v) => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
      const toJson = (v) => {
        if (!v || v === "" || (Array.isArray(v) && v.length === 0)) return "[]";
        if (typeof v === 'object') return JSON.stringify(v);
        let s = String(v).trim();
        return s.startsWith("[") ? s : "[" + s + "]";
      };
      sheetMatran.getRange("A:A").setNumberFormat("@");
      const rowData = [
        "'" + supper(toStr(data.gvId)), 
        "'" + supper(toStr(data.makiemtra)), 
        toStr(data.name), 
        toJson(data.topics),
        toNum(data.duration), 
        toJson(data.numMC), 
        toNum(data.scoreMC), 
        toJson(data.mcL3),
        toJson(data.mcL4), 
        toJson(data.numTF), 
        toNum(data.scoreTF), 
        toJson(data.tfL3),
        toJson(data.tfL4), 
        toJson(data.numSA), 
        toNum(data.scoreSA), 
        toJson(data.saL3), 
        toJson(data.saL4),
         "'" + supper(toStr(data.gvId)), 
        "'" + supper(toStr(data.makiemtra) + "." + toStr(data.gvId)),
        "'" + toStr(data.openDate),
        "'" + toStr(data.closeDate),
        now,
        "'" + supper(toStr(data.lop || data.class || data.targetClass || "")),
        data.thuong || ""
      ];
      const key = supper(data.gvPass + "." + data.gvId);
      const sheetId = ssAdmin.getSheetByName("idgv");
      const datapass = sheetId.getRange("F2:F" + sheetId.getLastRow()).getValues();
      let kiemtra = 0;
      for (let i = 0; i < datapass.length; i++) {

        if (datapass[i][0] && datapass[i][0].toString().trim() === key) {

        kiemtra = 1;

      break;
        }
      }
      if (kiemtra === 0) {
  return createResponse("error", "⚠️ Sai mật khẩu hoặc ID rồi thầy/cô ơi!");
}
      const vals = sheetMatran.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < vals.length; i++) {
        if (supper(vals[i][0]) === supper(toStr(data.gvId)) && supper(vals[i][1]) === supper(toStr(data.makiemtra))) {
          rowIndex = i + 1; break;
        }
      }
      if (rowIndex > 0) { sheetMatran.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]); }
      else { sheetMatran.appendRow(rowData); }
      return createResponse("success", "✅ Đã tạo ma trận " + data.makiemtra + " thành công!");
    }

    // 3. NHÁNH LƯU CÂU HỎI MỚI (saveQuestions)
    if (action === 'saveQuestions') {
  var now = new Date();
  var lastRow = sheetNH.getLastRow();
  
  // 1. Lấy ID cuối cùng trong Sheet (ép kiểu số)
  var lastIdInSheet = 0;
  if (lastRow > 0) {
    lastIdInSheet = Number(sheetNH.getRange(lastRow, 1).getValue()) || 0;
  }

  // 2. Chỉ kiểm tra item đầu tiên của mảng data gửi lên
  if (data.length > 0) {
    var firstItemId = Number(data[0].id);

    // Nếu ID đầu tiên nhỏ hơn hoặc bằng ID cuối trong sheet -> Chặn luôn
    if (firstItemId <= lastIdInSheet) {
      return createResponse("error", "Dữ liệu đã tồn tại hoặc ID không hợp lệ (ID đầu tiên " + firstItemId + " không lớn hơn " + lastIdInSheet + ")");
    }
  } else {
    return createResponse("error", "Không có dữ liệu để lưu");
  }

  // 3. Nếu vượt qua kiểm tra, tiến hành map và lưu toàn bộ
  var rows = data.map(function (item) {
  // Lấy giá trị answer thô, ép về chuỗi an toàn
  let rawAnswer = item.answer != null ? String(item.answer) : "";

  // Nếu là câu trả lời ngắn: đổi tất cả dấu phẩy thành dấu chấm và xoá khoảng trắng thừa
  if (item.type === "short-answer") {
    rawAnswer = rawAnswer.replace(/,/g, '.').trim();
  }
  return [
    item.id,
    item.classTag,
    item.type,
    item.part,
    item.question,
    item.options || "",
    rawAnswer,
    item.loigiai || "",
    now
  ];
});

  sheetNH.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  sheetNH.getRange("D:H").setWrap(true);

  return createResponse("success", "Đã lưu thành công " + rows.length + " câu hỏi!");
}
    // 5. CẬP NHẬT CÂU HỎI (updateQuestion)
    if (action === 'updateQuestion') {
  var item = data.data;
  Logger.log(e.postData.contents); 
  var allRows = sheetNH.getDataRange().getValues();
  
  // Kiểm tra ID từ client gửi lên có bị trống không
  var targetId = item.id || item.idquestion;
  if (!targetId) return resJSON({ status: 'error', message: 'ID gửi lên bị trống!' });

  for (var i = 1; i < allRows.length; i++) {
    // CHỐT CHẶN: Nếu ô ID trong Sheet bị trống thì bỏ qua, không so sánh
    if (allRows[i][0] === "" || allRows[i][0] === null || typeof allRows[i][0] === 'undefined') {
      continue; 
    }

    // So sánh an toàn sau khi đã chắc chắn ô đó có dữ liệu
    if (allRows[i][0].toString() === targetId.toString()) {
      // Ghi dữ liệu vào các cột tương ứng (Cột 2: classTag, 5: Question...)
      sheetNH.getRange(i + 1, 2).setValue(item.classTag || "");
      sheetNH.getRange(i + 1, 5).setValue(item.question || "");
      sheetNH.getRange(i + 1, 6).setValue(item.options || "");
      sheetNH.getRange(i + 1, 7).setValue(item.answer || "");
      sheetNH.getRange(i + 1, 8).setValue(item.loigiai || "");
      sheetNH.getRange(i + 1, 9).setValue(new Date().toLocaleString('vi-VN'));

      return resJSON({ status: 'success' });
    }
  }
  return resJSON({ status: 'error', message: 'Không tìm thấy ID: ' + targetId });
}

// #07 Thi lẻ và PDF
// Lấy link thi PDF
// #07 Thi lẻ và PDF
    // =================================================

    if (action === "getExamLink") {
  const idgv = (data.idgv || "").toString().replace(/'/g, "").trim().toUpperCase();
  const maDe = (data.maDe || "").toString().trim().toUpperCase();
  const sbd  = (data.sbd || "").toString().trim().toUpperCase();
  const pass = (data.password || "").toString().trim();
  const reqSheetId = data.sheetId || "";

  if (!idgv || !maDe || !pass || !sbd) {
    return resJSON({
      status: "error",
      message: "Thiếu Số báo danh, IDGV, mã đề hoặc mật khẩu!"
    });
  }
    
  const isUserValid = verifyhocsinh(sbd, idgv, pass, reqSheetId);
  if (!isUserValid) {
    return resJSON({ 
      status: "fail", 
      message: "Sai Số báo danh, Mã GV hoặc Mật khẩu học sinh!" 
    });
  }

  const examLinkdn = verifyExams(maDe, idgv, reqSheetId);
  if (examLinkdn === false) {
    return resJSON({ 
      status: "fail", 
      message: "Mã đề thi không hợp lệ hoặc không thuộc giáo viên này!" 
    });
  }
  const examLink = examLinkdn + "&sbd=" + sbd + "&pass=" + pass;
  if (!examLinkdn || examLinkdn.toString().trim() === "") {
    return resJSON({
      status: "error",
      message: "Kỳ thi hợp lệ nhưng Giáo viên chưa cấu hình link đề thi!"
    });
  }

  return resJSON({
    status: "success",
    message: "Đã tìm thấy link!",
    data: {
      link: examLink
    }
  });
}

    // 2308sua1: Sửa studentGetExam để so khớp linh hoạt SBD, IDGV và mã đề theo cả khóa ghép H/O lẫn từng cột A/F
    if (action === "studentGetExam") {
      try {
        const sbd = data.sbd ? data.sbd : "";
        const pass = data.pass ? data.pass : "";
        const examCode = data.examCode ? data.examCode : "";
        const idgv = data.idgv ? data.idgv : "";
        const reqSheetId = data.sheetId || "";
        const keyds = supper(sbd + "." + idgv);
        const keyexams = supper(examCode + "." + idgv);
        const keysbd = supper(examCode + "." + sbd + "." + idgv);

        const ss2 = getSS2(reqSheetId, idgv);

        const sheetDS = ss2.getSheetByName("danhsach");
        const sheetData = ss2.getSheetByName("exam_data");
        const sheetExam = ss2.getSheetByName("exams");
        const sheetKQ = ss2.getSheetByName("ketqua");
        if (!sheetDS) return createResponse("error", "Không tìm thấy sheet danhsach!");
        const dataDS = sheetDS.getDataRange().getValues();        
        if (dataDS.length < 2) {
          return createResponse("error", "Danh sách thí sinh trống!");
      }    

var student = null;
for (var i = 1; i < dataDS.length; i++) {
  var rowSBD = supper(dataDS[i][7] || "");  
  var colSbd = supper(dataDS[i][0] || "");
  var colIdgv = N9(dataDS[i][5] || "");
  var rowPass = (dataDS[i][8] || "").toString().trim();
  
  if (rowPass === pass.toString().trim() && (rowSBD === keyds || (colSbd === supper(sbd) && (colIdgv === N9(idgv) || supper(dataDS[i][5] || "") === supper(idgv))))) {
    student = dataDS[i];
    break;
  }
}

if (!student) {
  return createResponse("error", "SBD hoặc Mật khẩu học sinh không chính xác!");
}
if (!sheetExam) return createResponseW("error", "Không tìm thấy sheet exams!");
const exRow = sheetExam.getDataRange().getValues().find(r => 
  supper(r[14] || "") === keyexams || (supper(r[0] || "") === supper(examCode) && (N9(r[1] || "") === N9(idgv) || supper(r[1] || "") === supper(idgv)))
);
        if (!exRow) return createResponseW("error", "Không tìm thấy mã đề: " + examCode + " / GV: " + idgv);
const now = new Date();

const openTime = exRow[12] instanceof Date 
  ? exRow[12] 
  : new Date(exRow[12]);

const closeTime = exRow[11] instanceof Date 
  ? exRow[11] 
  : new Date(exRow[11]);

        const maxAttempts = parseInt(exRow[13], 10) || 1;
        let exRowKq = [];

        if (sheetKQ && sheetKQ.getLastRow() > 1) {
        exRowKq = sheetKQ.getRange(2, 1, sheetKQ.getLastRow()-1, 11).getValues();
          }
        const currentAttempts = exRowKq.filter(r => 
      r[10].toString().trim() === keysbd).length;

    if (sbd !== "8888") {       
      if (openTime && now < openTime) {
  return createResponseW("error", 
    "⏳ Bài thi chưa mở. Thời gian mở: " +
    Utilities.formatDate(openTime, "GMT+7", "yyyy/MM/dd HH:mm")
  );
}     

if (closeTime && now > closeTime) {
  return createResponseW("error", 
    "⛔ Bài thi đã đóng lúc: " +
    Utilities.formatDate(closeTime, "GMT+7", "yyyy/MM/dd HH:mm")
  );
}
       if (currentAttempts >= maxAttempts) {
        return createResponseW("error", `Bạn đã hết lượt thi! Mã đề ${examCode} chỉ cho phép thi tối đa ${maxAttempts} lần.`);
      }
    }
        const toInt = (v, def = 0) => {
          const n = parseInt(v?.toString().trim(), 10);
          return isNaN(n) ? def : n;
        };

        const toFloat = (v, def = 0) => {
          if (v === null || v === undefined) return def;
          const s = v.toString().replace(",", ".");
          const n = parseFloat(s);
          return isNaN(n) ? def : n;
        };

        const toDateISO = (v) => {
          if (v instanceof Date) {
            return Utilities.formatDate(v, "GMT+7", "yyyy-MM-dd");
          }
          const s = v?.toString().trim();
          return s || "";
        };

        if (!sheetData) return createResponseW("error", "Không tìm thấy sheet exam_data!");
        const allRows = sheetData.getDataRange().getValues();
        const filteredQuestions = allRows.slice(1)
          .filter(r => supper(r[8] || "") === keyexams || (supper(r[0] || "") === supper(examCode) && (N9(r[7] || "") === N9(idgv) || supper(r[7] || "") === supper(idgv))))
          .map(r => {
            let raw = r[4];
            if (!raw) return null;

                let contentStr = raw.toString().trim();
                    try {
                return JSON.parse(contentStr);
                  } catch (e) {
                      try {
                       let fixed = contentStr.replace(/\\/g, "\\\\").replace(/\\\\"/g, "\\\"");
                      return JSON.parse(fixed);
                        } catch (e2) {
                      return {
                    type: "mcq",
                      question: contentStr,
                    id: r[1],
              error: "Lỗi định dạng JSON"
    };
  }
}
          })
          .filter(Boolean);

        return createResponseW("success", "OK", {
          studentName: student[1],
          studentClass: student[2],
          duration: toInt(exRow[8], 33),
          minSubmitTime: toInt(exRow[9], 0),
          maxTabSwitches: toInt(exRow[10], 3),
          maxthi: maxAttempts,
          deadline: Utilities.formatDate(closeTime, "GMT+7", "yyyy/MM/dd HH:mm"),
          openTime: Utilities.formatDate(openTime, "GMT+7", "yyyy/MM/dd HH:mm"),
          scoreMCQ: toFloat(exRow[3], 0),
          scoreTF: toFloat(exRow[5], 0),
          scoreSA: toFloat(exRow[7], 0),

          questions: filteredQuestions
        });

      } catch (error) {
        return createResponseW("error", "Lỗi GAS: " + error.toString());
      }
    }
    if (action === 'saveOnlySolutions') {
      const examCode = data.examCode;
      const idgv = data.idgv;
      const reqSheetId = data.sheetId || "";
      const ss2 = getSS2(reqSheetId, idgv);

      const sheet = ss2.getSheetByName("exam_data");
      if (!sheet) return createResponse("error", "Không tìm thấy sheet!");

      const lastRow = sheet.getLastRow();
      const solutions = data.solutions;

      const range = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      let updatedCount = 0;

      solutions.forEach(solText => {
        const idMatch = solText.match(/id\s*:\s*"?([\w.]+)"?/);
        let found = false;

        if (idMatch) {
          const solId = idMatch[1].toString();
          const key = supper(examCode + "." + solId + "." + idgv)
          for (let i = 1; i < range.length; i++) {
           
            if (range[i][9].toString() === key) {
              sheet.getRange(i + 1, 6).setValue(solText);
              range[i][5] = solText;
              updatedCount++;
              found = true;
              break;
            }
          }
        }

        if (!found) {
          for (let i = 1; i < range.length; i++) {
            if (range[i][0].toString() === examCode.toString() && (!range[i][5] || range[i][5].toString().trim() === "")) {
              sheet.getRange(i + 1, 6).setValue(solText);
              range[i][5] = solText;
              updatedCount++;
              found = true;
              break;
            }
          }
        }
      });
      sheet.getRange("D:H").setWrap(true);

      return createResponse("success", `Đã nạp xong ${updatedCount} lời giải cho mã ${examCode}!`);
    }

    if (action === "saveOnlyQuestions") {
  const examCode = data.examCode;
  const idgv = data.idgv;
  const reqSheetId = data.sheetId || "";
  const ss2 = getSS2(reqSheetId, idgv);

  const sheet = ss2.getSheetByName("exam_data") || ss2.insertSheet("exam_data");
  const qArray = data.questions;
  const force = data.force || false; 
  
  if (!Array.isArray(qArray)) return createResponse("error", "questions không phải mảng!");

  const fullData = sheet.getDataRange().getValues();

  if (qArray.length === 1 && !force) {
    const targetId = qArray[0].id.toString();
    let rowIdx = -1;

    for (let i = 0; i < fullData.length; i++) {
      if ((fullData[i][0] || "").toString() === examCode.toString() && (fullData[i][1] || "").toString() === targetId) {
        rowIdx = i + 1;
        break;
      }
    }

    if (rowIdx !== -1) {
      const q = qArray[0];
      let finalLG = (q.loigiai && q.loigiai.trim() !== "") ? q.loigiai : "Đang cập nhật...";
      const rowToUpdate = [
        examCode, 
        q.id || "", 
        q.classTag || "1001.a", 
        q.type || "mcq", 
        q.question || "", 
        finalLG, 
        new Date(),
        "'" + idgv,
        examCode + "." + idgv,
        examCode + "." + q.id + "." + idgv
      ];
      sheet.getRange(rowIdx, 1, 1, 10).setValues([rowToUpdate]);
      return createResponse("success", `Đã cập nhật riêng câu ID: ${targetId}`);
    }
  }
  const exists = fullData.some(row => row[0].toString() === examCode.toString());
  if (exists && !force) return createResponse("exists", `Mã đề đã có dữ liệu!`);

  if (exists && force) {
    for (let i = fullData.length - 1; i >= 0; i--) {
      if ((fullData[i][0] || "").toString() === examCode.toString()) sheet.deleteRow(i + 1);
    }
  }
  const rows = qArray.map(q => [
    examCode, 
    q.id || "", 
    q.classTag || "1001.a", 
    q.type || "mcq", 
    q.question || "", 
    (q.loigiai && q.loigiai.trim() !== "") ? q.loigiai : "Đang cập nhật...", 
    new Date(),
     "'" + idgv,
    examCode + "." + idgv,
    examCode + "." + q.id + "." + idgv
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
  var lastRow = sheet.getLastRow();
      sheet.getRange("E:F").setWrap(true);

  return createResponse("success", `Đã nạp ${rows.length} câu vào mã ${examCode}`);
}

    if (action === "saveExamConfig") {
      const examCode = (data.examCode || "").toString().trim();
      const idgv = (data.idgv || "").toString().trim();
      const reqSheetId = data.sheetId || "";
      const ss2 = getSS2(reqSheetId, idgv);

      const sheetExamsGV = ss2.getSheetByName("exams") || ss2.insertSheet("exams");
      const passGV = (data.passGV || "").toString().trim();
      const cfg = data.config;
      const key = supper(passGV + "." + idgv);
      const keyId = supper(examCode + "." + idgv);
      const sheetId = ssAdmin.getSheetByName("idgv");
      const datapass = sheetId.getRange("F2:F" + sheetId.getLastRow()).getValues();
      let kiemtra = 0;
      for (let i = 0; i < datapass.length; i++) {

        if (datapass[i][0] && datapass[i][0].toString().trim() === key) {

        kiemtra = 1;

      break;
        }
      }
      if (kiemtra === 0) {
  return createResponse("error", "⚠️ Sai mật khẩu hoặc ID rồi thầy/cô ơi!");
}

      const isForce = data.force === true || data.force === "true";

      const vals = sheetExamsGV.getDataRange().getValues();
      let existingRow = -1;
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][14] && vals[i][14].toString().trim() === keyId) {
          existingRow = i + 1;
          break;
        }
      }

      if (existingRow !== -1 && !isForce) {
        return createResponse("exists", "Mã đề đã tồn tại!");
      }
        sheetExamsGV.getRange("B:B").setNumberFormat("@");
      const rowData = [
        supper(examCode), 
        "'" + supper(idgv), 
        cfg.numMCQ, 
        cfg.scoreMCQ, 
        cfg.numTF, 
        cfg.scoreTF,
        cfg.numSA, 
        cfg.scoreSA, 
        cfg.duration, 
        cfg.mintime, 
        cfg.tab, 
        cfg.close, 
        cfg.open, 
        cfg.maxthi,
        keyId, "", "", "", "", 
        "Word"
      ];
      
      if (existingRow !== -1) {
        sheetExamsGV.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
        return createResponse("success", "✅ Đã GHI ĐÈ cấu hình đề " + examCode);
      } else {
        sheetExamsGV.appendRow(rowData);
        return createResponse("success", "✅ Đã lưu cấu hình mới cho đề " + examCode);
      }
    }

    if (action === "uploadWord") {
      const sheetExams = ss.getSheetByName("Exams") || ss.insertSheet("Exams");
      const sheetBank = ss.getSheetByName("QuestionBank") || ss.insertSheet("QuestionBank");
      sheetExams.appendRow([data.config.title, data.idNumber, data.config.duration, data.config.minTime, data.config.tabLimit, JSON.stringify(data.config.points)]);
      data.questions.forEach(function (q) { sheetBank.appendRow([data.config.title, q.part, q.type, q.classTag, q.question, q.answer, q.image]); });
      return createResponse("success", "UPLOAD_DONE");
    }

if (action === "getListMade") {
  const idgv = (N9(data.idgv) || "").toString().trim();
  const reqSheetId = data.sheetId || "";
  const ss2 = getSS2(reqSheetId, idgv);

  if (!idgv) {
    return resJSON({ status: "error", message: "Vui lòng nhập IDGV!" });
  }

  let listResult = [];

  // 1. XỬ LÝ SHEET 'matran'
  const sheetMatran = ss2.getSheetByName("matran");
  if (sheetMatran) {
    const lastRowMatran = sheetMatran.getLastRow();
    if (lastRowMatran >= 2) {
      const dataMatran = sheetMatran.getRange(2, 1, lastRowMatran - 1, 2).getValues();
      for (let i = 0; i < dataMatran.length; i++) {
        const currentIdgv = (N9(dataMatran[i][0]) || "").toString().trim();
        const currentMaDe = (dataMatran[i][1] || "").toString().trim();
        
        if (currentIdgv === idgv && currentMaDe) {
          listResult.push({
            maDe: currentMaDe,
            theLoai: "Ma Trận"
          });
        }
      }
    }
  }

  // 2. XỬ LÝ SHEET 'exams'
  const sheetExams = ss2.getSheetByName("exams");
  if (sheetExams) {
    const lastRowExams = sheetExams.getLastRow();
    if (lastRowExams >= 2) {
      const dataExams = sheetExams.getRange(2, 1, lastRowExams - 1, 16).getValues();
      for (let i = 0; i < dataExams.length; i++) {
        const currentMaDe = (dataExams[i][0] || "").toString().trim();
        const currentIdgv = (N9(dataExams[i][1]) || "").toString().trim();
        const valueCotP = (dataExams[i][15] || "").toString().trim();

        if (currentIdgv === idgv && currentMaDe) {
          const loaiDe = valueCotP === "" ? "Word" : "PDF";
          listResult.push({
            maDe: currentMaDe,
            theLoai: loaiDe
          });
        }
      }
    }
  }

  // Trả về dữ liệu kết quả tổng hợp
  return resJSON({
    status: "success",
    message: "Lấy danh sách mã đề thành công!",
    data: listResult
  });
}
    // 8. NHÁNH THEO TYPE (quiz, rating, ketqua)
    if (data.type === 'rating') {
      let sheetRate = ss.getSheetByName("danhgia") || ss.insertSheet("danhgia");
      sheetRate.appendRow([new Date(), data.stars, data.name, data.class, data.idNumber, data.comment || "", data.taikhoanapp]);
      return createResponse("success", "Đã nhận đánh giá");
    }
    if (data.type === 'quiz') {
      let sheetQuiz = ss.getSheetByName("ketquaQuiZ") || ss.insertSheet("ketquaQuiZ");
      sheetQuiz.appendRow([new Date(), data.examCode || "QUIZ", data.name || "N/A", data.class || "", data.school || "", data.phoneNumber || "", data.score || 0, data.totalTime || "00:00", data.stk || "", data.bank || ""]);
      return createResponse("success", "Đã lưu kết quả Quiz");
    }

    // 9. LƯU KẾT QUẢ THI TỔNG HỢP (Mặc định nếu có data.examCode)
    //if (data.examCode) {
     // let sheetResult = ss.getSheetByName("ketqua") || ss.insertSheet("ketqua");
      //sheetResult.appendRow([
       // new Date(), 
        //data.examCode, 
       // data.sbd, 
       // data.name, 
       // data.class, 
       // data.score, 
       // data.totalTime, 
       // data.idgv, 
        //JSON.stringify(data.details)]);
      //return createResponse("success", "Đã lưu kết quả thi");
   // }
// Kết thúc Dopost
    return createResponse("error", "Không khớp lệnh nào!");

  }
  catch (err) {
    return createResponse("error", err.toString());
  } finally {
    lock.releaseLock();
  }
}

// Hết dopost ###
// #09 CÁC HÀM PHỤ TRỢ (Để hết vào đây)
function getLinkFromRouting(idNumber) {
  const sheet = ssAdmin.getSheetByName("idgv");
  const data = sheet.getDataRange().getValues();
  const id = String(idNumber).trim();
  for (let i = 1; i < data.length; i++) {
    // Cột A: idNumber, Cột C: linkscript
    if (data[i][0].toString().trim() === id) {
      return data[i][2].toString().trim();
    }
  }
  return null;
}



function replaceIdInBlock(block, newId) {
  if (block.match(/id\s*:\s*\d+/)) return block.replace(/id\s*:\s*\d+/, "id: " + newId);
  return block.replace("{", "{\nid: " + newId + ",");
}


function getAppConfig() {
  var sheetCD = ss.getSheetByName("dangcd");
  // var dataCD = sheetCD.getDataRange().getValues();
  var dataCD = sheetCD.getDataRange().getDisplayValues();
  const lastRow = sheetCD.getLastRow();

  var topics = [];
  var classesMap = {}; // Dùng để lọc danh sách lớp không trùng lặp
  var maxtotal = dataCD[lastRow - 1][8] || 0;
  var maxcau = "(" + maxtotal + " câu hỏi" + ")";

  // Chạy từ dòng 2 (bỏ tiêu đề)
  for (var i = 1; i < lastRow - 1; i++) {
    var lop = dataCD[i][0];   // Cột A: lop
    var idcd = dataCD[i][1];  // Cột B: idcd
    var namecd = dataCD[i][2]; // Cột C: namecd
    var total = dataCD[i][8]; // cột I ghi tổng số câu   

    if (lop && lop.trim() !== "") {
      // 1. Đẩy vào danh sách chuyên đề
      topics.push({
        grade: lop,
        id: idcd,
        name: namecd,
        total: parseInt(total) || 0
        
      });

      // 2. Thu thập danh sách lớp (để nạp vào CLASS_ID bên React)
      // Ví dụ: Trong sheet có lớp 10, 11, 12 thì CLASS_ID sẽ có các lớp tương ứng
      classesMap[lop] = true;
    }
  }

  return {
    topics: topics,
    classes: Object.keys(classesMap).sort(function (a, b) { return a - b; }),    // Trả về [9, 10, 11, 12] chẳng hạn
    maxcau: maxcau
  };
}

function getAppConfigmt() {
  try {
    // Lưu ý: Đảm bảo ssAdmin đã được khai báo ở đầu script của bạn
    var sheetCD = ss.getSheetByName("dangcd");
    if (!sheetCD) return { topics: [] };

    var dataCD = sheetCD.getDataRange().getValues();
    var topics = [];

    // Chạy từ dòng 2 (bỏ tiêu đề)
    for (var i = 1; i < dataCD.length; i++) {
      var lop = dataCD[i][0];    // Cột A: lop
      var idcd = dataCD[i][1];   // Cột B: idcd
      var namecd = dataCD[i][2]; // Cột C: namecd
      var total = dataCD[i][8]; // Cột I

      if (idcd) {
        topics.push({
          grade: lop,            // Khối lớp (10, 11, 12)
          id: String(idcd),      // ID chuyên đề (để lưu vào matrix)
          name: String(namecd),   // Tên để hiển thị cho GV chọn
          total: total || 0
        });
      }
    }

    return { topics: topics };
  } catch (e) {
    return { topics: [], error: e.toString() };
  }
}
function parseDocByParagraph_(docId) {
  const body = DocumentApp.openById(docId).getBody();
  const paras = body.getParagraphs();

  let part = "";
  let current = null;
  const questions = [];

  paras.forEach(p => {
    const text = p.getText().trim();
    if (!text) return;

    // PHẦN
    if (/^Phần\s*I/i.test(text)) part = "MCQ";
    if (/^Phần\s*II/i.test(text)) part = "TF";
    if (/^Phần\s*III/i.test(text)) part = "SA";

    // CÂU HỎI
    if (/^Câu\s+\d+/i.test(text)) {
      if (current) questions.push(current);
      current = {
        part,
        question: text,
        options: [],
        answers: [],
        key: ""
      };
      return;
    }

    if (!current) return;

    // PHẦN III – KEY
    if (part === "SA") {
      const m = text.match(/<key\s*=\s*([^>]+)>/i);
      if (m) current.key = m[1].trim();
      else current.question += "\n" + text;
      return;
    }

    // PHẦN I & II – OPTION
    if (/^[A-D]\./.test(text)) {
      const letter = text[0];
      const isUnderline = hasUnderline_(p);
      current.options.push(text);

      if (isUnderline) {
        current.answers.push(letter);
      }
    } else {
      current.question += "\n" + text;
    }
  });

  if (current) questions.push(current);
  return questions;
}
// kiểm tra gạch chân
function hasUnderline_(paragraph) {
  const text = paragraph.editAsText();
  for (let i = 0; i < text.getText().length; i++) {
    if (text.getUnderline(i)) return true;
  }
  return false;
}
// chuẩn hóa trước khi ghi exam_data
function normalizeQuestion_(q) {
  if (q.part === "MCQ") {
    return {
      type: "MCQ",
      answer: q.answers[0] || ""
    };
  }

  if (q.part === "TF") {
    return {
      type: "TF",
      answer: q.answers.join(",")
    };
  }

  if (q.part === "SA") {
    return {
      type: "SA",
      answer: q.key
    };
  }
}

function shuffle(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// ==== Ghi exam_data


function parseQuestionFromCell(text, id) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  const qLine = lines.find(l => l.startsWith('?'));
  const question = qLine ? qLine.slice(1).trim() : '';
  const options = lines.filter(l => /^[A-D]\./.test(l)).map(l => l.slice(2).trim());
  const ansLine = lines.find(l => l.startsWith('='));
  const ansIndex = ansLine ? ansLine.replace('=', '').trim().charCodeAt(0) - 65 : -1;
  return { id, type: 'mcq', question, o: options, a: options[ansIndex] || '' };
}

function findDuplicateQuestions(targetTag) {  
  const sheet = ss.getSheetByName("nganhang"); 
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1); // Bỏ dòng tiêu đề
  const [targetClass, targetType, tileText] = supper(targetTag).split(".");
  const tile = Number(tileText);
  
  // BƯỚC 1: Lọc danh sách câu hỏi theo targetTag và lưu lại số dòng gốc (rowNumber)
  const filteredRows = [];  
  for (let i = 0; i < rows.length; i++) {
  // Lấy 4 số đầu của classTag (Cột 1)
  const classCode = String(rows[i][1]).substring(0, 4); 
  const t = String(rows[i][2]).toLowerCase().trim();
  
  // Khai báo let type ở đây để reset giá trị theo từng dòng
  let type = "mcq"; 
  if (t === "true-false" || t === "tf") {
    type = "tf";
  } else if (t === "short-answer" || t === "sa") {
    type = "sa";
  }  
  
  
  // Dùng hàm supper thầy viết để so sánh không sợ lệch chữ hoa/thường hay khoảng trắng
  if (supper(classCode) === targetClass && supper(type) === targetType) {
    filteredRows.push({
      rowData: rows[i],
      actualRowIndex: i + 2, // Dòng thực tế trên Sheet      
    });
  }
}
  
  const results = [];
  const processedIdx = new Set();

  // BƯỚC 2: Quét trùng trên mảng đã lọc chắt lọc
  for (let i = 0; i < filteredRows.length; i++) {
    if (processedIdx.has(i)) continue;
    
    let group = { 
      mainId: filteredRows[i].rowData[0], 
      score: 0, 
      items: [getRowObj(filteredRows[i].rowData, headers, filteredRows[i].actualRowIndex)] 
    };
    
    for (let j = i + 1; j < filteredRows.length; j++) {
      if (processedIdx.has(j)) continue;
      
      // So sánh dữ liệu câu i và câu j
      let score = calculateSimilarity(filteredRows[i].rowData, filteredRows[j].rowData);
      
      if (score >= tile) { 
        group.items.push(getRowObj(filteredRows[j].rowData, headers, filteredRows[j].actualRowIndex));
        if (score > group.score) group.score = score;
        processedIdx.add(j);
      }
    }
    
    if (group.items.length > 1) {
      results.push(group);
      processedIdx.add(i);
    }
  }
  return { status: "success", data: results };
}

function calculateSimilarity(q1, q2) {
  let score = 0;
  // Cột: 0:id, 1:classTag, 4:question, 5:options, 6:answer
  
  // 1. Answer (20%) - Bỏ latex $, khoảng trắng
  const a1 = String(q1[6]).replace(/\$|\s/g, '');
  const a2 = String(q2[6]).replace(/\$|\s/g, '');
  if (a1 !== "" && a1 === a2) score += 20;

  // 2. Options (30%) - Parse, làm sạch từng đáp án và so sánh không cần thứ tự
  const optStr1 = q1[5] ? String(q1[5]).trim() : "";
  const optStr2 = q2[5] ? String(q2[5]).trim() : "";

  try {
    if (optStr1 === "" && optStr2 === "") {
      score += 30; // Cả hai cùng trống (Dạng tự luận/điền số)
    } else {
      // Parse ra mảng, ép tất cả phần tử về dạng chuỗi, xóa khoảng trắng/chữ hoa chữ thường và dấu $
      const arr1 = JSON.parse(optStr1 || "[]").map(function(item) {
        return String(item).replace(/\$|\s/g, '').toLowerCase();
      });
      const arr2 = JSON.parse(optStr2 || "[]").map(function(item) {
        return String(item).replace(/\$|\s/g, '').toLowerCase();
      });
      
      // Sắp xếp và gộp lại để so sánh không quan trọng thứ tự A, B, C, D
      const o1 = arr1.sort().join('|');
      const o2 = arr2.sort().join('|');
      
      if (o1 !== "" && o1 === o2) score += 30;
    }
  } catch(e) {
    // Nếu lỗi parse JSON (do chuỗi lỗi), ta cứu bằng cách so sánh chuỗi thuần túy sau khi xóa khoảng trắng
    const rawO1 = optStr1.replace(/\$|\s/g, '').toLowerCase();
    const rawO2 = optStr2.replace(/\$|\s/g, '').toLowerCase();
    if (rawO1 !== "" && rawO1 === rawO2) score += 30;
  }

  // 3. Question (40%) - Xóa khoảng trắng và chữ hoa/thường
  const txt1 = String(q1[4]).replace(/\s+/g, '').toLowerCase();
  const txt2 = String(q2[4]).replace(/\s+/g, '').toLowerCase();
  if (txt1 !== "" && txt1 === txt2) score += 40; 

  // CHẠM TRẦN: Vì đã bỏ điều kiện 4 nên điểm tối đa là 90. 
  // Nếu đạt từ 85 trở lên coi như trùng tuyệt đối (Trả về 99)
  if (score >= 85) return 99;
  return score;
}
function getRowObj(row, headers, rowIdx) {
  let obj = { rowIdx: rowIdx };
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}

function deleteQuestionRow(rowIdx) {
  try {
    const sheet = ss.getSheetByName("nganhang");
    sheet.deleteRow(parseInt(rowIdx));
    return { status: "success" };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}
// ======= sửa câu hỏi =====================================================
function updateQuestion(payload) {
  try {
    const data = payload.data;
    const sheet = sheetNH;
    const fullData = sheet.getDataRange().getValues();
    const headers = fullData[0];
    
    // 1. Kiểm tra ID gửi lên có tồn tại không
    if (!data.id) return { status: "error", message: "ID gửi lên bị trống!" };

    // 2. Duyệt tìm dòng
    for (var i = 1; i < fullData.length; i++) {
      // KIỂM TRA: Nếu ô ID bị trống thì bỏ qua dòng này, không .toString() nữa
      if (!fullData[i][0]) continue; 

      // So sánh ID an toàn
      if ((fullData[i][0] || "").toString() === data.id.toString()) {
        const rowNum = i + 1;
        
        // Cập nhật các cột dựa trên tên Header
        Object.keys(data).forEach(key => {
          const colIdx = headers.indexOf(key);
          if (colIdx !== -1) {
            sheet.getRange(rowNum, colIdx + 1).setValue(data[key]);
          }
        });
        
        return { status: "success" };
      }
    }
    return { status: "error", message: "Không tìm thấy ID: " + data.id };
  } catch (e) {
    return { status: "error", message: "Lỗi hệ thống: " + e.toString() };
  }
}
// 3107sua4: Lọc danh sách mã exams theo idgv của giáo viên
function getExamsList(type, idgv, sheetId) {
  let sheetName;
  let colIdgv;
  let colExams;

  if (type === "ketqua") {
    sheetName = "ketqua";
    colIdgv = 8;     // Cột H (idgv)
    colExams = 2;    // Cột B (exams)
  }
  else if (type === "matran") {
    sheetName = "matran";
    colIdgv = 18;    // Cột R (idgv)
    colExams = 2;    // Cột B (exams)
  }
  else if (type === "exams") {
    sheetName = "exams";
    colIdgv = 2;     // Cột B (idgv)
    colExams = 1;    // Cột A (Exams)
  }
  else if (type === "exam_data") {
    sheetName = "exam_data";
    colIdgv = 8;     // Cột H (idgv)
    colExams = 1;    // Cột A (exams)
  }
  else {
    return createResponse("error", "Type không hợp lệ");
  }

  const ss2 = getSS2(sheetId, idgv);
  const sheet = ss2.getSheetByName(sheetName);
  if (!sheet) {
    return createResponse("error", "Không tìm thấy sheet " + sheetName);
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return createResponse("success", "OK", []);
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  const idgvTarget = idgv ? supper(idgv) : "";
  const idgvN9 = idgv ? N9(idgv) : "";

  const examsList = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIdgv = row[colIdgv - 1];
    if (!idgv || (rowIdgv && (supper(rowIdgv) === idgvTarget || N9(rowIdgv) === idgvN9))) {
      const examVal = row[colExams - 1] ? row[colExams - 1].toString().trim() : "";
      if (examVal) {
        examsList.push(examVal);
      }
    }
  }

  const unique = [...new Set(examsList)];

  return createResponse("success", "OK", unique);
}
// 3107sua4: Reset dữ liệu theo idgv hoặc theo mã đề chính xác
function resetData(type, password, mode, exams, idgv, sheetId) {  
  const idgvStr = (idgv || "").toString().trim();
  const examStr = (exams || "").toString().trim();

  let colIdgv = 0;
  let colExamsIdgv = 0;
  let sheetName = "";

  if (type === "ketqua") {
    sheetName = "ketqua";
    colIdgv = 8;        // Cột H (idgv)
    colExamsIdgv = 10;  // Cột J (exams.idgv)
  }
  else if (type === "matran") {
    sheetName = "matran";
    colIdgv = 18;       // Cột R (idgv)
    colExamsIdgv = 19;  // Cột S (exams.idgv)
  }
  else if (type === "exams") {
    sheetName = "exams";
    colIdgv = 2;        // Cột B (idgv)
    colExamsIdgv = 15;  // Cột O (exams.idgv)
  }
  else if (type === "exam_data") {
    sheetName = "exam_data";
    colIdgv = 8;        // Cột H (idgv)
    colExamsIdgv = 9;   // Cột I (exams.idgv)
  }
  else return createResponse("error", "Loại dữ liệu (Type) không hợp lệ");

  let rowsDeleted = 0;
  var ss2 = getSS2(sheetId, idgvStr);

  if (mode === "all") {
    rowsDeleted = deleteFastAll(idgvStr, colIdgv, sheetName, ss2);
    return createResponse("success", "Đã dọn sạch " + rowsDeleted + " dòng trong sheet " + sheetName);
  }

  if (mode === "byExams") {
    if (!examStr) return createResponse("error", "Thiếu mã bài tập (exams)");
    
    let keyexamsid = "";
    if (supper(examStr).indexOf("." + supper(idgvStr)) !== -1 || supper(examStr).indexOf("." + N9(idgvStr)) !== -1) {
      keyexamsid = examStr;
    } else {
      keyexamsid = examStr + "." + idgvStr;
    }

    rowsDeleted = deleteFast(keyexamsid, colExamsIdgv, sheetName, ss2);
    return createResponse("success", "Đã xóa " + rowsDeleted + " dòng của " + examStr + " (" + sheetName + ")");
  }

  return createResponse("error", "Chế độ (Mode) không hợp lệ");
}
// =============================================================Kết thúc Reset chung=========================================================================

// xem điểm
function getScore(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    
    const searchExams = supper(params.exams || "");
    const searchSbd   = supper(params.sbd || "");
    const searchIdgv  = N9(params.idgv || "");
    const reqSheetId  = params.sheetId || "";

    if (!searchExams || !searchSbd || !searchIdgv) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "error", message: "Thiếu thông tin tra cứu!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss2 = getSS2(reqSheetId, searchIdgv);
    const sheet = ss2.getSheetByName("ketqua");
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "error", message: "Không tìm thấy sheet kết quả!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "not_found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

    for (let i = 0; i < data.length; i++) {
      const dbExams = supper(data[i][1] || "");
      const dbSbd   = supper(data[i][2] || "");
      const dbIdgv  = N9(data[i][7] || "");

      if (dbSbd === searchSbd && dbExams === searchExams && dbIdgv === searchIdgv) {
        return ContentService
          .createTextOutput(JSON.stringify({
            status: "success",
            data: {
              exams: data[i][1],
              sbd: data[i][2],
              name: data[i][3],
              class: data[i][4],
              tongdiem: data[i][5],
              time: data[i][6]
            }
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "not_found" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "Lỗi hệ thống: " + error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function createResponseW(status, message, data = null) {
  const output = { status: status, message: message };
  if (data !== null) output.data = data;
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
function createResponse(status, message, data) {
  const output = { status: status, message: message };
  if (data) output.data = data;
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// Giữ lại resJSON để phục vụ các đoạn code cũ đang gọi tên này
function resJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function N9(id) {
  if (id === null || id === undefined) return "";
  var str = id.toString().replace(/[^0-9a-zA-Z]/g, "").trim();
  if (str.length > 1) {
    str = str.toLowerCase().slice(-9);
  }
  return str;
}
function Right(text, n) {
  return text.toString().trim().slice(-n);
}

function supper(text) {
  if (text === null || text === undefined) return "";
   return text.toString().replace(/'/g, "").toUpperCase().trim()
}
  

// Hàm xóa nhiều dòng //
/**
 * Xóa dữ liệu cực nhanh và GIỮ LẠI dòng tiêu đề (Header)
 */

// 3107sua4: Xóa toàn bộ dữ liệu theo IDGV
function deleteFastAll(text, colNumber, name, targetSS) {  
  var target = targetSS || ss;
  var sheet = target.getSheetByName(name);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow <= 1) return 0;

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  var keySupper = supper(text);
  var keyN9 = N9(text);

  var filteredData = data.filter(function(row) {
    var cellVal = row[colNumber - 1];
    var cellSupper = supper(cellVal);
    var cellN9 = N9(cellVal);

    return !(cellSupper === keySupper || (cellN9 && cellN9 === keyN9));
  });

  var deletedCount = data.length - filteredData.length;

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  if (filteredData.length > 0) {
    sheet.getRange(2, 1, filteredData.length, lastCol).setValues(filteredData);
  }

  return deletedCount;
}

// 3107sua4: Xóa dữ liệu theo mã exams.idgv
function deleteFast(text, colNumber, name, targetSS) {  
  var target = targetSS || ss;
  var sheet = target.getSheetByName(name);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow <= 1) return 0;

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  var keyTarget = supper(text);
  var keyTargetN9 = "";
  if (text && text.indexOf(".") !== -1) {
    var parts = text.split(".");
    keyTargetN9 = supper(parts[0] + "." + N9(parts[1]));
  }

  var filteredData = data.filter(function(row) {
    var cellVal = row[colNumber - 1];
    var cellSupper = supper(cellVal);

    if (keyTargetN9) {
      return !(cellSupper === keyTarget || cellSupper === keyTargetN9);
    }
    return cellSupper !== keyTarget;
  });

  var deletedCount = data.length - filteredData.length;  

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  if (filteredData.length > 0) {
    sheet.getRange(2, 1, filteredData.length, lastCol).setValues(filteredData);
  }

  return deletedCount;
}


// Đăng ký GameShow
  function register(e) {
  const phone = (e.parameter.phone || "").trim();
  const pass = (e.parameter.pass || "").trim();

  if (!phone || !pass) {
    return createResponse("error", "Thiếu dữ liệu!");
  }

  const sheet = ssAdmin.getSheetByName("gameshow");
  const data = sheet.getDataRange().getValues();

  // kiểm tra trùng
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == phone) {
      return createResponse("exists", "Số điện thoại đã tồn tại!");
    }
  }

  sheet.appendRow([
    new Date(),
    "'" + phone,
    pass,
    "VIP0",
    ""
  ]);

  return createResponse("success", "Đăng ký thành công!");
}

function login(e) {
  const phone = (e.parameter.phone || "").trim();
  const pass = (e.parameter.pass || "").trim();

  const sheet = ssAdmin.getSheetByName("gameshow");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    // 🔥 FIX QUAN TRỌNG: bỏ dấu '
    const phoneSheet = data[i][1].toString().replace("'", "").trim();
    const passSheet = data[i][2].toString().trim();

    if (phoneSheet === phone && passSheet === pass) {
      return createResponse("success", "OK", {
        phone: phoneSheet,
        vip: data[i][3] || "VIP0",
        name: data[i][4] || ""
      });
    }
  }

  return createResponse("fail", "Sai tài khoản hoặc mật khẩu!");
}

function adminLogin(e) {
  const id = (e.parameter.id || "").trim();
  const pass = (e.parameter.pass || "").trim();

  if (!id || !pass) {
    return createResponse("error", "Thiếu dữ liệu!");
  }

  if (id === idadmin && supper(pass) === supper(passAdmin)) {
    return createResponse("success", "OK");
  }

  return createResponse("error", "Sai tài khoản!");
}


function getUsers(e) {
  const id = e.parameter.id;
  const pass = e.parameter.pass;

  if (id !== idadmin || pass !== passAdmin) {
    return jsonOut({ status: "error", message: "Unauthorized" });
  }
  const sheet = ssAdmin.getSheetByName("gameshow");

  const data = sheet.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    users.push({
      phone: data[i][1],
      vip: data[i][3],
      name: data[i][4]
    });
  }

  return jsonOut({ status: "success", data: users });
}

function updatePassword(e) {
  const phone = (e.parameter.phone || "").trim();
  const newPass = (e.parameter.pass || "").trim();

  if (!phone || !newPass) {
    return createResponse("error", "Thiếu dữ liệu!");
  }

  const sheet = ssAdmin.getSheetByName("gameshow");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    // 🔥 xử lý dấu ' ở số điện thoại
    const phoneSheet = data[i][1].toString().replace("'", "").trim();

    if (phoneSheet === phone) {
      sheet.getRange(i + 1, 3).setValue(newPass); // cột C = password

      return createResponse("success", "Đổi mật khẩu thành công!");
    }
  }

  return createResponse("error", "Không tìm thấy tài khoản!");
}

function getTeachers() {
  const sheet = ssAdmin.getSheetByName("gameshow");
  const data = sheet.getDataRange().getValues();

  const result = [];

  for (let i = 1; i < data.length; i++) {
    result.push({
      phone: data[i][1].toString().replace("'", ""),
      password: data[i][2],
      vip: data[i][3] || "VIP0",
      createdAt: data[i][0]
    });
  }

  return createResponse("success", "OK", result);
}

function updateTeacher(e) {
  const phone = (e.parameter.phone || "").trim();
  const pass = (e.parameter.pass || "").trim();
  const vip = (e.parameter.vip || "").trim();

  const sheet = ssAdmin.getSheetByName("gameshow");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const phoneSheet = data[i][1].toString().replace("'", "").trim();

    if (phoneSheet === phone) {

      if (pass) sheet.getRange(i + 1, 3).setValue(pass);
      if (vip) sheet.getRange(i + 1, 4).setValue(vip);

      return createResponse("success", "Đã cập nhật!");
    }
  }

  return createResponse("error", "Không tìm thấy GV!");
}
// Các hàm hỗ trợ tìm câu trùng

// Hàm làm sạch văn bản để so sánh chính xác
function cleanForCompare(txt) {
  if (!txt) return "";
  return String(txt).toLowerCase()
    .replace(/<[^>]*>/g, "") // Bỏ HTML
    .replace(/\\s/g, "")     // Bỏ khoảng trắng trong LaTeX
    .replace(/\s+/g, "")     // Bỏ mọi khoảng trắng
    .trim();
}

// Hàm tính % tương đồng giữa 2 chuỗi văn bản (Dùng thuật toán Dice's Coefficient đơn giản)
function textSimilarity(str1, str2) {
  let s1 = cleanForCompare(str1);
  let s2 = cleanForCompare(str2);
  if (s1 === s2) return 100;
  if (s1.length < 2 || s2.length < 2) return 0;

  let bigrams1 = new Set();
  for (let i = 0; i < s1.length - 1; i++) bigrams1.add(s1.substring(i, i + 2));
  
  let intersect = 0;
  for (let i = 0; i < s2.length - 1; i++) {
    if (bigrams1.has(s2.substring(i, i + 2))) intersect++;
  }

  return (2.0 * intersect) / (s1.length + s2.length - 2) * 100;
}
function getNGramSimilarity(str1, str2) {
  const n = 5; // Độ dài cụm ký tự theo ý bạn
  const clean = (txt) => String(txt).replace(/\s+/g, ""); // Xóa mọi khoảng trắng để so sánh chính xác
  
  let s1 = clean(str1);
  let s2 = clean(str2);
  
  if (s1 === s2) return 100;
  if (s1.length < n || s2.length < n) return 0;

  // Tạo tập hợp các cụm 5 ký tự của chuỗi 1
  let nGrams1 = new Set();
  for (let i = 0; i <= s1.length - n; i++) {
    nGrams1.add(s1.substring(i, i + n));
  }

  // Đếm xem chuỗi 2 có bao nhiêu cụm trùng
  let matches = 0;
  let totalNGrams2 = s2.length - n + 1;
  for (let i = 0; i <= s2.length - n; i++) {
    if (nGrams1.has(s2.substring(i, i + n))) {
      matches++;
    }
  }

  // Tính % trung bình dựa trên cả hai chuỗi
  let totalNGrams1 = s1.length - n + 1;
  return (matches * 2) / (totalNGrams1 + totalNGrams2) * 100;
}


// --- HÀM BỔ TRỢ BĂM SHA256 (Thầy dán hàm này ở cuối file Script tổng) Xóa ảnh cloud ---
function SHA256_(input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  var output = "";
  for (var i = 0; i < rawHash.length; i++) {
    var v = rawHash[i];
    if (v < 0) v += 256;
    if (v < 16) output += "0";
    output += v.toString(16);
  }
  return output;
}
// 2308sua1: Sửa verifyhocsinh trả về boolean chuẩn và so khớp linh hoạt SBD/IDGV/Khóa ghép H
function verifyhocsinh(sbd, idgv, pass, sheetId) {
  const ss2 = getSS2(sheetId, idgv);
  const sheet = ss2.getSheetByName("danhsach");
  if (!sheet) {
    return false;
  }

  const cleanSbd = (sbd || "").toString().trim().toUpperCase();
  const cleanIdgv = (idgv || "").toString().trim().toUpperCase();
  const cleanPass = (pass || "").toString().trim();
  const keyds = supper(cleanSbd + "." + cleanIdgv);

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const rowSbd  = (values[i][0] || "").toString().trim().toUpperCase();
    const rowIdgv = (values[i][5] || "").toString().replace(/'/g, "").trim().toUpperCase();
    const rowSbdKey = (values[i][7] || "").toString().trim().toUpperCase();
    const rowPass = (values[i][8] || "").toString().replace(/'/g, "").trim();

    if (rowPass === cleanPass && ((rowSbd === cleanSbd && (N9(rowIdgv) === N9(cleanIdgv) || rowIdgv === cleanIdgv)) || rowSbdKey === keyds)) {
      return true;
    }
  }

  return false; 
}
// 2308sua1: Sửa verifyExams trả về boolean/link chuẩn và so khớp linh hoạt mã đề/IDGV/Khóa ghép O
function verifyExams(examcode, idgv, sheetId) {
  const ss2 = getSS2(sheetId, idgv);
  const sheet = ss2.getSheetByName("exams");
  if (!sheet) {
    return false;
  }

  const cleanExamcode = (examcode || "").toString().trim().toUpperCase();
  const cleanIdgv = (idgv || "").toString().trim().toUpperCase();
  const keyexams = supper(cleanExamcode + "." + cleanIdgv);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const rowCode  = (values[i][0] || "").toString().trim().toUpperCase();
    const rowIdgv = (values[i][1] || "").toString().replace(/'/g, "").trim().toUpperCase();    
    const rowExamKey = (values[i][14] || "").toString().trim().toUpperCase();

    if (((rowCode === cleanExamcode && (N9(rowIdgv) === N9(cleanIdgv) || rowIdgv === cleanIdgv)) || rowExamKey === keyexams)) {
      return values[i][18] || true;
    }
  }
  return false; 
}
// Hàm kiểm tra ngày đóng - mở
function opencloseDate(sheetDateVal, type) {
  // Xử lý ô trống theo ý thầy: "Không nhập nghĩa là để thời gian mãi mãi"
  if (!sheetDateVal || sheetDateVal.toString().trim() === "") {
    if (type === 'open') return true;   // Ô mở trống -> Coi như ĐÃ vượt qua mốc mở (mở mãi mãi)
    if (type === 'close') return false; // Ô đóng trống -> Coi như CHƯA vượt qua mốc đóng (mở mãi mãi)
    return false;
  }
  
  const now = new Date();
  let targetDate = null;
  
  if (sheetDateVal instanceof Date) {
    targetDate = sheetDateVal;
  } else {
    targetDate = new Date(sheetDateVal.toString().trim().replace(' ', 'T'));
  }
  
  if (isNaN(targetDate.getTime())) return false; 
  
  return now > targetDate;
}

/**
 * Hàm lấy mật khẩu giáo viên từ file Admin
 * @param {string|number} idgv - ID giáo viên (đó cũng chính là TÊN SHEET chứa thông tin GV)
 * @return {string} Mật khẩu giáo viên (hoặc "" nếu không tìm thấy)
 */
function passteacher(idgv) {
  try {
    if (!idgv) return "";
    
    // 1. Mở file SS Admin theo ID File
    // ⚠️ THẦY ĐỔI 'ID_FILE_SS_ADMIN' THÀNH ID THỰC TẾ CỦA FILE ADMIN NHÉ!    
    
    // 2. Tìm sheet có tên chính là idgv
    const sheet = ssAdmin.getSheetByName("idgv");
    if (!sheet) {
      Logger.log("Không tìm thấy sheet của GV: " + idgv);
      return "";
    }

    // 3. Lấy toàn bộ dữ liệu cột A (idgv) và cột C (passGV)
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return ""; // Sheet rỗng hoặc chỉ có hàng tiêu đề

    // Lấy cột A (cột 1) đến cột C (cột 3)
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    // 4. Tìm dòng có cột A trùng khớp với idgv và trả về cột C (index 2)
    const targetId = N9(idgv);
    for (let i = 0; i < data.length; i++) {
      const rowId = N9(data[i][0]); // Cột A
      if (rowId === targetId) {
        return String(data[i][2]).trim();     // Cột C (Mật khẩu)
      }
    }

    // Nếu duyệt hết mà không thấy ID trùng khớp ở cột A
    return "";

  } catch (error) {
    Logger.log("Lỗi trong hàm passteacher: " + error.toString());
    return "";
  }
}

// 3107them5: Kiểm tra xác thực IDGV và mật khẩu
function checkTeacherAuth(idgv, password) {
  if (!idgv || !password) return false;
  var targetId = N9(idgv);
  var passStr = String(password).trim();

  var p1 = passteacher(idgv);
  if (p1 && p1.trim() === passStr) return true;

  try {
    var sheet = ssAdmin.getSheetByName("idgv");
    if (sheet) {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var rowId = N9(rows[i][0]);
        var col2Pass = String(rows[i][1] || "").trim();
        var col3Pass = String(rows[i][2] || "").trim();
        var col6Pass = String(rows[i][5] || "").trim();
        if (rowId === targetId && (col2Pass === passStr || col3Pass === passStr || col6Pass === passStr)) {
          return true;
        }
      }
    }
  } catch(e) {}

  return false;
}

// 2508sua1: Lấy danh sách các mã đề theo thể loại (Matrix, Word) ở sheet ketqua của ss2 (bắt buộc dùng ss2, không cần kiểm tra idgv) chamlai
function getRegradeExamsList(idgv, sheetId, theloai) {  
  // 2508sua1: Khối try catch bắt lỗi
  try {
    // 2508sua1: Chuẩn hóa thể loại được yêu cầu
    var targetTheloai = (theloai || "").trim().toLowerCase();
    // 2508sua1: Kiểm tra bắt buộc có idgv và thể loại (loại trừ pdf)
    if (!idgv || !theloai || targetTheloai === "pdf") return createResponse("error", "Thiếu thông tin rồi bạn ơi");
    // 2508sua1: Mở sheet ketqua trong ss2 thông qua getSS2Sheet
    var sheet = getSS2Sheet(sheetId, idgv, "ketqua");
    // 2508sua1: Fallback nếu getSS2Sheet không trả về
    if (!sheet) {
      // 2508sua1: Mở ss2
      var ss2 = getSS2(sheetId, idgv);
      // 2508sua1: Lấy sheet ketqua
      sheet = ss2 ? (ss2.getSheetByName("ketqua") || ss2.getSheetByName("KetQua") || ss2.getSheetByName("KETQUA")) : null;
    }
    // 2508sua1: Nếu không có sheet ketqua thì trả về danh sách rỗng
    if (!sheet) return createResponse("success", "OK", []);

    // 2508sua1: Lấy số dòng cuối cùng của sheet ketqua
    var lastRow = sheet.getLastRow();
    // 2508sua1: Nếu không có dòng dữ liệu nào từ dòng 2 trở đi thì trả về rỗng
    if (lastRow < 2) return createResponse("success", "OK", []);

    // 2508sua1: Lấy dữ liệu 12 cột từ dòng 2
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    // 2508sua1: Khởi tạo mảng lưu danh sách mã đề
    var examsList = [];

    // 2508sua1: Duyệt qua từng dòng trong sheet ketqua
    for (var i = 0; i < data.length; i++) {
      // 2508sua1: Lấy dữ liệu dòng hiện tại
      var row = data[i];
      // 2508sua1: Lấy thể loại từ cột L (Index 11)
      var theloaisheet = row[11] ? String(row[11]).trim().toLowerCase() : "";     
      // 2508sua1: So sánh chính xác theo thể loại yêu cầu
      if (theloaisheet === targetTheloai) {
        // 2508sua1: Lấy mã đề từ cột B (Index 1) và bỏ dấu nháy đơn đầu dòng nếu có
        var examCode = row[1] ? String(row[1]).trim().replace(/^'/, '') : "";
        // 2508sua1: Nếu có mã đề hợp lệ thì thêm vào mảng
        if (examCode) {
          // 2508sua1: Đẩy mã đề vào danh sách
          examsList.push(examCode);
        }
      }
    }

    // 2508sua1: Lọc bỏ các mã đề bị trùng lặp
    var uniqueExams = [...new Set(examsList)];
    // 2508sua1: Trả về kết quả thành công kèm mảng mã đề
    return createResponse("success", "OK", uniqueExams);
  } catch (err) {
    // 2508sua1: Xử lý và trả về thông báo lỗi khi có ngoại lệ
    return createResponse("error", "Lỗi lấy danh sách mã đề: " + err.toString());
  }
}

// 2608sua4: Chấm lại bài thi Matran & Word tách riêng luồng xử lý và dùng chung phần ghi kết quả
function regradeExams(idgv, password, examCode, sheetId, theloai) {
  try {
    // 1. Chuẩn hóa dữ liệu đầu vào
    var idgvStr = String(idgv || "").trim();
    var passStr = String(password || "").trim();
    var examStr = String(examCode || "").trim().replace(/^'/, '');
    var theloaiStr = String(theloai || "").trim().toLowerCase();

    if (!idgvStr) return createResponse("error", "Vui lòng nhập ID Giáo viên!");
    if (!passStr) return createResponse("error", "Vui lòng nhập Mật khẩu!");
    if (!theloaiStr) return createResponse("error", "Vui lòng chọn thể loại đề thi PDF/Word/Matrix!");
    if (!examStr) return createResponse("error", "Vui lòng chọn Mã đề!");

    // 2. Xác thực giáo viên
    var isAuth = checkTeacherAuth(idgvStr, passStr);
    if (!isAuth) {
      return createResponse("error", "Mật khẩu hoặc ID Giáo viên không đúng!");
    }

    // 3. Lấy Bảng tính ss2 và Sheet ketqua
    var ss2 = getSS2(sheetId, idgvStr);
    var sheetKq = ss2 ? (ss2.getSheetByName("ketqua") || ss2.getSheetByName("KetQua") || ss2.getSheetByName("KETQUA")) : null;   

    if (!sheetKq) return createResponse("error", "Sheet ketqua không tồn tại trong ss2!");

    var lastRowKq = sheetKq.getLastRow();
    if (lastRowKq < 2) return createResponse("error", "Sheet ketqua không có dữ liệu!");

    // Sắp xếp dữ liệu theo cột Mã đề (Cột 2)
    if (typeof sapxep_1 === "function") sapxep_1(2, sheetKq);

    var dataKq = sheetKq.getDataRange().getValues();    
    var targetExamSupper = (typeof supper === "function") ? supper(examStr) : examStr.toUpperCase();

    var matchingRowIndices = [];
    var matchingDetails = [];    
    var arraydiemcu = [];
    var arraydiemchamlai = [];

    // 4. Lọc bài làm cần chấm lại theo Mã đề & Thể loại
    for (var i = 1; i < dataKq.length; i++) {
      var row = dataKq[i];
      var rowExams = row[1] ? String(row[1]).trim().replace(/^'/, '') : "";
      var rowTheloai = row[11] ? String(row[11]).replace(/'/g, '').trim().toLowerCase() : "";

      var matchesExam = ((typeof supper === "function" ? supper(rowExams) : rowExams.toUpperCase()) === targetExamSupper);
      var matchesType = (rowTheloai === theloaiStr);                    

      if (matchesExam && matchesType) {
        matchingRowIndices.push(i + 1); // Dòng thực tế trên Sheet (1-indexed)
        matchingDetails.push(row[12]);   // Cột M: Detail bài làm
        arraydiemcu.push(row[5]);        // Cột F: Điểm cũ
        var diemP = row[15] !== undefined ? row[15] : "";
        arraydiemchamlai.push(diemP);
      }
    }

    if (matchingRowIndices.length === 0) {
      return createResponse("error", "Không tìm thấy bài làm nào của mã đề " + examStr + " thuộc thể loại " + theloaiStr + "!");
    }

    var startRow = matchingRowIndices[0];
    var numRows = matchingRowIndices.length;

    // 5. Chấm điểm tách biệt theo thể loại đề Word / Matrix
    var gradeResult = null;

    if (theloaiStr === "word") {
      gradeResult = regradeWordExams_(ss2, targetExamSupper, matchingDetails);
    } else if (theloaiStr === "matrix" || theloaiStr === "matran" || theloaiStr === "ma trận") {
      gradeResult = regradeMatrixExams_(ss2, targetExamSupper, matchingDetails);
    } else {
      return createResponse("error", "Thể loại '" + theloaiStr + "' hiện chưa được hỗ trợ chấm lại!");
    }

    if (!gradeResult || !gradeResult.success) {
      return createResponse("error", (gradeResult && gradeResult.message) ? gradeResult.message : "Chấm lại không thành công!");
    }

    var arraydiem = gradeResult.arraydiem || [];
    var arraynx = gradeResult.arraynx || [];

    // 6. GHI HÀNG LOẠT (BULK WRITE) TỪ DÒNG matchingRowIndices[0] - Ghi điểm mới và bảo lưu điểm cũ
    if (numRows > 0) {
      // Đóng gói mảng 2 chiều cho Cột N, O, P, Q (Cột 14, 15, 16, 17)
      var bulkDataNOP = [];
      for (var r = 0; r < numRows; r++) {
        var nxVal = arraynx[r] ? arraynx[r][0] : "";
        var oldScoreVal = arraydiemcu[r] !== undefined ? arraydiemcu[r] : "";
        var oldDiemchamlai = arraydiemchamlai[r] !== undefined ? arraydiemchamlai[r] : "";
        
        if (String(oldDiemchamlai).trim() !== "") {
          bulkDataNOP.push([nxVal, "Chấm lại", oldDiemchamlai, oldScoreVal]);          
        } else {
          bulkDataNOP.push([nxVal, "Chấm lại", oldScoreVal, ""]); 
        }
      }

      // Ghi hàng loạt Cột F (Điểm mới) - Cột 6
      sheetKq.getRange(startRow, 6, numRows, 1).setValues(arraydiem);

      // Ghi hàng loạt 4 cột N, O, P, Q (Nhận xét, Trạng thái, Điểm gốc, Điểm lần chấm trước) - Cột 14
      sheetKq.getRange(startRow, 14, numRows, 4).setValues(bulkDataNOP);
    }

    return createResponse("success", "Đã chấm lại thành công " + numRows + " bài làm cho mã đề " + examStr + "!");

  } catch (err) {
    return createResponse("error", "Lỗi trong quá trình chấm lại: " + err.toString());
  }
}

/**
 * 2608sua4: Hàm chấm lại riêng cho thể loại đề WORD
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss2 - Bảng tính ss2 của GV
 * @param {string} targetExamSupper - Mã đề dạng chữ hoa
 * @param {Array<string|Object>} matchingDetails - Mảng chứa detail của các bài làm cần chấm
 * @return {Object} { success: boolean, arraydiem: Array<Array<number>>, arraynx: Array<Array<string>>, message?: string }
 */
function regradeWordExams_(ss2, targetExamSupper, matchingDetails) {
  // 1. Lấy thang điểm đề Word từ sheet exams
  var scMCQ = 0.25;
  var scTF = 1.0;
  var scSA = 0.5;

  var sheetExams = ss2.getSheetByName("exams");
  if (sheetExams && sheetExams.getLastRow() >= 2) {     
    var dataExams = sheetExams.getRange(2, 1, sheetExams.getLastRow() - 1, sheetExams.getLastColumn()).getValues();
    for (var eIdx = 0; eIdx < dataExams.length; eIdx++) {
      var exRow = dataExams[eIdx];
      var rowExKey = (typeof supper === "function") ? supper(String(exRow[0] || "")) : String(exRow[0] || "").toUpperCase();
      var exTheloai = String(exRow[19] || exRow[20] || "").trim().toLowerCase();
      
      if (rowExKey === targetExamSupper && (exTheloai === "word" || !exRow[19])) {
        var val3 = parseNum(exRow[3]);
        if (val3 > 0) scMCQ = val3;
        var val5 = parseNum(exRow[5]);
        if (val5 > 0) scTF = val5;
        var val7 = parseNum(exRow[7]);
        if (val7 > 0) scSA = val7;
        break;
      }
    }
  }

  // 2. Lấy dữ liệu câu hỏi từ sheet exam_data trong ss2
  var sheetWord = ss2.getSheetByName("exam_data");
  if (!sheetWord || sheetWord.getLastRow() < 2) {
    return { success: false, message: "Không tìm thấy sheet exam_data hoặc sheet không có dữ liệu!" };
  }

  var dataWord = sheetWord.getDataRange().getValues();
  var wordMap = {};

  for (var i = 1; i < dataWord.length; i++) {
    var idq = String(dataWord[i][1] || "").trim();
    if (!idq) continue;

    var typeq = String(dataWord[i][3] || "").trim().toLowerCase();
    var questionq = dataWord[i][4] || "";

    // Bóc answer từ JSON của exam_data
    var examMap = parseExamData(questionq);
    var ansE = examMap[idq];

    if (ansE === undefined || ansE === null) {
      ansE = dataWord[i][6] !== undefined && dataWord[i][6] !== "" ? dataWord[i][6] : dataWord[i][5];
    }

    wordMap[idq] = {
      type: typeq,
      answer: ansE
    };
  }

  // 3. Tiến hành chấm từng bài làm của học sinh
  var arraydiem = [];
  var arraynx = [];

  for (var k = 0; k < matchingDetails.length; k++) {
    var rawDetail = matchingDetails[k];
    var parsedStudent = parseDetailData(rawDetail);
    var listIdkq = parsedStudent.arrayIddetail || [];
    var listanswerkq = parsedStudent.arrayanswer || [];      
    var totalScore = 0;

    for (var j = 0; j < listIdkq.length; j++) {
      var detailId = String(listIdkq[j] || "").trim();
      var bankItem = wordMap[detailId];
      if (!bankItem) continue;

      var qType = String(bankItem.type || "").toLowerCase().trim();
      var ansExam = bankItem.answer;
      var ansStudent = listanswerkq[j];

      if (qType === "mcq" || qType.indexOf("phần i") !== -1 || (!qType && !Array.isArray(ansExam))) {
        if (normalizeAns(ansStudent) === normalizeAns(ansExam)) {
          totalScore += scMCQ;
        }
      } 
      else if (qType === "true-false" || qType === "tf" || qType.indexOf("phần ii") !== -1) {
        var tfStudent = parseTfOptions(ansStudent);
        var expectedTf = Array.isArray(ansExam) ? ansExam : parseTfOptions(ansExam);
        var point = pointtf(scTF, tfStudent, expectedTf);
        totalScore += point;
      } 
      else if (qType === "short-answer" || qType === "sa" || qType.indexOf("phần iii") !== -1) {
        if (normalizeAns(ansStudent) === normalizeAns(ansExam)) {
          totalScore += scSA;        
        }
      }
    }     
    
    var finalScore = Math.round(totalScore * 100) / 100;      
    var numericScore = typeof finalScore === "number" ? finalScore : (parseFloat(String(finalScore).replace(",", ".")) || 0);
    var nx = (typeof layNhanXet === "function") ? layNhanXet(numericScore) : "Hoàn thành bài thi";
    
    arraydiem.push([finalScore]); 
    arraynx.push([nx]);           
  }

  return {
    success: true,
    arraydiem: arraydiem,
    arraynx: arraynx
  };
}

/**
 * 2608sua4: Hàm chấm lại riêng cho thể loại đề MATRIX (Ma trận)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss2 - Bảng tính ss2 của GV
 * @param {string} targetExamSupper - Mã đề dạng chữ hoa
 * @param {Array<string|Object>} matchingDetails - Mảng chứa detail của các bài làm cần chấm
 * @return {Object} { success: boolean, arraydiem: Array<Array<number>>, arraynx: Array<Array<string>>, message?: string }
 */
function regradeMatrixExams_(ss2, targetExamSupper, matchingDetails) {
  // 1. Lấy thang điểm đề Ma trận từ sheet matran
  var scMCQ = 0.25;
  var scTF = 1.0;
  var scSA = 0.5;

  var sheetMatran = ss2.getSheetByName("matran");
  if (sheetMatran && sheetMatran.getLastRow() >= 2) {
    var dataMatran = sheetMatran.getRange(2, 1, sheetMatran.getLastRow() - 1, 19).getValues();
    for (var mIdx = 0; mIdx < dataMatran.length; mIdx++) {
      var mtRow = dataMatran[mIdx];
      var rowMtKey = (typeof supper === "function") ? supper(String(mtRow[1] || "")) : String(mtRow[1] || "").toUpperCase();
      if (rowMtKey === targetExamSupper) {
        var val6 = parseNum(mtRow[6]);
        if (val6 > 0) scMCQ = val6;
        var val10 = parseNum(mtRow[10]);
        if (val10 > 0) scTF = val10;
        var val14 = parseNum(mtRow[14]);
        if (val14 > 0) scSA = val14;
        break;
      }
    }
  }

  // 2. Lấy dữ liệu câu hỏi từ sheet nganhang
  var sheetNHRef = (typeof sheetNH !== "undefined" && sheetNH) ? sheetNH : (typeof ss !== "undefined" && ss ? ss.getSheetByName("nganhang") : null);
  if (!sheetNHRef && ss2) {
    sheetNHRef = ss2.getSheetByName("nganhang");
  }
  if (!sheetNHRef) {
    try {
      var activeSS = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSS) sheetNHRef = activeSS.getSheetByName("nganhang");
    } catch (eActive) {}
  }

  if (!sheetNHRef || sheetNHRef.getLastRow() < 2) {
    return { success: false, message: "Không tìm thấy sheet nganhang hoặc sheet không có dữ liệu!" };
  }

  var dataNH = sheetNHRef.getDataRange().getValues();
  var nhMap = {};

  for (var i = 1; i < dataNH.length; i++) {
    var idqRow = String(dataNH[i][0] || "").trim();
    if (!idqRow) continue;

    var typeqRow = String(dataNH[i][2] || "").trim().toLowerCase();
    var ansRow = null;
    
    if (typeqRow === "true-false" || typeqRow === "tf") {
      // Dùng parseTfOptions cho Cột F (Cột 5) của Ngân hàng đề
      ansRow = parseTfOptions(dataNH[i][5]); 
      if (ansRow.length === 0) ansRow = parseTfOptions(dataNH[i][6]);
    } else {
      // Chuẩn hóa chuỗi đáp án MCQ / Short-answer ở Cột G (Cột 6)
      ansRow = normalizeAns(dataNH[i][6]);
    }
    
    nhMap[idqRow] = {
      type: typeqRow,
      ans: ansRow
    };
  }

  // 3. Tiến hành chấm từng bài làm của học sinh
  var arraydiem = [];
  var arraynx = [];

  for (var k = 0; k < matchingDetails.length; k++) {
    var rawDetail = matchingDetails[k];
    var parsedStudent = parseDetailData(rawDetail);
    var listIdkq = parsedStudent.arrayIddetail || [];
    var listanswerkq = parsedStudent.arrayanswer || [];      
    var totalScore = 0;

    for (var j = 0; j < listIdkq.length; j++) {
      var targetId = String(listIdkq[j] || "").trim();
      var bankItem = nhMap[targetId];
      if (!bankItem) continue;

      var qType = String(bankItem.type || "").toLowerCase().trim();
      var ansExam = bankItem.ans;
      var ansStudent = listanswerkq[j];

      if (qType === "mcq" || qType.indexOf("phần i") !== -1 || (!qType && !Array.isArray(ansExam))) {
        if (normalizeAns(ansStudent) === normalizeAns(ansExam)) {
          totalScore += scMCQ;
        }
      } 
      else if (qType === "true-false" || qType === "tf" || qType.indexOf("phần ii") !== -1) {
        var tfStudent = parseTfOptions(ansStudent);
        var expectedTf = Array.isArray(ansExam) ? ansExam : parseTfOptions(ansExam);
        var point = pointtf(scTF, tfStudent, expectedTf);
        totalScore += point;
      } 
      else if (qType === "short-answer" || qType === "sa" || qType.indexOf("phần iii") !== -1) {
        if (normalizeAns(ansStudent) === normalizeAns(ansExam)) {
          totalScore += scSA;        
        }
      }
    }     
    
    var finalScore = Math.round(totalScore * 100) / 100;      
    var numericScore = typeof finalScore === "number" ? finalScore : (parseFloat(String(finalScore).replace(",", ".")) || 0);
    var nx = (typeof layNhanXet === "function") ? layNhanXet(numericScore) : "Hoàn thành bài thi";
    
    arraydiem.push([finalScore]); 
    arraynx.push([nx]);           
  }

  return {
    success: true,
    arraydiem: arraydiem,
    arraynx: arraynx
  };
}
// 2508ketthucsua1

// Hàm chuẩn hóa lại ngân hàng
function normalizeQuestionBank_1() {
  // Sử dụng biến ss toàn cục được khai báo ở đầu file của bạn
  var sheet = ss.getSheetByName("nganhang") || ss.getSheets()[0];
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error("Bảng tính trống hoặc không có dữ liệu để chuẩn hóa!");
  }
  
  // Đọc toàn bộ dữ liệu từ dòng 2 đến hết (bỏ qua dòng tiêu đề số 1)
  var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  var values = range.getValues();
  
  var activeCount = 0;
  var deletedCount = 0;
  var rowsToDelete = []; // Lưu lại các dòng thực tế cần xóa

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var actualRowIndex = i + 2; // Số thứ tự dòng thực tế trên Sheet
    
    // Nếu dòng trống hoàn toàn (cột ID rỗng và các cột khác không có chữ) thì bỏ qua
    if (!row[0] && row.join("").trim() === "") continue;
    
    var hasChange = false;
    
    // 1. Quét sạch tất cả các cụm <key...> hoặc </key...> lỗi trong toàn bộ các cột
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== null && row[c] !== undefined) {
        var valStr = row[c].toString();
        if (/<\/?[kK][eE][yY][^>]*>/g.test(valStr)) {
          row[c] = valStr.replace(/<\/?[kK][eE][yY][^>]*>/g, '').trim();
          hasChange = true;
        }
      }
    }
    
    // Thứ tự cột cố định:
    // 0: idquestion (A) | 1: classTag (B) | 2: type (C) | 3: part (D) | 4: question (E)
    // 5: options (F)    | 6: answer (G)   | 7: loigiai (H) | 8: date (I)
    var typeRaw = row[2] !== null ? row[2].toString().trim() : "";
    var optionRaw = row[5];
    var answerRaw = row[6];
    
    // Kiểm tra trống toàn diện
    var isOptionEmpty = checkValueEmpty(optionRaw);
    var isAnswerEmpty = checkValueEmpty(answerRaw);
    
    // 2. MỤC TIÊU 4: Nếu F rỗng và G rỗng -> XÓA NGAY dòng đó
    if (isOptionEmpty && isAnswerEmpty) {
      rowsToDelete.push(actualRowIndex);
      deletedCount++;
      continue;
    }
    
    var targetType = "";
    var targetPart = "";
    
    // 3. Phân loại chuẩn theo logic yêu cầu
    if (!isOptionEmpty && !isAnswerEmpty) {
      targetType = "mcq";
      targetPart = "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn";
    } 
    else if (!isOptionEmpty && isAnswerEmpty) {
      targetType = "true-false";
      targetPart = "PHẦN II. Câu trắc nghiệm đúng sai";
    } 
    else if (isOptionEmpty && !isAnswerEmpty) {
      targetType = "short-answer";
      targetPart = "PHẦN III. Câu trắc nghiệm trả lời ngắn";
    }
    
    // CHỈ KHI TYPE HIỆN TẠI KHÁC TYPE CHUẨN (HOẶC CÓ THẺ KEY LỖI) THÌ MỚI GHI ĐÈ
    if (typeRaw !== targetType || hasChange) {
      row[2] = targetType;
      row[3] = targetPart;
      
      sheet.getRange(actualRowIndex, 1, 1, lastColumn).setValues([row]);
      activeCount++;
    }
  }
  
  // 4. Tiến hành xóa các dòng rác (Duyệt ngược từ dưới lên để tránh bị chạy lệch index dòng)
  for (var d = rowsToDelete.length - 1; d >= 0; d--) {
    sheet.deleteRow(rowsToDelete[d]);
  }
  
  return {
    activeCount: activeCount,
    deletedCount: deletedCount
  };
}

// Hàm bổ trợ kiểm tra giá trị thực sự trống trên Google Sheets
function checkValueEmpty(val) {
  if (val === null || val === undefined) return true;
  
  var str = val.toString().trim();
  
  // Các trường hợp được coi là trống trong cấu trúc ngân hàng câu hỏi
  if (str === "" || str === "0" || str === "[]" || str === "{}" || str === "['']" || str === '[""]') {
    return true;
  }
  return false;
}
// Hàm chuẩn hóa ngân hàng câu hỏi (Tối ưu In-Memory)
function normalizeQuestionBank() {
  var sheet = ss.getSheetByName("nganhang") || ss.getSheets()[0];
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error("Bảng tính trống hoặc không có dữ liệu để chuẩn hóa!");
  }
  
  // Đọc toàn bộ dữ liệu từ dòng 2 đến hết
  var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  var values = range.getValues();
  
  var cleanedValues = []; // Mảng chứa các dòng sạch sẽ giữ lại
  var activeCount = 0;
  var deletedCount = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    
    // Nếu dòng trống hoàn toàn (cột ID rỗng và các cột khác rỗng) -> Xóa
    if (!row[0] && row.join("").trim() === "") {
      deletedCount++;
      continue;
    }
    
    // 1. Quét sạch tất cả các cụm <key...> VÀ Chuẩn hóa MathJax trên toàn bộ các cột
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== null && row[c] !== undefined) {
        var valStr = row[c].toString();
        
        // Sửa cụm <key...>
        if (/<\/?[kK][eE][yY][^>]*>/g.test(valStr)) {
          valStr = valStr.replace(/<\/?[kK][eE][yY][^>]*>/g, '').trim();
        }
        
        // Chuẩn hóa lỗi thiếu \ trong MathJax
        valStr = fixMathJaxString(valStr);
        
        row[c] = valStr;
      }
    }
    
    // Cột 5 (F): options, Cột 6 (G): answer
    var optionRaw = row[5];
    var answerRaw = row[6];
    
    var isOptionEmpty = checkValueEmpty(optionRaw);
    var isAnswerEmpty = checkValueEmpty(answerRaw);
    
    // 2. Nếu F rỗng và G rỗng -> XÓA dòng đó
    if (isOptionEmpty && isAnswerEmpty) {
      deletedCount++;
      continue;
    }
    
    // 3. Phân loại chuẩn theo logic
    var targetType = "";
    var targetPart = "";
    
    if (!isOptionEmpty && !isAnswerEmpty) {
      targetType = "mcq";
      targetPart = "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn";
    } 
    else if (!isOptionEmpty && isAnswerEmpty) {
      targetType = "true-false";
      targetPart = "PHẦN II. Câu trắc nghiệm đúng sai";
    } 
    else if (isOptionEmpty && !isAnswerEmpty) {
      targetType = "short-answer";
      targetPart = "PHẦN III. Câu trắc nghiệm trả lời ngắn";
    }
    
    // Cập nhật Type (cột 2) và Part (cột 3)
    row[2] = targetType;
    row[3] = targetPart;
    
    // Đưa dòng hợp lệ vào mảng kết quả
    cleanedValues.push(row);
    activeCount++;
  }
  
  // 4. Ghi đè lại dữ liệu sạch vào Sheet chỉ bằng 1 lần ghi
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, lastColumn).clearContent();
  
  // Ghi mảng dữ liệu mới nếu có dữ liệu
  if (cleanedValues.length > 0) {
    sheet.getRange(2, 1, cleanedValues.length, lastColumn).setValues(cleanedValues);
  }
  
  return {
    activeCount: activeCount,
    deletedCount: deletedCount
  };
}
/**
 * Hàm sửa các lỗi gõ thiếu dấu \ trong công thức MathJax/LaTeX
 */
function fixMathJaxString(str) {
  if (!str || typeof str !== 'string') return str;

  return str
    // 1. Sửa vec{a}, overrightarrow{AB}... bị thiếu \ ở đầu
    .replace(/(^|[^\\])\b(vec|overrightarrow|overleftarrow|hat|bar|tilde|dot|ddot)\{/g, '$1\\$2{')

    // 2. Sửa left{ hoặc left( ... bị thiếu \ ở left
    .replace(/(^|[^\\])\b(left|right)([\{\}\(\)\[\]\|\.\ \t])/g, '$1\\$2$3')

    // 3. Khắc phục riêng trường hợp \left... thiếu \right. hoặc thiếu dấu . ở cuối right
    // Chuyển left{ thành \left\{ nếu thiếu \ trước ngoặc nhọn
    .replace(/\\left\{/g, '\\left\\{')
    // Nếu có \right bị đứng một mình ở cuối mà không có dấu . hoặc ngoặc đi kèm -> tự động thêm \right.
    .replace(/\\right(?!\s*[\{\}\(\)\[\]\|\.])/g, '\\right.')

    // 4. Sửa các môi trường bị thiếu \ trước begin / end (ví dụ: begin{aligned}, end{cases})
    .replace(/(^|[^\\])\b(begin|end)\{/g, '$1\\$2{')

    // 5. Sửa các lệnh toán học thông dụng khác bị gõ thiếu \ ở đầu
    .replace(/(^|[^\\])\b(frac|sqrt|limits|int|sum|prod|lim|alpha|beta|gamma|delta|pi|theta|infty|le|ge|neq|approx|times|div|cdot)\b/g, '$1\\$2');
}
function layNhanXet(diem) {
  const nx1 = "🌟 Bài làm rất tốt, cần tiếp tục phát huy nhé";
  const nx2 = "👍 Bài làm tương đối tốt, nắm khá chắc kiến thức cơ bản. Cần cố gắng và chăm chỉ hơn nữa";
  const nx3 = "🙂 Bài làm mới đạt ở mức trung bình, cần cố gắng và nỗ lực nhiều hơn nữa";
  const nx4 = "⚠️ Bài làm chưa được tốt, báo động nhẹ. Cần phải xem lại thái độ học tập";
  const nx5 = "🚨 Kết quả rất báo động, có vẻ như em chưa thực sự chú ý đến việc học hoặc chưa có phương pháp học hiệu quả. Hãy chia sẻ để mình tốt hơn nhé";
  const nx6 = "🚨🚨Kết quả rất báo động, nếu còn tiếp diễn có thể em sẽ bị loại khỏi cuộc chơi";
  let nx = "";

  if (diem >= 8.5) {
    nx = nx1;
  } else if (diem > 7) {
    // 7 < diem < 8.5
    nx = nx2;
  } else if (diem >= 6) {
    // 6 <= diem <= 7
    nx = nx3;
  } else if (diem >= 5) {
    // 5 <= diem < 6
    nx = nx4;
  } else if (diem >= 3) {
    nx = nx5;
  } else {
    nx = nx6;
  }
  return nx;
}
/**
 * Hàm sắp xếp tổng quát cho Sheet
 * @param {number} cot - Thứ tự cột cần sắp xếp (Ví dụ: 2 là cột B)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Đối tượng Sheet cần sắp xếp
 */
function sapxep_1(cot, sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // Nếu có dữ liệu từ dòng 2 trở đi thì mới sắp xếp (tránh xếp mất dòng Header)
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    dataRange.sort({ column: cot, ascending: true });
  }
}
function sapxep_2(cot1, x, cot2, y, sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // Nếu có dữ liệu từ dòng 2 trở đi thì mới sắp xếp
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    
    dataRange.sort([
      { column: cot1, ascending: x === 1 },
      { column: cot2, ascending: y === 1 }
    ]);
  }
}

function parseNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  // Chuyển sang dạng chuỗi, đổi dấu phẩy thành dấu chấm rồi parse số
  var cleanVal = String(val).replace(',', '.');
  var num = parseFloat(cleanVal);
  return isNaN(num) ? 0 : num;
}

/**
 * Tách dữ liệu detail bài làm của học sinh thành 2 mảng ID và Answer tương ứng.
 * @param {string|Array} detailInput - Chuỗi JSON hoặc mảng Object từ cột detail.
 * @return {Object} Đối tượng chứa { arrayIddetail, arrayanswer }
 */
function parseDetailData(detailInput) {
  var arrayIddetail = [];
  var arrayanswer = [];

  if (!detailInput) {
    return { arrayIddetail: arrayIddetail, arrayanswer: arrayanswer };
  }

  var list = [];
  try {
    // Nếu dữ liệu truyền vào là chuỗi JSON thì parse ra mảng
    list = typeof detailInput === "string" ? JSON.parse(detailInput) : detailInput;
  } catch (e) {
    Logger.log("Lỗi parse JSON detail: " + e.toString());
    return { arrayIddetail: arrayIddetail, arrayanswer: arrayanswer };
  }

  if (Array.isArray(list)) {
    for (var i = 0; i < list.length; i++) {
      var item = list[i];
      if (item && typeof item === "object") {
        // Lấy ID câu hỏi
        var qId = item.id || item.idquestion || item.questionId || "";
        arrayIddetail.push(String(qId).trim());

        // Lấy Answer của câu hỏi
        var ans = item.answer !== undefined ? item.answer : (item.ans !== undefined ? item.ans : item.a);
        arrayanswer.push(ans);
      }
    }
  }

  return {
    arrayIddetail: arrayIddetail,
    arrayanswer: arrayanswer
  };
}

// ==========================================
/**
 * Tách dữ liệu câu hỏi trong exam_data thành 2 mảng ID và Answer.
 * @param {Array<Object|string>} examDataInput - Danh sách các câu hỏi (dạng Object hoặc mảng JSON) từ exam_data.
 * @return {Object} Đối tượng chứa { arrayexamId, arrayexamAnswer }[cite: 1]
 */
function parseExamData(examDataInput) {
  var examMap = {};

  if (!examDataInput) return examMap;

  var list = [];
  try {
    list = typeof examDataInput === "string" ? JSON.parse(examDataInput) : examDataInput;
  } catch (e) {
    Logger.log("Lỗi parse JSON exam_data: " + e.toString());
    return examMap;
  }

  // Đảm bảo list luôn là mảng để duyệt
  if (!Array.isArray(list)) list = [list];

  for (var i = 0; i < list.length; i++) {
    var q = list[i];
    if (typeof q === "string") {
      try { q = JSON.parse(q); } catch (e) { continue; }
    }

    if (q && typeof q === "object") {
      var qId = String(q.id !== undefined ? q.id : "").trim();
      var type = String(q.type || "").toLowerCase().trim();
      var part = String(q.part || "").toLowerCase().trim();
      var ansValue = null;

      // 1. Trả lời ngắn (PHẦN III / short-answer)
      if (type === "short-answer" || type === "sa" || part.indexOf("iii") !== -1) {
        ansValue = q.a !== undefined ? q.a : (q.answer !== undefined ? q.answer : q.ans);
      }
      // 2. Trắc nghiệm Đúng/Sai (PHẦN II / true-false)
      else if (type === "true-false" || type === "tf" || part.indexOf("ii") !== -1) {
        var sList = Array.isArray(q.s) ? q.s : [];
        ansValue = sList.map(function(item) {
          if (item && typeof item === "object") {
            return item.a !== undefined ? item.a : item.answer;
          }
          return item;
        });
      } 
      // 3. Trắc nghiệm MCQ (PHẦN I)
      else {
        ansValue = q.a !== undefined ? q.a : (q.answer !== undefined ? q.answer : q.ans);
      }

      if (qId) {
        examMap[qId] = ansValue;
      }
    }
  }

  return examMap;
}


/**
 * Hàm tách riêng danh sách phương án lựa chọn (options) từ exam_data.
 * @param {Array<Object|string>|string} examDataInput - Dữ liệu câu hỏi trong exam_data.
 * @return {Array<Array>} Mảng chứa danh sách options của từng câu hỏi.
 */
/**
 * Hàm trích xuất mảng giá trị Boolean [true/false] từ danh sách mệnh đề Đúng/Sai.
 * @param {Array<Object>|string} tfInput - Mảng chứa các mệnh đề dạng [{"text": "...", "a": true}, ...] hoặc chuỗi JSON.
 * @return {Array<boolean>} Mảng Boolean tương ứng [true, false, ...]
 */
function parseTfOptions(tfInput) {
  var arrayTfBool = [];

  if (!tfInput) return arrayTfBool;

  var list = [];
  try {
    list = typeof tfInput === "string" ? JSON.parse(tfInput) : tfInput;
  } catch (e) {
    Logger.log("Lỗi parse JSON True-False: " + e.toString());
    return arrayTfBool;
  }

  if (Array.isArray(list)) {
    for (var i = 0; i < list.length; i++) {
      var item = list[i];
      if (item && typeof item === "object") {
        var val = item.a !== undefined ? item.a : item.answer;
        // Ép kiểu chuẩn về Boolean
        var boolVal = (val === true || String(val).toLowerCase() === "true" || val === 1);
        arrayTfBool.push(boolVal);
      }
    }
  }

  return arrayTfBool; // Trả thẳng mảng [true, false, false, true]
}

// ==========================================
// Test thử với dữ liệu của thầy
// ==========================================
function testParseTf() {
  var dataInput = [
    {"text":"Giá trị trung bình của mẫu số liệu bảng $1$ xấp xỉ bằng $58,7$ (làm tròn đến hàng phần chục).","a":true},
    {"text":"Ghép nhóm mẫu số liệu bảng $1$ với lớp ghép đầu tiên là $[25; 34)$ ta có mốt của mẫu số liệu sau khi ghép nhóm (làm tròn kết quả đến hàng phần mười) bằng $49,8$.","a":true},
    {"text":"Bảng tần số ghép nhóm của mẫu số liệu bảng $1$ có lớp ghép đầu tiên là $[25; 34)$ như bảng $2$ sau:","a":false},
    {"text":"Kích thước mẫu của mẫu số liệu bảng $1$ bằng $30$.","a":true}
  ];

  var result = parseTfOptions(dataInput);
  Logger.log(result); 
  // Kết quả ra đúng chuẩn: [true, true, false, true]
}

/**
 * Tính điểm câu hỏi Trắc nghiệm Đúng/Sai dựa trên số ý trả lời đúng.
 * @param {number} point - Điểm tối đa của câu hỏi.
 * @param {Array<boolean>} p1 - Mảng đáp án 1 (ví dụ: đáp án học sinh).
 * @param {Array<boolean>} p2 - Mảng đáp án 2 (ví dụ: đáp án chuẩn).
 * @return {number} Số điểm đạt được (đã làm tròn 2 chữ số thập phân).
 */
function pointtf(point, p1, p2) {
  // Kiểm tra dữ liệu đầu vào
  if (!Array.isArray(p1) || !Array.isArray(p2) || p1.length === 0 || p2.length === 0) {
    return 0;
  }

  // Đếm số ý trùng khớp giữa p1 và p2
  var count = 0;
  var len = Math.min(p1.length, p2.length);

  for (var i = 0; i < len; i++) {
    // Ép kiểu Boolean để so sánh chuẩn xác tuyệt đối (tránh lỗi do chuỗi "true"/1)
    var bool1 = (p1[i] === true || String(p1[i]).toLowerCase() === "true" || p1[i] === 1);
    var bool2 = (p2[i] === true || String(p2[i]).toLowerCase() === "true" || p2[i] === 1);

    if (bool1 === bool2) {
      count++;
    }
  }

  // Bảng tỷ lệ điểm tương ứng theo số ý đúng [0, 1, 2, 3, 4]
  var scaleMap = [0, 0.1, 0.25, 0.5, 1];

  // Lấy tỷ lệ (nếu vượt quá 4 ý thì mặc định lấy 1)
  var scale = scaleMap[count] !== undefined ? scaleMap[count] : 1;

  // Tính điểm và làm tròn 2 chữ số thập phân
  var resultPoint = Number(point) * scale;
  return Math.round(resultPoint * 100) / 100;
}

// ==========================================
// Test thử theo ví dụ
// ==========================================
function testPointTf() {
  var p1 = [true, true, false, true];
  var p2 = [true, false, true, true];

  var result = pointtf(1, p1, p2);
  Logger.log("Điểm đạt được: " + result); // Kết quả trả về số 0.25
}

function demoSuDung() {
  // 1. Giả sử đây là dữ liệu ngân hàng câu hỏi từ exam_data
  var examDataInput = [
    { "id": 3, "part": "PHẦN I", "type": "mcq", "a": "$12$" },
    { "id": 13, "part": "PHẦN II", "type": "true-false", "s": [{ "a": true }, { "a": false }, { "a": true }, { "a": false }] },
    { "id": 17, "part": "PHẦN III", "type": "short-answer", "a": "45" }
  ];

  // 2. Gọi hàm parseExamData và lưu vào biến result
  var result = parseExamData(examDataInput);

  // 3. Lấy mảng ID và mảng Answer ra sử dụng
  var listId = result.arrayexamId;       // Mảng ID: ["3", "13", "17"]
  var listAns = result.arrayexamAnswer;  // Mảng Answer: ["$12$", [true, false, true, false], "45"]

  // --- Ví dụ: Gọi phần tử đầu tiên (Phần tử chỉ số 0) ---
  Logger.log("ID câu thứ 1: " + listId[0]);      // Ra: "3"
  Logger.log("Đáp án câu thứ 1: " + listAns[0]); // Ra: "$12$"

  // --- Ví dụ: Gọi phần tử câu Đúng/Sai (Phần tử chỉ số 1) ---
  Logger.log("ID câu thứ 2: " + listId[1]);      // Ra: "13"
  Logger.log("Đáp án câu thứ 2: " + listAns[1]); // Ra: [true, false, true, false]
}


function normalizeAns(str) {
  if (str === null || str === undefined) return "";
  
  var s = String(str).trim();
  
  // Bỏ dấu nháy bọc ngoài nếu có (ví dụ: "\"$E \\subset F$.\"")
  if ((s.startsWith('"') && s.endsWith('"')) || (s.startsWith("'") && s.endsWith("'"))) {
    s = s.slice(1, -1).trim();
  }
  
  return s
    .replace(/\\\\/g, "\\")      // Chuyển \\ thành \ (đồng bộ LaTeX)
    .replace(/\s+/g, " ")        // Gom nhiều khoảng trắng/xuống dòng thành 1 khoảng trắng
    .trim()                      // Xóa khoảng trắng đầu/cuối
    .replace(/\.$/, "")          // Xóa duy nhất dấu chấm thừa ở cuối câu
    .trim()                      // Trim lại phòng trường hợp còn khoảng trắng trước dấu chấm
    .toLowerCase();
}
