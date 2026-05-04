// --- FILE TỔNG TRÊN GITHUB ---

function mainDoGet(e) {
const params = e.parameter;
  const type = params.type;
  const action = params.action || e.parameter.action;  
  
  
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
    const passAdmiN = sheet.getRange("H2").getValue().toString().trim();

    // Trả về trực tiếp chuỗi pass, không join gì cả
    return ContentService.createTextOutput(passAdmiN).setMimeType(ContentService.MimeType.TEXT);
  }
  // Ghi idinput vào sheet(danhsach)
  // Thêm vào trong hàm doGet(e) của thầy
// Ghi idinput vào sheet(danhsach)
  if (action === "getLastID") {  
    // Thêm dòng này để định nghĩa 'sheet' là sheet danhsach
    var sheet = ss.getSheetByName("danhsach"); 
    var val = sheet.getRange("J2").getValue();
    return ContentService.createTextOutput(val.toString());
  }

  if (action === "saveLastID") {
    var idMoi = e.parameter.id;  
    // Thêm dòng này để định nghĩa 'sheet' là sheet danhsach
    var sheet = ss.getSheetByName("danhsach"); 
    sheet.getRange("J2").setValue("'" + idMoi); 
    SpreadsheetApp.flush(); 
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
  // 6. XÁC MINH THÍ SINH
  if (type === 'verifyStudent') {
    const idNumber = params.idnumber;
    const sbd = params.sbd;
    const sheet = ss.getSheetByName("danhsach");
    const lastRow = sheet.getLastRow();
    // #vip
    if (lastRow < 2) {
    return createResponse("error", "Danh sách thí sinh trống!");
      }      
    // #vip
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const key = supper(sbd + "." + idNumber);      
    for (let i = 0; i < data.length; i++) {
      if ((data[i][7] || "").toString().trim() === key) {
        return createResponse("success", "OK", {
          name: data[i][1], 
          class: data[i][2], 
          limit: data[i][3],
          limittab: data[i][4], 
          taikhoanapp: data[i][6], 
          idnumber: idNumber, 
          sbd: sbd
        });         
      }
    }
    return createResponse("error", sbd + " không tồn tại!");
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
    var idTraCuu = supper(params.id);
    if (!idTraCuu) return ContentService.createTextOutput("Thiếu ID rồi!").setMimeType(ContentService.MimeType.TEXT);
    const lastRow = sheetNH.getLastRow();
   
        if (lastRow < 2) {
    return createResponse("error", "Ngân hàng đang trống!");
      }  
    var data = sheetNH.getRange(2, 1, lastRow - 1, 8).getValues();   
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString().trim() === idTraCuu) {
        var loigiai = data[i][7] || "";

        // Ép kiểu về String để đảm bảo không bị lỗi tệp
        return ContentService.createTextOutput(String(loigiai))
          .setMimeType(ContentService.MimeType.TEXT);
      }
    }
    return ContentService.createTextOutput("Không tìm thấy ID này!").setMimeType(ContentService.MimeType.TEXT);
  }
  // SỬ LÝ CÂU TRÙNG
  if (action == 'findDuplicateQuestions') {
    return ContentService.createTextOutput(JSON.stringify(findDuplicateQuestions()))
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
        return createResponse("success", "OK", {
          idquestion: dataNH[i][0], 
          classTag: dataNH[i][1], 
          type: dataNH[i][2],
          question: dataNH[i][4],
          options: dataNH[i][5],
          answer: dataNH[i][6],
          loigiai: dataNH[i][7],
          datetime: dataNH[i][8]
          
        });
      }
    }
    return resJSON({ status: 'error' });
  }

  // 8. LẤY MA TRẬN ĐỀ
  if (type === 'getExamCodes') {
    const teacherId = supper(params.idnumber);
    const sheet = ss.getSheetByName("matran");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
    return createResponse("error", "Ma trận trống!");
      }  

    const data = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
    const results = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0].toString().trim() === teacherId || row[0].toString() === "SYSTEM") {
        try {
          results.push({
            code: row[1].toString(), name: row[2].toString(), topics: JSON.parse(row[3]),
            fixedConfig: {
              duration: parseInt(row[4]), numMC: JSON.parse(row[5]), scoreMC: parseFloat(row[6]),
              mcL3: JSON.parse(row[7]), mcL4: JSON.parse(row[8]), numTF: JSON.parse(row[9]),
              scoreTF: parseFloat(row[10]), tfL3: JSON.parse(row[11]), tfL4: JSON.parse(row[12]),
              numSA: JSON.parse(row[13]), scoreSA: parseFloat(row[14]), saL3: JSON.parse(row[15]), saL4: JSON.parse(row[16])
            }
          });
        } catch (err) {}
      }
    }
    return createResponse("success", "OK", results);
  }

  // 9. LẤY TẤT CẢ CÂU HỎI (Hàm này thầy bị trùng, em gom lại bản chuẩn nhất)
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

    var qObj = {
      id: rows[i][0],
      classTag: rows[i][1] || "",
      type: rows[i][2] || "",
      part: rows[i][3] || "",
      question: rows[i][4] || "",
      a: rows[i][6] || "",
      loigiai: rows[i][7] || ""
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

  const sheet = ss.getSheetByName("exam_data");
  const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
    return createResponse("error", "Ngân hàng câu hỏi trống!");
      }  
  const examCodeInput = e.parameter.examCode || "";
  const questionIdInput = e.parameter.questionId || "";
  const idgv = e.parameter.idgv || "";
  const key = supper(examCodeInput + "." + questionIdInput + "." + idgv);

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  for (let i = 0; i < data.length; i++) {    
    if (supper(data[i][9] || "") === key) {

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
    const examCode = params.examCode;
    const idgv = params.idgv;
    const key = supper(examCode + "." + idgv);
    const sheet = ss.getSheetByName("exam_data");
    if (!sheet) return createResponse("error", "Chưa có dữ liệu exam_data");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
    return createResponse("error", "Ngân hàng trống!");
      }  
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const results = [];

    for (let i = 0; i < data.length; i++) {      
      if (supper(data[i][8] || "") === key) {
        try {
          // Cột E chứa JSON câu hỏi
          results.push(JSON.parse(data[i][4]));
        } catch (err) {
          results.push(data[i][4]);
        }
      }
    }
    return createResponse("success", "OK", results);
  }

// #04 chung
  // Tải điểm
if (action === "downloadScores") {
    const idgv = e.parameter.idgv; 
    const exams = e.parameter.exams;     
    const sheet = ss.getSheetByName("ketqua");
    
    const keycheck = (exams + "." + idgv).toUpperCase();
    
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
      .filter(row => String(row[9]).toUpperCase() === keycheck) // Lọc dựa trên cột I (index 9)
      .map(row => row.slice(0, 9)); // Cắt bỏ cột J, chỉ lấy từ cột A (index 0) đến I (index 8)
    
    return ContentService.createTextOutput(JSON.stringify({
      header: header,
      data: filteredData
    })).setMimeType(ContentService.MimeType.JSON);
}
  
// ===== LẤY LIST EXAMS =====
  if (action === "getExamsList") {
    return getExamsList(e.parameter.type, e.parameter.idgv );
  }

  // ===== RESET DATA =====
  if (action === "resetData") {
    const key = supper(e.parameter.password + "." + e.parameter.idgv );
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
    return resetData(
      e.parameter.type,
      e.parameter.password,
      e.parameter.mode,
      e.parameter.exams,
      e.parameter.idgv
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

  // 3. TOP 10
  if (type === 'top10') {
    const sheet = ssAdmin.getSheetByName("Top10Display");
    if (!sheet) return createResponse("error", "Không tìm thấy sheet Top10Display");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return createResponse("success", "Chưa có dữ liệu Top 10", []);
    const values = sheet.getRange(2, 1, Math.min(10, lastRow - 1), 10).getValues();
    const top10 = values.map((row, index) => ({
      rank: index + 1, name: row[0], phoneNumber: row[1], score: row[2],
      time: row[3], sotk: row[4], bank: row[5], idPhone: row[9]
    }));
    return createResponse("success", "OK", top10);
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
    const data = JSON.parse(e.postData.contents || "{}");
    const action = (data.action || e.parameter.action || "").toString();
    const res = (status, message, payload) =>
      ContentService.createTextOutput(
        JSON.stringify({ status, message, data: payload || null })
      ).setMimeType(ContentService.MimeType.JSON);
   const sheetKq = ss.getSheetByName("ketqua") || ss.insertSheet("ketqua");

  // Đảm bảo tiêu đề cột chuẩn nếu sheet mới tạo
  if (sheetKq.getLastRow() === 0) {
    sheetKq.appendRow(["Timestamp", "exams", "sbd", "name", "class", "tongdiem", "time", "idgv", "vipham", "exams.idgv", "exams.sbd.idgv"]);
  }

  // LOGIC CHUNG CHO CẢ 2 LOẠI (Vì cấu trúc cột ghi là giống nhau)
  // Tìm đến đoạn xử lý kết quả và thay bằng đoạn này:
// Thay thế đoạn xử lý submit trong mainDoPost
if (action === "submitExam" || action === "submitExamMatrix") {
  try {
    // 1. Phải đảm bảo ssTarget đã được khai báo bằng openById ở trên
   
    
    let sheetKq = ss.getSheetByName("ketqua");    
    if (!sheetKq) {
      sheetKq = ss.insertSheet("ketqua");
      sheetKq.appendRow(["Timestamp", "Mã đề", "SBD", "Họ tên", "Lớp", "Tổng điểm", "Thời gian làm", "IDGV", "Vi phạm", "exams.idgv", "exams.sbd.idgv"]);
    }

    // 2. CHUẨN HÓA DỮ LIỆU
    const exams = (data.exams || data.examCode || "").toString().toUpperCase();
    const idgv = (data.idgv || "").toString();
    const diem = data.tongdiem !== undefined ? data.tongdiem : (data.score || 0);
    const className  = (data.class || data.className || "Tự do").toString();
    const thoiGian = data.time || 0;
    const sbd = data.sbd || "";
    const tabCount = data.tabSwitches !== undefined ? data.tabSwitches : 0;

    // 3. TÌM HÀNG TRỐNG TIẾP THEO (Ép ghi thay vì dùng appendRow)
    // const lastRow = sheetKq.getLastRow();
    // const nextRow = lastRow + 1;
     const vals = sheetKq.getDataRange().getValues();
      let nextRow = -1;
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][1].trim() === "") {
          nextRow = i + 1; break;
        }
      }
    // NẾU KHÔNG TÌM THẤY HÀNG TRỐNG THÌ GHI VÀO DÒNG CUỐI CÙNG TIẾP THEO
    if (nextRow === -1) {
      nextRow = sheetKq.getLastRow() + 1;
    }

    // Đảm bảo nextRow không bao giờ nhỏ hơn 2 (tránh ghi đè tiêu đề hàng 1)
    if (nextRow < 2) nextRow = 2;
    // Chuẩn bị mảng dữ liệu 1 hàng
    const rowData = [
      data.timestamp || new Date().toLocaleString('vi-VN'), // A
      supper(exams),                                                // B
      supper(sbd),                                          // C
      supper(data.name || ""),                             // D
      supper(className),                                                 // E
      diem,                                                // F
      thoiGian,                                           // G           
      "'" + supper(idgv),
      tabCount,
      supper(exams + "." + idgv),                                    // S
      supper(exams + "." + sbd + "." + idgv)      
    ];

    // GHI ĐÈ VÀO RANGE CỤ THỂ
    sheetKq.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
   
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "success", 
      message: "Ghi điểm thành công vào file TOÁN!",
      rowRecorded: nextRow // Trả về số hàng đã ghi để thầy check bên Console React
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: "Lỗi ghi điểm: " + err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

    // 2. Nếu sau này thầy gửi dữ liệu đăng ký (có pass, phone...)
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
          return resJSON({ status: "success" });
        }
      }
      return resJSON({ status: "error", message: "ID hoặc Mật khẩu GV không đúng!" });
    }
// 6. XÁC MINH ADMIN (verifyAdmin)
    if (action === "verifyAdmin") {      
      if (data.password.toString().trim() === passAdmin) return resJSON({ status: "success", message: "Chào Admin!" });
      return resJSON({ status: "error", message: "Sai mật khẩu!" });
    }
// #06 Ma trận
// 1. NHÁNH LỜI GIẢI (saveLG)
    if (action === 'saveLG') {
      var lastRow = sheetNH.getLastRow();


      if (lastRow < 2) return ContentService.createTextOutput("⚠️ Sheet rỗng, chưa có ID để khớp thầy ơi!").setMimeType(ContentService.MimeType.TEXT);

      // 1. Tìm ô trống đầu tiên ở cột H
      var eValues = sheetNH.getRange(1, 8, lastRow, 1).getValues();
      var firstEmptyRow = 0;
      for (var i = 1; i < eValues.length; i++) {
        if (!eValues[i][0] || eValues[i][0].toString().trim() === "") {
          firstEmptyRow = i + 1;
          break;
        }
      }
      if (firstEmptyRow === 0) firstEmptyRow = lastRow + 1;

      // 2. Điền LG và ép ID theo cột A
      var count = 0;
      data.forEach(function (item, index) {
        var targetRow = firstEmptyRow + index;

        // Lấy ID "xịn" đang nằm ở cột A của hàng này
        var realId = sheetNH.getRange(targetRow, 1).getValue().toString();

        if (realId) {
          var rawLG = item.loigiai || item.lg || "";

          // Dùng Regex để tìm "id: ..." hoặc "id:..." và thay bằng ID xịn từ cột A
          // Đoạn này xử lý cả trường hợp có ngoặc kép hoặc không
          var fixedLG = rawLG.replace(/id\s*:\s*["']?[^"'\s]+["']?/g, 'id: "' + realId + '"');

          // Ghi vào cột H
          sheetNH.getRange(targetRow, 8).setValue(fixedLG);
          count++;
        }
      });
      sheetNH.getRange("D:H").setWrap(true);

      // Tự chỉnh chiều cao từ dòng 2 trở xuống
      

      return ContentService.createTextOutput("🚀 Đã xong! Điền tiếp " + count + " lời giải. ID trong LG đã được đồng bộ theo ID câu hỏi.").setMimeType(ContentService.MimeType.TEXT);
    }
    // 2. NHÁNH MA TRẬN (saveMatrix)
    if (action === "saveMatrix") {
      
      const sheetMatran = ss.getSheetByName("matran") || ss.insertSheet("matran");
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
        supper(toStr(data.makiemtra)), 
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
        supper(toStr(data.makiemtra) + "." + toStr(data.gvId))
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
        if (vals[i][0].toString() === toStr(data.gvId) && vals[i][1].toString() === toStr(data.makiemtra)) {
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

      var startRow = sheetNH.getLastRow() + 1;

      var rows = data.map(function (item) {
        return [
          item.id,
          item.classTag,
          item.type,
          item.part,
          item.question,
          item.options || "",
          item.answer || "",
          item.loigiai || "",
          now
        ];
      });

      if (rows.length > 0) {
        sheetNH.getRange(startRow, 1, rows.length, rows[0].length)
          .setValues(rows);
      }
      var lastRow = sheetNH.getLastRow();
      sheetNH.getRange("D:H").setWrap(true);

      // Tự chỉnh chiều cao từ dòng 2 trở xuống     

      return createResponse("success", "Đã lưu " + rows.length + " câu hỏi thành công!");
    }
    // 5. CẬP NHẬT CÂU HỎI (updateQuestion)
    if (action === 'updateQuestion') {
  var item = data.data;
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

// #07 Thi lẻ

    // =================================================================== TRỘN ĐỀ ===========================================

    if (action === "studentGetExam") {
      try {
        const sbd = data.sbd ? data.sbd : "";
        const examCode = data.examCode ? data.examCode : "";
        const idgv = data.idgv ? data.idgv : "";
        const keyds = supper(sbd + "." + idgv);
        const keyexams = supper(examCode + "." + idgv);
        const keysbd = supper(examCode + "." + sbd + idgv);

        const sheetDS = ss.getSheetByName("danhsach");
        const sheetData = ss.getSheetByName("exam_data");
        const sheetExam = ss.getSheetByName("exams");
        const sheetKQ = ss.getSheetByName("ketqua"); // Bảng lưu kết quả thi
        const dataDS = sheetDS.getDataRange().getValues();        
        if (dataDS.length < 2) {
          return createResponse("error", "Danh sách thí sinh trống!");
      }    

       // 1. Tìm học sinh bằng vòng lặp (An toàn và nhanh nhất)
var student = null;
for (var i = 1; i < dataDS.length; i++) {
  var rowSBD = supper(dataDS[i][7] || "");  
  
  // So sánh chuẩn cả 2 điều kiện
  if (rowSBD === keyds) {
    student = dataDS[i];
    break; // Tìm thấy rồi thì thoát vòng lặp luôn
  }
}

// 2. Kiểm tra nếu không tìm thấy
if (!student) {
  return createResponse("error", "SBD hoặc IDGV không chính xác!");
}
const exRow = sheetExam.getDataRange().getValues().find(r => 
  supper(r[14]) === keyexams
);
        if (!exRow) return createResponseW("error", "Không tìm thấy mã đề: " + examCode + " / GV: " + idgv);
        // ===== CHECK THỜI GIAN MỞ / ĐÓNG =====
const now = new Date();

const openTime = exRow[12] instanceof Date 
  ? exRow[12] 
  : new Date(exRow[12]);

const closeTime = exRow[11] instanceof Date 
  ? exRow[11] 
  : new Date(exRow[11]);

        // --- BỔ SUNG: CHẶN SỐ LẦN THI ---
        // Cột N là index 13. Lấy số lần thi tối đa cho phép.
        const maxAttempts = parseInt(exRow[13], 10) || 1;
        let exRowKq = [];

        if (sheetKQ.getLastRow() > 1) {
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
        // chuẩn hóa
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

        // 2. Lấy câu hỏi - ĐOẠN ĐÃ TỐI ƯU
        const allRows = sheetData.getDataRange().getValues();
        const filteredQuestions = allRows.slice(1)
          .filter(r => supper(r[8]) === keyexams)
          .map(r => {
            let raw = r[4];
            if (!raw) return null;

            // Thay thế đoạn từ dòng 130 đến 135 bằng đoạn này:
                let contentStr = raw.toString().trim();
                    try {
                        // Ưu tiên 1: Parse trực tiếp dữ liệu chuẩn
                return JSON.parse(contentStr);
                  } catch (e) {
                  // Ưu tiên 2: Chỉ xử lý nếu JSON thực sự có vấn đề về dấu gạch chéo (Escape)
                      try {
                    // Chỉ nhân đôi dấu gạch chéo nếu cần thiết, không dùng Regex xóa ký tự ẩn
                       let fixed = contentStr.replace(/\\/g, "\\\\").replace(/\\\\"/g, "\\\"");
                      return JSON.parse(fixed);
                        } catch (e2) {
                        // Ưu tiên 3: Trả về object lỗi để không làm treo app
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

        // 3. Trả về (Em bỏ qua bước trộn để test xem nó có lên đủ câu không đã)
        return createResponseW("success", "OK", {
          studentName: student[1],
          studentClass: student[2],
          duration: toInt(exRow[8], 33),
          minSubmitTime: toInt(exRow[9], 0),     // minitime
          maxTabSwitches: toInt(exRow[10], 3),        // tab limit
          maxthi: maxAttempts,
          deadline: Utilities.formatDate(closeTime, "GMT+7", "yyyy/MM/dd HH:mm"),
          openTime: Utilities.formatDate(openTime, "GMT+7", "yyyy/MM/dd HH:mm"),
          scoreMCQ: toFloat(exRow[3], 0),
          scoreTF: toFloat(exRow[5], 0),
          scoreSA: toFloat(exRow[7], 0),

          questions: filteredQuestions // Gửi hết về xem có đủ không
        });

      } catch (error) {
        return createResponseW("error", "Lỗi GAS: " + error.toString());
      }
    }
    if (action === 'saveOnlySolutions') {
      const sheet = ss.getSheetByName("exam_data");
      if (!sheet) return createResponse("error", "Không tìm thấy sheet!");

      const lastRow = sheet.getLastRow();
      const solutions = data.solutions; // Mảng các chuỗi {...}
      const examCode = data.examCode;
      const idgv = data.idgv;
      

      // Đọc dữ liệu để làm bản đồ
      const range = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      let updatedCount = 0;

      solutions.forEach(solText => {
        // 1. Thử tìm ID trong khối lời giải
        const idMatch = solText.match(/id\s*:\s*"?([\w.]+)"?/);
        let found = false;

        if (idMatch) {
          const solId = idMatch[1].toString();
          const key = supper(examCode + "." + solId + "." + idgv)
          // Dò đúng dòng có Mã đề + ID
          for (let i = 1; i < range.length; i++) {
           
            if (range[i][9].toString() === key) {
              sheet.getRange(i + 1, 6).setValue(solText);
              range[i][5] = solText; // Cập nhật vào mảng tạm để tránh ghi đè
              updatedCount++;
              found = true;
              break;
            }
          }
        }  // ######

        // 2. Nếu không có ID hoặc không tìm thấy dòng khớp ID -> Tìm dòng trống đầu tiên của mã đề đó
        if (!found) {
          for (let i = 1; i < range.length; i++) {
            if (range[i][0].toString() === examCode.toString() && (!range[i][5] || range[i][5].toString().trim() === "")) {
              sheet.getRange(i + 1, 6).setValue(solText);
              range[i][5] = solText; // Đánh dấu là đã điền
              updatedCount++;
              found = true;
              break;
            }
          }
        }
      });
            sheet.getRange("D:H").setWrap(true);
      // Tự chỉnh chiều cao từ dòng 2 trở xuống   

      return createResponse("success", `Đã nạp xong ${updatedCount} lời giải cho mã ${examCode}!`);
    }



    // 2. NHÁNH NẠP CÂU HỎI (Khớp 100% với React ở trên)
    if (action === "saveOnlyQuestions") {
  const sheet = ss.getSheetByName("exam_data") || ss.insertSheet("exam_data");
  const qArray = data.questions;
  const examCode = data.examCode;
  const idgv = data.idgv;
  const force = data.force || false; 
  
  if (!Array.isArray(qArray)) return createResponse("error", "questions không phải mảng!");

  const fullData = sheet.getDataRange().getValues();

  // --- LOGIC MỚI: KIỂM TRA NẾU LÀ SỬA CÂU LẺ (Mảng chỉ có 1 phần tử) ---
  if (qArray.length === 1 && !force) {
    const targetId = qArray[0].id.toString();
    let rowIdx = -1;

    // Tìm xem ID câu hỏi này đã nằm ở dòng nào của Mã đề này chưa
    for (let i = 0; i < fullData.length; i++) {
      if ((fullData[i][0] || "").toString() === examCode.toString() && (fullData[i][1] || "").toString() === targetId) {
        rowIdx = i + 1;
        break;
      }
    }

    // Nếu tìm thấy dòng cũ, tiến hành ghi đè đúng dòng đó
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
  // --- LOGIC CŨ CỦA THẦY: LƯU CẢ BỘ ---
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

      // Tự chỉnh chiều cao từ dòng 2 trở xuống
      
  return createResponse("success", `Đã nạp ${rows.length} câu vào mã ${examCode}`);
}


    // 1. LƯU CẤU HÌNH (Ghi về Spreadsheet của GV) =========================================================================
    if (action === "saveExamConfig") {
      const sheetExamsGV = ss.getSheetByName("exams") || ss.insertSheet("exams");
      const examCode = (data.examCode || "").toString().trim();
      const idgv = (data.idgv || "").toString().trim();
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

      // Lấy force từ data (Body JSON)
      const isForce = data.force === true || data.force === "true";

      const vals = sheetExamsGV.getDataRange().getValues();
      let existingRow = -1;
      // Dò tìm mã đề
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][14] && vals[i][14].toString().trim() === keyId) {
          existingRow = i + 1;
          break;
        }
      }

      // Nếu tìm thấy mã đề mà KHÔNG chọn ghi đè thì mới trả về "exists"
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
        keyId
      ];
      
      if (existingRow !== -1) {
        // THỰC HIỆN GHI ĐÈ tại đây
        sheetExamsGV.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
        return createResponse("success", "✅ Đã GHI ĐÈ cấu hình đề " + examCode);
      } else {
        // THÊM MỚI
        sheetExamsGV.appendRow(rowData);
        return createResponse("success", "✅ Đã lưu cấu hình mới cho đề " + examCode);
      }
    }
// 7. LƯU TỪ WORD (uploadWord)
    if (action === "uploadWord") {
      const sheetExams = ss.getSheetByName("Exams") || ss.insertSheet("Exams");
      const sheetBank = ss.getSheetByName("QuestionBank") || ss.insertSheet("QuestionBank");
      sheetExams.appendRow([data.config.title, data.idNumber, data.config.duration, data.config.minTime, data.config.tabLimit, JSON.stringify(data.config.points)]);
      data.questions.forEach(function (q) { sheetBank.appendRow([data.config.title, q.part, q.type, q.classTag, q.question, q.answer, q.image]); });
      return createResponse("success", "UPLOAD_DONE");
    }

// #08 Chung
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

  var topics = [];
  var classesMap = {}; // Dùng để lọc danh sách lớp không trùng lặp

  // Chạy từ dòng 2 (bỏ tiêu đề)
  for (var i = 1; i < dataCD.length; i++) {
    var lop = dataCD[i][0];   // Cột A: lop
    var idcd = dataCD[i][1];  // Cột B: idcd
    var namecd = dataCD[i][2]; // Cột C: namecd
    var total = dataCD[i][8]; // cột I ghi tổng số câu

    if (lop) {
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
    classes: Object.keys(classesMap).sort(function (a, b) { return a - b; }) // Trả về [9, 10, 11, 12] chẳng hạn
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
// tìm câu trùng=========================================================================================================================================
function findDuplicateQuestions() {  
  const sheet = ss.getSheetByName("nganhang"); 
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1); // Bỏ dòng tiêu đề
  
  const results = [];
  const processedIdx = new Set();

  for (let i = 0; i < rows.length; i++) {
    if (processedIdx.has(i)) continue;
    
    let group = { 
      mainId: rows[i][0], 
      score: 0, 
      items: [getRowObj(rows[i], headers, i + 2)] 
    };
    
    for (let j = i + 1; j < rows.length; j++) {
      if (processedIdx.has(j)) continue;
      
      let score = calculateSimilarity(rows[i], rows[j]);
      
      if (score >= 50) { 
        group.items.push(getRowObj(rows[j], headers, j + 2));
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

  // 2. Options (30%) - Parse và so sánh không cần thứ tự
  try {
    const o1 = JSON.parse(q1[5] || "[]").sort().join('|');
    const o2 = JSON.parse(q2[5] || "[]").sort().join('|');
    if (o1 !== "" && o1 === o2) score += 30;
  } catch(e) {}

  // 3. Question (40%) - Xóa khoảng trắng và chữ hoa/thường
  const txt1 = String(q1[4]).replace(/\s+/g, '').toLowerCase();
  const txt2 = String(q2[4]).replace(/\s+/g, '').toLowerCase();
  if (txt1 !== "" && txt1 === txt2) score += 40;

  // 4. ClassTag (5%) - So sánh mã xyzt (4 số đầu)
  const tag1 = String(q1[1]).substring(0, 4);
  const tag2 = String(q2[1]).substring(0, 4);
  if (tag1 !== "" && tag1 === tag2) score += 5;

  if (score >= 95) return 99;
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
// lọc mã exems chung
function getExamsList(type, idgv) {

  let sheetName;
  let columnIndex;

  if (type === "ketqua") {
    sheetName = "ketqua";
    columnIndex = 10; // cột I
  }

  else if (type === "matran") {
    sheetName = "matran";
    columnIndex = 19; // cột S
  }

  else if (type === "exams") {
    sheetName = "exams";
    columnIndex = 16; // cột P
  }

  else if (type === "exam_data") {
    sheetName = "exam_data";
    columnIndex = 9; // cột I
  }

  else {
    return createResponse("error", "Type không hợp lệ");
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return createResponse("error", "Không tìm thấy sheet " + sheetName);
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return createResponse("success", "OK", []);
  }

  const examsColumn = sheet
    .getRange(2, columnIndex, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(v => v && v !== "");

  const unique = [...new Set(examsColumn)];

  return createResponse("success", "OK", unique);
}
// Reset chung
function resetData(type, password, mode, exams, idgv) {  
  // Chuẩn hóa mã trước khi so sánh
  const keyid = N9(idgv);
  const idgvStr = idgv.toString().trim();
  const exam = exams.split(".")[0];
  const keyexamsid = supper(exam + "." + idgvStr);

  let sheetName = "";
  let colums = 0;

  // 1. Xác định Sheet và Cột mốc
  if (type === "ketqua") {
    sheetName = "ketqua";
    colums = 8;
  }
  else if (type === "matran") {
    sheetName = "matran";
    colums = 18;
  }
  else if (type === "exams") {
    sheetName = "exams";
    colums = 15;
  }
  else if (type === "exam_data") {
    sheetName = "exam_data";
    colums = 8;
  }
  else return createResponse("error", "Loại dữ liệu (Type) không hợp lệ");

  let rowsDeleted = 0;

  // 2. Xử lý xóa theo MODE
  if (mode === "all") {
    // Xóa toàn bộ theo IDGV (cột colums)
    rowsDeleted = deleteFastAll(keyid, colums, sheetName);
    return createResponse("success", "Đã dọn sạch " + rowsDeleted + " dòng trong sheet " + sheetName);
  }

  if (mode === "byExams") {
    if (!exams) return createResponse("error", "Thiếu mã bài tập (exams)");
    
    // Xóa theo mã cụ thể (cột colums + 1)
    rowsDeleted = deleteFast(keyexamsid, colums + 1, sheetName);
    return createResponse("success", "Đã xóa " + rowsDeleted + " dòng của  " + exams + " (" + sheetName + ")");
  }

  // 3. Nếu không rơi vào 2 mode trên
  return createResponse("error", "Chế độ (Mode) không hợp lệ");
}
// =============================================================Kết thúc Reset chung=========================================================================

// xem điểm
function getScore(e) {
  const sbd = e.parameter.sbd;
  const exams = e.parameter.exams;
  const idgv = e.parameter.idgv;
  const key = supper(exams + "." + sbd + "." + idgv);

  const sheet = ss.getSheetByName("ketqua");
  const data = sheet.getDataRange().getValues();

  const results = data.slice(1).filter(row =>
    row[10].toString().trim() === key
  );

  if (results.length === 0) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "not_found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const row = results[0];

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "success",
      data: {
        exams: row[1],
        sbd: row[2],
        name: row[3],
        class: row[4],
        tongdiem: row[5],
        time: row[6]
      }
    }))
    .setMimeType(ContentService.MimeType.JSON);
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
  return id.toString().toUpperCase().trim().slice(-9);
}
function supper(text) {
  if (text === null || text === undefined) return "";
   return text.toString().replace(/'/g, "").toUpperCase().trim()
}
  

// Hàm xóa nhiều dòng //
/**
 * Xóa dữ liệu cực nhanh và GIỮ LẠI dòng tiêu đề (Header)
 */

  function deleteFastAll(text, number, name) {  
  var sheet = ss.getSheetByName(name);
 if (!sheet) return createResponse("exists", "Sheet " + name + " không tồn tại!");

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow <= 1) createResponse("exists", "Sheet " + name + " đang trống dữ liệu!");

  // 👉 chỉ lấy data (bỏ header)
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  var key = N9(text);

  var filteredData = data.filter(function(row) {
    return N9(row[number - 1]) !== key;
  });

  var deletedCount = data.length - filteredData.length;

  // 👉 clear data cũ
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // 👉 ghi lại data mới từ dòng 2
  if (filteredData.length > 0) {
    sheet.getRange(2, 1, filteredData.length, lastCol)
         .setValues(filteredData);
  }

  return deletedCount;
}

 function deleteFast(text, number, name) {  
  var sheet = ss.getSheetByName(name);
 if (!sheet) return createResponse("exists", "Sheet " + name + " không tồn tại!");

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow <= 1) createResponse("exists", "Sheet " + name + " đang trống dữ liệu!");

  // 👉 chỉ lấy data (bỏ header)
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  var key = supper(text);

 var filteredData = data.filter(function(row, index) {
  var cell = row[number - 1];
  var val = supper(cell);

  if (index < 5) { // chỉ log vài dòng đầu
    Logger.log("👉 KEY: [" + key + "]");
    Logger.log("👉 CELL: [" + val + "]");
  }

  return val !== key;
});

  var deletedCount = data.length - filteredData.length;  

  // 👉 clear data cũ
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // 👉 ghi lại data mới từ dòng 2
  if (filteredData.length > 0) {
    sheet.getRange(2, 1, filteredData.length, lastCol)
         .setValues(filteredData);
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

  if (id === idadmin && pass === passAdmin) {
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
