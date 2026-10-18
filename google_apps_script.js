/**
 * ==============================================================================
 * DỰ ÁN: Qr Grandstream by CLG - Cloud Google Sheets Webhook
 * Tác giả: CLG Team
 * Chức năng: Nhận webhook từ barKoder, kiểm tra trùng MAC, lưu vào Google Sheet
 *            và cung cấp API dữ liệu trực tiếp cho GitHub Pages Dashboard.
 * ==============================================================================
 */

// Tên trang tính lưu trữ dữ liệu thiết bị
var SHEET_NAME = "Data Thiết Bị";

/**
 * Khởi tạo cấu trúc bảng tính nếu chưa có
 */
function initSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = ["STT", "MAC Address", "Serial Number", "Password", "Thời Gian Quét"];
    sheet.appendRow(headers);
    
    // Định dạng tiêu đề
    var headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#2563eb");
    headerRange.setFontColor("#ffffff");
    headerRange.setHorizontalAlignment("center");
    
    // Căn chỉnh độ rộng cột
    sheet.setColumnWidth(1, 80);   // STT
    sheet.setColumnWidth(2, 200);  // MAC Address
    sheet.setColumnWidth(3, 200);  // Serial Number
    sheet.setColumnWidth(4, 180);  // Password
    sheet.setColumnWidth(5, 180);  // Thời Gian Quét
    sheet.setFrozenRows(1);
  }
}

/**
 * Xử lý Webhook POST từ ứng dụng quét mã (barKoder)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Khóa 15 giây để chống xung đột khi quét nhiều máy liên tục
  if (!lock.tryLock(15000)) {
    return createJsonResponse({
      status: false,
      message: "Hệ thống đang bận xử lý thiết bị khác, vui lòng thử lại sau 2 giây!"
    });
  }

  try {
    var rawText = (e && e.postData && e.postData.contents) ? e.postData.contents : "";
    var body = {};

    if (rawText) {
      try {
        body = JSON.parse(rawText);
      } catch (err) {
        body = (e && e.parameter) ? e.parameter : {};
      }
    } else if (e && e.parameter) {
      body = e.parameter;
    }

    var rawData = body.data || [];
    if (typeof rawData === "string") {
      try {
        rawData = JSON.parse(rawData);
      } catch (err) {}
    }

    // Trích xuất danh sách mã quét được
    var codesBatch = [];
    if (Array.isArray(rawData)) {
      for (var i = 0; i < rawData.length; i++) {
        var item = rawData[i];
        if (typeof item === "object" && item !== null) {
          var val = (item.value || "").toString().trim();
          var isB64 = item.base64 || false;
          if (val) {
            if (isB64) {
              try {
                var decoded = Utilities.newBlob(Utilities.base64Decode(val)).getDataAsString();
                val = decoded.trim();
              } catch (eDec) {}
            }
            if (val) {
              codesBatch.push(val);
            }
          }
        } else if (typeof item === "string" && item.trim()) {
          codesBatch.push(item.trim());
        }
      }
    }

    // Yêu cầu quét đủ 3 mã
    if (codesBatch.length < 3) {
      return createJsonResponse({
        status: false,
        message: "Đã nhận " + codesBatch.length + " mã. Cần quét đủ 3 mã (MAC, SN, Password)!"
      });
    }

    var targetCodes = codesBatch.slice(0, 3);
    var mac = "", sn = "", password = "";

    // Phân loại tự động thông minh
    for (var j = 0; j < targetCodes.length; j++) {
      var code = targetCodes[j];
      var cleaned = code.replace(/[: -]/g, "").trim();
      
      // MAC address: 12 ký tự hex
      if (/^[0-9A-Fa-f]{12}$/.test(cleaned)) {
        mac = cleaned;
      } 
      // Serial Number: toàn số hoặc đúng 10 ký tự
      else if (/^\d+$/.test(cleaned) || cleaned.length === 10) {
        sn = cleaned;
      } 
      // Password: chuỗi còn lại
      else {
        password = cleaned;
      }
    }

    // Gán các trường còn thiếu nếu quét theo định dạng đặc biệt
    var unassigned = [];
    for (var k = 0; k < targetCodes.length; k++) {
      var c = targetCodes[k];
      if (c !== mac && c !== sn && c !== password) {
        unassigned.push(c);
      }
    }
    if (!mac && unassigned.length > 0) mac = unassigned.shift();
    if (!sn && unassigned.length > 0) sn = unassigned.shift();
    if (!password && unassigned.length > 0) password = unassigned.shift();

    mac = mac.toUpperCase();
    var now = new Date();
    var timeStr = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd HH:mm:ss");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    initSheet(sheet);

    // Kiểm tra TRÙNG LẶP MAC trong Google Sheet
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var macRange = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (var r = 0; r < macRange.length; r++) {
        var existingMac = String(macRange[r][0]).replace(/[: -]/g, "").trim().toUpperCase();
        if (existingMac === mac) {
          // Lưu thông tin cảnh báo trùng vào Script Properties
          PropertiesService.getScriptProperties().setProperty(
            "LATEST_DUPLICATE", 
            JSON.stringify({ mac: mac, time: timeStr, row: r + 2 })
          );
          
          return createJsonResponse({
            status: false,
            message: "TRÙNG MAC: " + mac + " đã tồn tại trong Google Sheet (Dòng " + (r + 2) + ")!"
          });
        }
      }
    }

    // Xóa cảnh báo trùng cũ nếu quét mới thành công
    PropertiesService.getScriptProperties().deleteProperty("LATEST_DUPLICATE");

    // Thêm dòng mới vào Google Sheet
    var newSTT = lastRow; // Dòng 1 là tiêu đề
    sheet.appendRow([newSTT, mac, sn, password, timeStr]);

    // Định dạng dòng mới
    var currentRow = sheet.getLastRow();
    var newRange = sheet.getRange(currentRow, 1, 1, 5);
    newRange.setHorizontalAlignment("center");
    sheet.getRange(currentRow, 2).setFontFamily("Courier New").setFontWeight("bold").setFontColor("#1d4ed8");
    sheet.getRange(currentRow, 3).setFontFamily("Courier New");

    return createJsonResponse({
      status: true,
      message: "ĐÃ LƯU THÀNH CÔNG! (Tổng: " + newSTT + ")"
    });

  } catch (error) {
    return createJsonResponse({
      status: false,
      message: "Lỗi Google Apps Script: " + error.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Xử lý GET request: Trả về dữ liệu cho Dashboard trực tuyến trên GitHub Pages
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    initSheet(sheet);

    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";

    // Hành động xóa cảnh báo trùng lặp
    if (action === "clear_duplicates") {
      PropertiesService.getScriptProperties().deleteProperty("LATEST_DUPLICATE");
      return createJsonResponse({ status: true, message: "Đã xóa cảnh báo trùng." });
    }

    // Hành động thêm thiết bị test thử nghiệm
    if (action === "test_scan") {
      var mockCodes = {
        data: [
          { value: "000B82" + Math.floor(100000 + Math.random() * 900000).toString(16).toUpperCase() },
          { value: "246" + Math.floor(1000000 + Math.random() * 9000000) },
          { value: "TestPass@" + Math.floor(100 + Math.random() * 900) }
        ]
      };
      return doPost({ postData: { contents: JSON.stringify(mockCodes) } });
    }

    // Lấy toàn bộ dữ liệu thiết bị
    var lastRow = sheet.getLastRow();
    var items = [];
    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        if (row[0] !== "" && row[1] !== "") {
          var timeVal = row[4];
          if (timeVal instanceof Date) {
            timeVal = Utilities.formatDate(timeVal, "GMT+7", "yyyy-MM-dd HH:mm:ss");
          }
          items.push([
            row[0], 
            String(row[1]), 
            String(row[2]), 
            String(row[3]), 
            String(timeVal), 
            "Đã lưu Google Sheet"
          ]);
        }
      }
    }

    // Lấy thông tin cảnh báo trùng gần nhất
    var dupProp = PropertiesService.getScriptProperties().getProperty("LATEST_DUPLICATE");
    var duplicates = dupProp ? [JSON.parse(dupProp)] : [];

    // Đảo ngược để hiển thị thiết bị mới quét nhất lên đầu
    items.reverse();

    return createJsonResponse({
      status: true,
      total: items.length,
      items: items,
      duplicates: duplicates,
      sheet_url: ss.getUrl()
    });

  } catch (err) {
    return createJsonResponse({
      status: false,
      total: 0,
      items: [],
      error: err.toString()
    });
  }
}

/**
 * Hàm tiện ích định dạng phản hồi chuẩn JSON có hỗ trợ CORS
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
