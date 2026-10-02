/* =========================================================================
   Backend: Code.gs
   ระบบยืมและจองห้องคอมพิวเตอร์ (Computer Room Booking System)
   ========================================================================= */

// ชื่อ Sheet ที่ใช้งาน
const SHEET_USERS = "Users";
const SHEET_BOOKINGS = "Bookings";

/**
 * ตรวจสอบและสร้าง Sheet พร้อมหัวคอลัมน์อัตโนมัติหากยังไม่มีในระบบ
 */
function getOrCreateSheet(sheetName, defaultHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (defaultHeaders && defaultHeaders.length > 0) {
      sheet.appendRow(defaultHeaders);
      sheet
        .getRange(1, 1, 1, defaultHeaders.length)
        .setFontWeight("bold")
        .setBackground("#0f172a")
        .setFontColor("#ffffff");
    }
  }
  return sheet;
}

/**
 * Render หน้าเว็บหลัก Index.html
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ระบบยืมและจองห้องคอมพิวเตอร์ | Computer Room Booking System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * ระบบสมัครสมาชิก (Register)
 */
function processRegister(username, password, fullName) {
  try {
    if (!username || !password || !fullName) {
      return { success: false, message: "กรุณากรอกข้อมูลให้ครบถ้วน" };
    }

    const sheet = getOrCreateSheet(SHEET_USERS, [
      "UserID",
      "Username",
      "Password",
      "FullName",
      "Role",
      "Status",
    ]);
    const users = getSheetDataAsObjects(SHEET_USERS);

    // ตรวจสอบชื่อผู้ใช้ซ้ำ (ไม่สนใจตัวพิมพ์เล็ก-ใหญ่)
    const existingUser = users.find(
      (u) =>
        String(u.Username).trim().toLowerCase() ===
        String(username).trim().toLowerCase(),
    );
    if (existingUser) {
      return {
        success: false,
        message: "ชื่อผู้ใช้นี้ (Username) ถูกใช้งานแล้ว กรุณาใช้ชื่ออื่น",
      };
    }

    // Gen รหัสผู้ใช้ใหม่ เช่น USR240924153000
    const userId =
      "USR" +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyMMddHHmmss",
      );

    // บันทึกลง Sheet (กำหนดสิทธิ์เริ่มต้นเป็น User และสถานะ Active)
    sheet.appendRow([
      userId,
      username.trim(),
      password.trim(),
      fullName.trim(),
      "User",
      "Active",
    ]);

    return {
      success: true,
      message: "สมัครสมาชิกสำเร็จ! ท่านสามารถเข้าสู่ระบบได้ทันที",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาดในการสมัครสมาชิก: " + error.message,
    };
  }
}

/**
 * แปลงข้อมูลใน Sheet ให้กลายเป็น Array of Objects
 */
function getSheetDataAsObjects(sheetName) {
  const sheet = getOrCreateSheet(sheetName, []);
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const headers = data.shift();
  return data.map((row) => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

/**
 * ระบบเข้าสู่ระบบ (Authentication)
 */
function processLogin(username, password) {
  try {
    const sheet = getOrCreateSheet(SHEET_USERS, [
      "UserID",
      "Username",
      "Password",
      "FullName",
      "Role",
      "Status",
    ]);

    if (sheet.getLastRow() <= 1) {
      sheet.appendRow([
        "USR001",
        "admin",
        "1234",
        "ผู้ดูแลระบบ (Admin)",
        "Admin",
        "Active",
      ]);
      sheet.appendRow([
        "USR002",
        "user",
        "1234",
        "ผู้ใช้งานทั่วไป (User)",
        "User",
        "Active",
      ]);
    }

    const users = getSheetDataAsObjects(SHEET_USERS);
    const user = users.find(
      (u) =>
        String(u.Username).trim() === String(username).trim() &&
        String(u.Password).trim() === String(password).trim(),
    );

    if (user) {
      if (user.Status !== "Active") {
        return {
          success: false,
          message: "บัญชีของคุณถูกระงับการใช้งาน โปรดติดต่อผู้ดูแลระบบ",
        };
      }
      return {
        success: true,
        user: {
          UserID: user.UserID,
          Username: user.Username,
          FullName: user.FullName,
          Role: user.Role,
        },
      };
    } else {
      return {
        success: false,
        message:
          "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง (ลองใช้ admin / 1234 หรือ user / 1234)",
      };
    }
  } catch (error) {
    return { success: false, message: "ข้อผิดพลาดระบบ: " + error.message };
  }
}

// ตัวอย่างฟังก์ชันบันทึกการยืมทันทีใน Code.gs
function handleBorrowNow(data) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Bookings");

  // สร้างฟอร์แมตเวลาปัจจุบัน (ประเทศไทย GMT+7)
  const now = new Date();
  const currentDateTime = Utilities.formatDate(
    now,
    "GMT+7",
    "yyyy-MM-dd HH:mm:ss",
  );

  const newBookingId =
    "BK" + Utilities.formatDate(now, "GMT+7", "yyyyMMddHHmmss");

  // บันทึกแถวใหม่ลง Sheet (เรียงลำดับคอลัมน์ตาม Sheet ของคุณ)
  sheet.appendRow([
    newBookingId, // BookingID
    data.userId, // UserID
    data.fullName, // FullName
    data.roomName, // RoomName
    "ยืมทันที", // ActionType
    currentDateTime, // StartDateTime -> บันทึกเวลาที่กดทำรายการจริง
    "กำลังใช้งาน", // EndDateTime -> สถานะเริ่มต้นเมื่อยืมทันที
    currentDateTime, // Timestamp
  ]);

  return { success: true, message: "บันทึกการยืมทันทีเรียบร้อยแล้ว" };
}

/**
 * ดึงประวัติการยืม-จองของผู้ใช้รายบุคคล
 */
function getUserBookings(userId) {
  try {
    const allBookings = getSheetDataAsObjects(SHEET_BOOKINGS);
    const userBookings = allBookings
      .filter((b) => String(b.UserID) === String(userId))
      .reverse();
    return { success: true, data: userBookings };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * บันทึกการยืม/จองห้องคอมพิวเตอร์
 */
function submitBooking(
  userId,
  roomName,
  actionType,
  startDateTime,
  endDateTime,
) {
  try {
    const sheet = getOrCreateSheet(SHEET_BOOKINGS, [
      "BookingID",
      "UserID",
      "RoomName",
      "ActionType",
      "StartDateTime",
      "EndDateTime",
      "Timestamp",
    ]);
    const bookingId =
      "BK" +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyyMMddHHmmss",
      );
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss",
    );

    let startStr = startDateTime;
    let endStr = endDateTime;

    if (actionType === "ใช้ห้องทันที") {
      startStr = timestamp;
      endStr = "กำลังใช้งาน";
    } else {
      startStr = startDateTime.replace("T", " ");
      endStr = endDateTime.replace("T", " ");
    }

    sheet.appendRow([
      bookingId,
      userId,
      roomName,
      actionType,
      startStr,
      endStr,
      timestamp,
    ]);

    const users = getSheetDataAsObjects(SHEET_USERS);
    const userObj = users.find((u) => String(u.UserID) === String(userId));
    const userName = userObj ? userObj.FullName : userId;

    const lineMsg =
      `🔔 มีรายการทำรายการใหม่!\n` +
      `📌 รหัสรายการ: ${bookingId}\n` +
      `👤 ผู้ทำรายการ: ${userName}\n` +
      `💻 ห้อง: ${roomName}\n` +
      `📝 ประเภท: ${actionType}\n` +
      `⏰ เริ่มต้น: ${startStr}\n` +
      `⌛ สิ้นสุด: ${endStr}`;

    try {
      sendLineMessage(lineMsg);
    } catch (lineErr) {
      Logger.log("LINE Notify Warning: " + lineErr.message);
    }

    return { success: true, message: `ทำรายการ${actionType} สำเร็จแล้ว!` };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * บันทึกการคืนห้องคอมพิวเตอร์
 */
function returnRoom(bookingId) {
  try {
    const sheet = getOrCreateSheet(SHEET_BOOKINGS, [
      "BookingID",
      "UserID",
      "RoomName",
      "ActionType",
      "StartDateTime",
      "EndDateTime",
      "Timestamp",
    ]);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const bookingIdColIndex = headers.indexOf("BookingID");
    const endDateTimeColIndex = headers.indexOf("EndDateTime");
    const roomNameColIndex = headers.indexOf("RoomName");
    const userIdColIndex = headers.indexOf("UserID");

    if (bookingIdColIndex === -1 || endDateTimeColIndex === -1) {
      return {
        success: false,
        message: "ไม่พบหัวคอลัมน์ที่จำเป็นในตาราง Bookings",
      };
    }

    let rowIndex = -1;
    let roomName = "";
    let userId = "";

    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][bookingIdColIndex]).trim() === String(bookingId).trim()
      ) {
        rowIndex = i + 1;
        userId = data[i][userIdColIndex];
        roomName = data[i][roomNameColIndex];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "ไม่พบรหัสรายการยืมนี้ในระบบ" };
    }

    const nowStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss",
    );

    // อัปเดต EndDateTime
    sheet.getRange(rowIndex, endDateTimeColIndex + 1).setValue(nowStr);

    const users = getSheetDataAsObjects(SHEET_USERS);
    const userObj = users.find((u) => String(u.UserID) === String(userId));
    const userName = userObj ? userObj.FullName : userId;

    const lineMsg =
      `🔄 มีการคืนห้องเรียบร้อยแล้ว!\n` +
      `📌 รหัสรายการ: ${bookingId}\n` +
      `👤 ผู้ทำรายการ: ${userName}\n` +
      `💻 ห้อง: ${roomName}\n` +
      `⏰ เวลาคืนห้อง: ${nowStr}`;

    try {
      sendLineMessage(lineMsg);
    } catch (lineErr) {
      Logger.log("LINE Notify Warning: " + lineErr.message);
    }

    return { success: true, message: "บันทึกการคืนห้องสำเร็จเรียบร้อยแล้ว!" };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.message };
  }
}

/**
 * ดึงข้อมูลผู้ใช้งานทั้งหมด และรายการยืม-จอง ทั้งหมดสำหรับ Admin
 */
function getAdminData() {
  try {
    const users = getSheetDataAsObjects(SHEET_USERS);
    const bookings = getSheetDataAsObjects(SHEET_BOOKINGS).reverse();

    const enrichedBookings = bookings.map((b) => {
      const u = users.find((user) => String(user.UserID) === String(b.UserID));
      return {
        ...b,
        FullName: u ? u.FullName : "ไม่ทราบชื่อ",
      };
    });

    return { success: true, users: users, bookings: enrichedBookings };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * สลับสถานะผู้ใช้งาน (Active / Inactive) แบบ Real-time
 */
function toggleUserStatus(userId, currentStatus) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
    if (!sheet) return { success: false, message: "ไม่พบชีต Users" };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf("UserID");
    const statusColIndex = headers.indexOf("Status");

    const newStatus = currentStatus === "Active" ? "Inactive" : "Active";

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIndex]) === String(userId)) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue(newStatus);
        return {
          success: true,
          newStatus: newStatus,
          message: `อัปเดตสถานะผู้ใช้เป็น ${newStatus} เรียบร้อยแล้ว`,
        };
      }
    }
    return { success: false, message: "ไม่พบผู้ใช้งานในระบบ" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * ส่งข้อความแจ้งเตือนผ่าน LINE Messaging API Push Message
 */
function sendLineMessage(message) {
  if (
    message === null ||
    message === undefined ||
    String(message).trim() === ""
  ) {
    throw new Error("ข้อความที่จะส่งไปยัง LINE ต้องไม่เป็นค่าว่าง");
  }

  const properties = PropertiesService.getScriptProperties();
  const channelAccessToken = properties.getProperty(
    "LINE_CHANNEL_ACCESS_TOKEN",
  );
  const targetId = properties.getProperty("LINE_TARGET_ID");

  if (!channelAccessToken) {
    throw new Error("ไม่พบ LINE_CHANNEL_ACCESS_TOKEN ใน Script Properties");
  }

  if (!targetId) {
    throw new Error("ไม่พบ LINE_TARGET_ID ใน Script Properties");
  }

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: targetId,
    messages: [
      {
        type: "text",
        text: String(message).trim(),
      },
    ],
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + channelAccessToken,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(
      "ส่งข้อความ LINE ไม่สำเร็จ (" + statusCode + "): " + responseText,
    );
  }

  return true;
}

function testSendLineMessage() {
  sendLineMessage("✅ ทดสอบส่งข้อความจาก Google Apps Script สำเร็จ");
}
